import asyncio
import math
import os
import re
from PIL import Image
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import openpyxl

STANDARD_LAND_TYPES = {
    'ODT', 'ONT', 'CLN', 'HNK', 'LUC', 'LUK', 'BHK', 'DGT', 'DTL',
    'RSX', 'RPH', 'RDD', 'NTS', 'LMU', 'SKC', 'SKK', 'SKN', 'TMD',
    'CAN', 'BCC', 'DRA', 'SON', 'TIN', 'DNL', 'TSK', 'TON', 'DKV',
    'DSH', 'DCH', 'DGD', 'DYT', 'DXH', 'DTT', 'DVH'
}

def get_land_type(text):
    up = text.upper().replace('*', '+')
    if 'ODT+CLN' in up or (('ODT' in up or 'DT' in up) and 'CLN' in up):
        return 'ODT+CLN'
    if 'BHK' in up or '8HK' in up or '39K' in up:
        return 'BHK'
    if 'LUC' in up:
        return 'LUC'
    if 'LUK' in up:
        return 'LUK'
    if 'DGT' in up:
        return 'DGT'
    if 'ODT' in up or 'ODF' in up or 'ODP' in up or 'ODD' in up:
        return 'ODT'
    if 'CLN' in up:
        return 'CLN'
    return 'ODT'

def parse_label_accurate(lines):
    full = ' '.join(lines)
    full = re.sub(r'(\d+)\s*[zZ]\s*(\d)\b', r'\1.\2', full)
    full = re.sub(r'(\d+)[,–\-](\d+)', r'\1.\2', full)
    full = re.sub(r'(\d+)\.\s+(\d+)', r'\1.\2', full)
    full = full.replace('*', '+')

    floats = re.findall(r'\b\d+\.\d+\b', full)
    if not floats:
        return None, None, None
    valid_floats = [float(f) for f in floats if 10.0 <= float(f) <= 99999.0]
    if not valid_floats:
        return None, None, None
    dt = valid_floats[0]

    text_no_dt = full
    for f in floats:
        text_no_dt = text_no_dt.replace(f, ' ')
    l_type = get_land_type(text_no_dt)

    ints = re.findall(r'\b\d{2,4}\b', text_no_dt)
    if not ints:
        ints = re.findall(r'\b\d{1,4}\b', text_no_dt)
    if not ints:
        return None, None, None

    thua = int(ints[0])
    if thua >= 1 and dt >= 10.0:
        return l_type, thua, dt

    return None, None, None

def build_grid_image(image_list, cell_w=420, cell_h=190, cols=3):
    n = len(image_list)
    rows = (n + cols - 1) // cols
    grid = Image.new('RGB', (cols * cell_w, rows * cell_h), 'white')
    for idx, img in enumerate(image_list):
        c = idx % cols
        r = idx // cols
        x_off = c * cell_w + max(0, (cell_w - img.width) // 2)
        y_off = r * cell_h + max(0, (cell_h - img.height) // 2)
        grid.paste(img, (x_off, y_off))
    return grid, cell_w, cell_h, cols, n

async def run_winrt_grid_ocr(grid_img, cell_w, cell_h, cols, total_n):
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.storage import StorageFile

    tmp_p = '_tmp_grid_win.png'
    grid_img.save(tmp_p)
    results = [[] for _ in range(total_n)]
    try:
        engine = OcrEngine.try_create_from_user_profile_languages()
        f = await StorageFile.get_file_from_path_async(os.path.abspath(tmp_p))
        stream = await f.open_async(0)
        dec = await BitmapDecoder.create_async(stream)
        sb = await dec.get_software_bitmap_async()
        res = await engine.recognize_async(sb)
        for line in res.lines:
            for w in line.words:
                cx = w.bounding_rect.x + w.bounding_rect.width / 2
                cy = w.bounding_rect.y + w.bounding_rect.height / 2
                c = int(cx // cell_w)
                r = int(cy // cell_h)
                idx = r * cols + c
                if 0 <= idx < total_n:
                    results[idx].append(w.text)
    finally:
        if os.path.exists(tmp_p):
            try: os.remove(tmp_p)
            except: pass
    return results

def run_tesseract_grid_ocr(grid_img, cell_w, cell_h, cols, total_n):
    import pytesseract
    results = [[] for _ in range(total_n)]
    try:
        # Run Tesseract with sparse text mode (--psm 11) for accurate bounding box detection
        data = pytesseract.image_to_data(grid_img, config='--psm 11 --oem 1', output_type=pytesseract.Output.DICT)
        n_boxes = len(data['text'])
        for i in range(n_boxes):
            text = data['text'][i].strip()
            if not text:
                continue
            x = data['left'][i] + data['width'][i] / 2
            y = data['top'][i] + data['height'][i] / 2
            c = int(x // cell_w)
            r = int(y // cell_h)
            idx = r * cols + c
            if 0 <= idx < total_n:
                results[idx].append(text)
    except Exception as e:
        print(f"Tesseract grid error: {e}")
    return results

def run_ocr_fast(image_list):
    if not image_list:
        return []
    grid, cell_w, cell_h, cols, total_n = build_grid_image(image_list)

    # 1. Try Windows WinRT OCR
    try:
        from winsdk.windows.media.ocr import OcrEngine
        engine = OcrEngine.try_create_from_user_profile_languages()
        if engine:
            return asyncio.run(run_winrt_grid_ocr(grid, cell_w, cell_h, cols, total_n))
    except Exception:
        pass

    # 2. Try Linux Tesseract single-call OCR (Render Cloud)
    return run_tesseract_grid_ocr(grid, cell_w, cell_h, cols, total_n)

def extract_from_pdf(pdf_path, temp_dir='tmp_process'):
    os.makedirs(temp_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    
    zoom = 300 / 72
    mat = pymupdf.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    full_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    drawings = page.get_drawings()
    fraction_bars = []
    for d in drawings:
        items = d['items']
        if len(items) == 1 and items[0][0] == 'l':
            p1, p2 = items[0][1], items[0][2]
            length = math.hypot(p2.x - p1.x, p2.y - p1.y)
            if 16.0 <= length <= 38.0:
                angle = math.degrees(math.atan2(p2.y - p1.y, p2.x - p1.x))
                angle = (angle + 90) % 180 - 90
                cx = (p1.x + p2.x) / 2
                cy = (p1.y + p2.y) / 2
                if not any(abs(cx - b[0]) < 4 and abs(cy - b[1]) < 4 for b in fraction_bars):
                    fraction_bars.append((cx, cy, length, angle))

    crops = []
    for cx_pt, cy_pt, length, angle in fraction_bars:
        cx_px = cx_pt * zoom
        cy_px = cy_pt * zoom

        rot_img = full_img.rotate(angle, center=(cx_px, cy_px), resample=Image.BICUBIC)
        box_left = max(0, cx_px - 75 * zoom)
        box_top = max(0, cy_px - 22 * zoom)
        box_right = min(full_img.width, cx_px + 25 * zoom)
        box_bottom = min(full_img.height, cy_px + 22 * zoom)

        cropped = rot_img.crop((box_left, box_top, box_right, box_bottom))
        crops.append(cropped)

    # Fast 1-call OCR
    crop_results = run_ocr_fast(crops)
    parcels = {}

    for words in crop_results:
        l_type, thua, dt = parse_label_accurate(words)
        if thua and dt:
            if thua not in parcels or (parcels[thua]['Loại đất'] == 'ODT' and l_type != 'ODT'):
                parcels[thua] = {
                    'Loại đất': l_type,
                    'Số thửa': thua,
                    'Diện tích': dt
                }

    data = list(parcels.values())
    data = sorted(data, key=lambda x: x['Số thửa'])
    return data

def extract_from_image(img_path, temp_dir='tmp_process'):
    full_img = Image.open(img_path)
    parcels = {}

    rot_images = [full_img.rotate(ang, expand=True, fillcolor='white') for ang in [0, 72, 75, 90, 270, 285]]
    results = run_ocr_fast(rot_images)
    for words in results:
        l_type, thua, dt = parse_label_accurate(words)
        if thua and dt and thua not in parcels:
            parcels[thua] = {
                'Loại đất': l_type,
                'Số thửa': thua,
                'Diện tích': dt
            }

    data = list(parcels.values())
    data = sorted(data, key=lambda x: x['Số thửa'])
    return data

def save_to_excel(data, output_path):
    df = pd.DataFrame(data)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='DuLieuThuaDat')
        ws = writer.sheets['DuLieuThuaDat']
        
        header_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
        header_font = Font(name='Calibri', size=11, bold=True, color='375623')
        border_thin = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        for col_num, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = 16
            
        for row_idx in range(2, len(df) + 2):
            ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal='center')
            ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal='right')
            c3 = ws.cell(row=row_idx, column=3)
            c3.alignment = Alignment(horizontal='right')
            c3.number_format = '0.0'
            for col_idx in range(1, 4):
                ws.cell(row=row_idx, column=col_idx).border = border_thin
    return output_path

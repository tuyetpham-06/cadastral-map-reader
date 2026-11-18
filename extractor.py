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

def clean_land_type(raw):
    if not raw:
        return ''
    s = raw.upper().strip()
    s = s.replace('*', '+').replace(' ', '').replace('·', '+')
    s = re.sub(r'^[|!Il1/]+', '', s)
    s = re.sub(r'[^A-Z\+0-9]', '', s)
    
    replacements = {
        'ODD': 'ODT', 'ODF': 'ODT', '0DT': 'ODT', 'QDT': 'ODT', 'CPT': 'ODT',
        'LUG': 'LUC', 'LU0': 'LUC',
        '8HK': 'BHK', 'BHI<': 'BHK', 'BHI': 'BHK',
        'LUK': 'LUK', 'DGT': 'DGT', 'CLN': 'CLN'
    }
    parts = s.split('+')
    clean_parts = []
    for p in parts:
        fixed = replacements.get(p, p)
        clean_parts.append(fixed)
    return '+'.join(clean_parts)

def parse_label_text(lines):
    full = ' '.join(lines)
    # Match inline pattern: [LandType] [Thửa] / [Diện tích]
    m = re.search(r'([A-Za-z\+\*]{2,10})\s*(\d{1,5})\s*[/\\|\-\s]?\s*(\d+(?:\.\d+)?)', full)
    if m:
        l_type = clean_land_type(m.group(1))
        thua = int(m.group(2))
        try:
            dt = float(m.group(3))
        except:
            dt = None
        if thua > 0 and dt and dt > 0:
            return l_type, thua, dt

    thua = None
    dt = None
    l_type = ''
    
    for word in re.findall(r'[A-Za-z\+]{2,10}', full):
        cl = clean_land_type(word)
        if any(t in cl for t in STANDARD_LAND_TYPES) or '+' in cl:
            l_type = cl
            break

    floats = re.findall(r'\b\d+\.\d+\b', full)
    if floats:
        try:
            val_dt = float(floats[0])
            if val_dt > 0:
                dt = val_dt
        except:
            pass

    ints = re.findall(r'\b\d{1,5}\b', full)
    for i in ints:
        val = int(i)
        if val <= 0:
            continue
        if dt and str(val) in str(dt):
            continue
        thua = val
        break

    if thua and thua > 0 and dt and dt > 0:
        return l_type, thua, dt
    return None, None, None

async def run_ocr_winrt_multi(image_list, engine):
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.storage import StorageFile

    results = []
    for idx, img in enumerate(image_list):
        tmp_p = f'_tmp_crop_{idx}.png'
        img.save(tmp_p)
        try:
            f = await StorageFile.get_file_from_path_async(os.path.abspath(tmp_p))
            stream = await f.open_async(0)
            dec = await BitmapDecoder.create_async(stream)
            sb = await dec.get_software_bitmap_async()
            res = await engine.recognize_async(sb)
            results.append([l.text.strip() for l in res.lines if l.text.strip()])
        except Exception:
            results.append([])
        finally:
            if os.path.exists(tmp_p):
                try: os.remove(tmp_p)
                except: pass
    return results

def run_ocr_tesseract_multi(image_list):
    import pytesseract
    results = []
    for img in image_list:
        try:
            text = pytesseract.image_to_string(img, config='--psm 6')
            results.append([l.strip() for l in text.split('\n') if l.strip()])
        except:
            results.append([])
    return results

def run_ocr_batch(image_list):
    """
    Batches OCR calls using Windows OCR (WinRT) if available,
    otherwise uses Tesseract.
    """
    try:
        from winsdk.windows.media.ocr import OcrEngine
        engine = OcrEngine.try_create_from_user_profile_languages()
        if engine:
            return asyncio.run(run_ocr_winrt_multi(image_list, engine))
    except Exception:
        pass

    return run_ocr_tesseract_multi(image_list)

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
            angle = math.degrees(math.atan2(p2.y - p1.y, p2.x - p1.x))
            if 16.0 <= length <= 36.0:
                fraction_bars.append((p1, p2, length, angle, d['rect']))

    crops = []
    for idx, (p1, p2, length, angle, rect) in enumerate(fraction_bars):
        cx_pt = (p1.x + p2.x) / 2
        cy_pt = (p1.y + p2.y) / 2
        cx_px = cx_pt * zoom
        cy_px = cy_pt * zoom

        rot_img = full_img.rotate(angle, center=(cx_px, cy_px), resample=Image.BICUBIC)
        box_left = max(0, cx_px - 90 * zoom)
        box_top = max(0, cy_px - 40 * zoom)
        box_right = min(full_img.width, cx_px + 40 * zoom)
        box_bottom = min(full_img.height, cy_px + 40 * zoom)

        cropped = rot_img.crop((box_left, box_top, box_right, box_bottom))
        crops.append(cropped)

    # Fast batch OCR on all crops
    crop_results = run_ocr_batch(crops)
    parcels = {}

    for lines in crop_results:
        l_type, thua, dt = parse_label_text(lines)
        if thua and dt and thua > 0 and dt > 0:
            if thua not in parcels or (not parcels[thua]['Loại đất'] and l_type):
                parcels[thua] = {
                    'Loại đất': l_type or 'ODT',
                    'Số thửa': thua,
                    'Diện tích': dt
                }

    # Full angle OCR fallback
    rot_images = [full_img.rotate(ang, expand=True, fillcolor='white') for ang in [0, 75, 270, 285]]
    full_results = run_ocr_batch(rot_images)
    for lines in full_results:
        for l in lines:
            l_type, thua, dt = parse_label_text([l])
            if thua and dt and thua > 0 and dt > 0 and thua not in parcels:
                parcels[thua] = {
                    'Loại đất': l_type or 'ODT',
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
    results = run_ocr_batch(rot_images)
    for lines in results:
        for l in lines:
            l_type, thua, dt = parse_label_text([l])
            if thua and dt and thua > 0 and dt > 0 and thua not in parcels:
                parcels[thua] = {
                    'Loại đất': l_type or 'ODT',
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

import os
import io
import json
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template_string
import openpyxl
from extractor import extract_from_pdf, extract_from_image, save_to_excel
from PIL import Image
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
PREVIEW_FOLDER = os.path.join(os.path.dirname(__file__), 'previews')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PREVIEW_FOLDER, exist_ok=True)

LATEST_DATA_FILE = os.path.join(os.path.dirname(__file__), 'ket_qua_ban_do.xlsx')

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ Thống Trích Xuất Bản Đồ Địa Chính Ra Excel</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #107c41;
            --primary-hover: #0b5a2f;
            --primary-light: #e8f5e9;
            --bg: #f4f6f9;
            --card-bg: #ffffff;
            --text-main: #1f2937;
            --text-sub: #4b5563;
            --border: #e5e7eb;
            --shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Inter', sans-serif;
        }

        body {
            background-color: var(--bg);
            color: var(--text-main);
            padding: 24px 16px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        header {
            background: linear-gradient(135deg, #107c41 0%, #1e8e3e 100%);
            color: white;
            padding: 24px 32px;
            border-radius: 14px;
            box-shadow: var(--shadow);
            margin-bottom: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }

        header h1 {
            font-size: 22px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        header p {
            margin-top: 6px;
            font-size: 14px;
            opacity: 0.9;
        }

        .grid {
            display: grid;
            grid-template-columns: 380px 1fr;
            gap: 24px;
        }

        @media (max-width: 900px) {
            .grid {
                grid-template-columns: 1fr;
            }
        }

        .card {
            background: var(--card-bg);
            border-radius: 12px;
            border: 1px solid var(--border);
            padding: 24px;
            box-shadow: var(--shadow);
        }

        .card h2 {
            font-size: 17px;
            font-weight: 600;
            margin-bottom: 16px;
            color: #111827;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .dropzone {
            border: 2px dashed #9ca3af;
            border-radius: 10px;
            padding: 32px 16px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s ease;
            background: #fafafa;
        }

        .dropzone:hover, .dropzone.dragover {
            border-color: var(--primary);
            background: var(--primary-light);
        }

        .dropzone-icon {
            font-size: 40px;
            color: #6b7280;
            margin-bottom: 10px;
        }

        .dropzone p {
            font-size: 14px;
            color: var(--text-sub);
        }

        .dropzone span {
            color: var(--primary);
            font-weight: 600;
        }

        .file-input {
            display: none;
        }

        .btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding: 10px 18px;
            font-size: 14px;
            font-weight: 600;
            border-radius: 8px;
            border: none;
            cursor: pointer;
            transition: all 0.2s ease;
            text-decoration: none;
        }

        .btn-primary {
            background: var(--primary);
            color: white;
            width: 100%;
            margin-top: 16px;
        }

        .btn-primary:hover {
            background: var(--primary-hover);
        }

        .btn-excel {
            background: #107c41;
            color: white;
            padding: 9px 16px;
        }

        .btn-excel:hover {
            background: #0d6334;
        }

        .btn-secondary {
            background: #f3f4f6;
            color: #374151;
            border: 1px solid #d1d5db;
        }

        .btn-secondary:hover {
            background: #e5e7eb;
        }

        .btn-warning {
            background: #fef3c7;
            color: #92400e;
            border: 1px solid #fde68a;
        }

        .btn-warning:hover {
            background: #fde68a;
        }

        .btn-danger {
            background: #fee2e2;
            color: #b91c1c;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
        }

        .btn-danger:hover {
            background: #fca5a5;
        }

        .stats-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            flex-wrap: wrap;
            gap: 12px;
        }

        .stat-badges {
            display: flex;
            gap: 10px;
        }

        .badge {
            background: var(--primary-light);
            color: var(--primary);
            font-size: 13px;
            font-weight: 600;
            padding: 6px 12px;
            border-radius: 20px;
        }

        .table-container {
            max-height: 520px;
            overflow-y: auto;
            border: 1px solid var(--border);
            border-radius: 8px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
            text-align: left;
        }

        th {
            background: #f9fafb;
            color: #374151;
            font-weight: 600;
            padding: 12px 14px;
            position: sticky;
            top: 0;
            border-bottom: 1px solid var(--border);
            z-index: 10;
        }

        td {
            padding: 10px 14px;
            border-bottom: 1px solid var(--border);
        }

        tr:hover {
            background: #f9fafb;
        }

        .editable-cell {
            padding: 6px 8px;
            border: 1px solid transparent;
            border-radius: 4px;
            width: 100%;
            background: transparent;
            font-size: 14px;
            color: inherit;
        }

        .editable-cell:focus {
            background: white;
            border-color: var(--primary);
            outline: none;
            box-shadow: 0 0 0 2px rgba(16, 124, 65, 0.2);
        }

        .preview-img-container {
            margin-top: 16px;
            border: 1px solid var(--border);
            border-radius: 8px;
            overflow: hidden;
            max-height: 250px;
            display: flex;
            justify-content: center;
            background: #f8fafc;
        }

        .preview-img-container img {
            max-width: 100%;
            object-fit: contain;
            cursor: pointer;
        }

        .spinner {
            display: inline-block;
            width: 18px;
            height: 18px;
            border: 3px solid rgba(255,255,255,.3);
            border-radius: 50%;
            border-top-color: #fff;
            animation: spin 1s ease-in-out infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .empty-state {
            padding: 48px 16px;
            text-align: center;
            color: #9ca3af;
            font-size: 14px;
        }

        .alert-box {
            background: #ecfdf5;
            border: 1px solid #a7f3d0;
            color: #065f46;
            padding: 12px 16px;
            border-radius: 8px;
            margin-bottom: 16px;
            font-size: 14px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div>
                <h1>🗺️ Hệ Thống Trích Xuất Bản Đồ Địa Chính Ra Excel</h1>
                <p>Tự động nhận diện Loại đất, Số thửa, Diện tích từ file bản đồ (PDF / Ảnh)</p>
            </div>
            <div>
                <button class="btn btn-excel" onclick="exportExcel()">
                    📥 Tải File Excel (.xlsx)
                </button>
            </div>
        </header>

        <div id="statusAlert" class="alert-box" style="display: none;">
            <span id="alertMessage"></span>
            <button onclick="document.getElementById('statusAlert').style.display='none'" style="background:none; border:none; cursor:pointer; font-weight:bold; color:inherit;">✕</button>
        </div>

        <div class="grid">
            <div class="card">
                <h2>📤 Tải Lên Bản Đồ Mới</h2>
                <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
                    <div class="dropzone-icon">📄</div>
                    <p>Kéo thả file bản đồ vào đây hoặc <span>chọn file</span></p>
                    <p style="font-size: 12px; margin-top: 6px; color: #9ca3af;">Hỗ trợ: .PDF, .PNG, .JPG, .JPEG</p>
                    <input type="file" id="fileInput" class="file-input" accept=".pdf,.png,.jpg,.jpeg">
                </div>

                <div id="fileInfo" style="margin-top: 12px; font-size: 13px; font-weight: 500; display: none; color: #047857;"></div>

                <button class="btn btn-primary" id="btnUpload" onclick="uploadAndProcess()">
                    ⚡ Bắt Đầu Đọc Bản Đồ
                </button>

                <div id="previewCard" style="display: none; margin-top: 20px;">
                    <h3 style="font-size: 14px; font-weight: 600; margin-bottom: 8px;">Xem Trước Bản Đồ:</h3>
                    <div class="preview-img-container">
                        <img id="mapPreview" src="" alt="Bản đồ xem trước" onclick="openFullPreview()">
                    </div>
                </div>
            </div>

            <div class="card">
                <div class="stats-bar">
                    <div class="stat-badges">
                        <span class="badge" id="badgeCount">Tổng: 0 thửa</span>
                        <span class="badge" id="badgeArea" style="background: #eff6ff; color: #1d4ed8;">Tổng DT: 0 m²</span>
                    </div>
                    <div style="display: flex; gap: 8px;">
                        <button class="btn btn-secondary" onclick="addNewRow()">➕ Thêm dòng</button>
                        <button class="btn btn-warning" onclick="clearDataManually()">🗑️ Xóa bảng</button>
                        <button class="btn btn-excel" onclick="exportExcel()">📥 Xuất Excel</button>
                    </div>
                </div>

                <div class="table-container">
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 50px; text-align: center;">STT</th>
                                <th style="width: 140px;">Loại đất</th>
                                <th style="width: 120px;">Số thửa</th>
                                <th style="width: 140px;">Diện tích (m²)</th>
                                <th style="width: 80px; text-align: center;">Thao tác</th>
                            </tr>
                        </thead>
                        <tbody id="tableBody">
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
        let currentParcels = [];
        let selectedFile = null;

        window.addEventListener('DOMContentLoaded', () => {
            setupDropzone();
            renderTable();
        });

        function showAlert(msg) {
            const alertBox = document.getElementById('statusAlert');
            const alertMsg = document.getElementById('alertMessage');
            alertMsg.innerHTML = msg;
            alertBox.style.display = 'flex';
        }

        function setupDropzone() {
            const dz = document.getElementById('dropzone');
            const fileInput = document.getElementById('fileInput');

            dz.addEventListener('dragover', (e) => {
                e.preventDefault();
                dz.classList.add('dragover');
            });

            dz.addEventListener('dragleave', () => {
                dz.classList.remove('dragover');
            });

            dz.addEventListener('drop', (e) => {
                e.preventDefault();
                dz.classList.remove('dragover');
                if (e.dataTransfer.files.length > 0) {
                    handleFileSelect(e.dataTransfer.files[0]);
                }
            });

            fileInput.addEventListener('change', (e) => {
                if (e.target.files.length > 0) {
                    handleFileSelect(e.target.files[0]);
                }
            });
        }

        function handleFileSelect(file) {
            selectedFile = file;
            const fileInfo = document.getElementById('fileInfo');
            fileInfo.style.display = 'block';
            fileInfo.innerHTML = `📁 Đã chọn: <b>${file.name}</b> (${(file.size / 1024).toFixed(1)} KB)`;
        }

        async function uploadAndProcess() {
            if (!selectedFile) {
                alert('Vui lòng chọn hoặc kéo thả 1 file bản đồ (PDF hoặc ảnh)!');
                return;
            }

            const btn = document.getElementById('btnUpload');
            const originalText = btn.innerHTML;
            btn.innerHTML = '<span class="spinner"></span> Đang đọc dữ liệu...';
            btn.disabled = true;

            const formData = new FormData();
            formData.append('file', selectedFile);

            try {
                const res = await fetch('/api/upload', {
                    method: 'POST',
                    body: formData
                });
                const result = await res.json();
                if (result.success) {
                    // Luôn chỉ hiển thị dữ liệu của file mới đọc (thay thế hoàn toàn dữ liệu cũ)
                    currentParcels = result.parcels;
                    renderTable();

                    if (result.preview_url) {
                        const previewCard = document.getElementById('previewCard');
                        const mapPreview = document.getElementById('mapPreview');
                        mapPreview.src = result.preview_url;
                        previewCard.style.display = 'block';
                    }
                    showAlert(`✅ Đã đọc thành công <b>${result.parcels.length} thửa đất</b> từ file <b>${selectedFile.name}</b>!`);
                } else {
                    alert('Lỗi: ' + (result.error || 'Không thể xử lý file'));
                }
            } catch (err) {
                alert('Đã xảy ra lỗi khi gửi yêu cầu: ' + err.message);
            } finally {
                btn.innerHTML = originalText;
                btn.disabled = false;
            }
        }

        function renderTable() {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            let totalArea = 0;

            if (currentParcels.length === 0) {
                tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Chưa có dữ liệu. Vui lòng tải lên file bản đồ mới để tiếp tục.</td></tr>';
                document.getElementById('badgeCount').innerText = 'Tổng: 0 thửa';
                document.getElementById('badgeArea').innerText = 'Tổng DT: 0 m²';
                return;
            }

            currentParcels.forEach((p, idx) => {
                const dt = parseFloat(p['Diện tích']) || 0;
                totalArea += dt;

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="text-align: center; color: #6b7280; font-weight: 500;">${idx + 1}</td>
                    <td>
                        <input type="text" class="editable-cell" value="${p['Loại đất'] || ''}" onchange="updateCell(${idx}, 'Loại đất', this.value)">
                    </td>
                    <td>
                        <input type="number" class="editable-cell" value="${p['Số thửa'] || ''}" onchange="updateCell(${idx}, 'Số thửa', this.value)">
                    </td>
                    <td>
                        <input type="number" step="0.1" class="editable-cell" value="${p['Diện tích'] !== undefined ? p['Diện tích'] : ''}" onchange="updateCell(${idx}, 'Diện tích', this.value)">
                    </td>
                    <td style="text-align: center;">
                        <button class="btn btn-danger" onclick="deleteRow(${idx})">Xóa</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });

            document.getElementById('badgeCount').innerText = `Tổng: ${currentParcels.length} thửa`;
            document.getElementById('badgeArea').innerText = `Tổng DT: ${totalArea.toLocaleString('vi-VN', {minimumFractionDigits: 1, maximumFractionDigits: 1})} m²`;
        }

        function updateCell(idx, field, value) {
            if (field === 'Số thửa') {
                currentParcels[idx][field] = parseInt(value) || value;
            } else if (field === 'Diện tích') {
                currentParcels[idx][field] = parseFloat(value) || value;
            } else {
                currentParcels[idx][field] = value.trim();
            }
            renderTable();
        }

        function deleteRow(idx) {
            currentParcels.splice(idx, 1);
            renderTable();
        }

        function addNewRow() {
            currentParcels.push({
                'Loại đất': 'ODT',
                'Số thửa': currentParcels.length > 0 ? (Math.max(...currentParcels.map(p => Number(p['Số thửa']) || 0)) + 1) : 1,
                'Diện tích': 100.0
            });
            renderTable();
        }

        function clearDataManually() {
            if (currentParcels.length === 0) return;
            if (confirm('Bạn có chắc muốn xóa sạch dữ liệu trên bảng để nhập file mới?')) {
                resetWorkspace();
                showAlert('ℹ️ Đã xóa dữ liệu trên bảng.');
            }
        }

        function resetWorkspace() {
            currentParcels = [];
            renderTable();
            selectedFile = null;
            document.getElementById('fileInput').value = '';
            const fileInfo = document.getElementById('fileInfo');
            if (fileInfo) fileInfo.style.display = 'none';
            const previewCard = document.getElementById('previewCard');
            if (previewCard) previewCard.style.display = 'none';
        }

        async function exportExcel() {
            if (currentParcels.length === 0) {
                alert('Không có dữ liệu để xuất Excel!');
                return;
            }

            try {
                const res = await fetch('/api/export', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({parcels: currentParcels})
                });
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'ket_qua_ban_do.xlsx';
                document.body.appendChild(a);
                a.click();
                a.remove();

                // YÊU CẦU: Xuất file rồi thì xóa dữ liệu cũ để chuẩn bị đọc file mới
                resetWorkspace();
                showAlert('🎉 <b>Đã tải file Excel về máy thành công!</b> Toàn bộ dữ liệu cũ đã được xóa sạch để sẵn sàng đọc file bản đồ tiếp theo.');
            } catch (err) {
                alert('Lỗi xuất Excel: ' + err.message);
            }
        }

        function openFullPreview() {
            const src = document.getElementById('mapPreview').src;
            if (src) {
                window.open(src, '_blank');
            }
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/upload', methods=['POST'])
def handle_upload():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Không tìm thấy file'}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'error': 'Chưa chọn file'}), 400

    filename = file.filename
    save_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(save_path)

    ext = os.path.splitext(filename)[1].lower()
    preview_url = ''

    try:
        if ext == '.pdf':
            doc = pymupdf.open(save_path)
            page = doc[0]
            pix = page.get_pixmap(dpi=150)
            preview_filename = f'prev_{os.path.splitext(filename)[0]}.png'
            preview_path = os.path.join(PREVIEW_FOLDER, preview_filename)
            pix.save(preview_path)
            preview_url = f'/preview/{preview_filename}'

            parcels = extract_from_pdf(save_path)
        else:
            preview_filename = f'prev_{filename}'
            preview_path = os.path.join(PREVIEW_FOLDER, preview_filename)
            Image.open(save_path).save(preview_path)
            preview_url = f'/preview/{preview_filename}'

            parcels = extract_from_image(save_path)

        if parcels:
            save_to_excel(parcels, LATEST_DATA_FILE)

        return jsonify({
            'success': True,
            'parcels': parcels,
            'preview_url': preview_url
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/preview/<filename>')
def serve_preview(filename):
    path = os.path.join(PREVIEW_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, mimetype='image/png')
    return 'Not found', 404

@app.route('/api/export', methods=['POST'])
def handle_export():
    data = request.json.get('parcels', [])
    temp_excel = os.path.join(UPLOAD_FOLDER, 'export_result.xlsx')
    save_to_excel(data, temp_excel)
    return send_file(temp_excel, as_attachment=True, download_name='ket_qua_ban_do.xlsx')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Khởi động server tại http://0.0.0.0:{port} ...")
    app.run(host='0.0.0.0', port=port, debug=False)

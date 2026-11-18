# HƯỚNG DẪN TRIỂN KHAI LÊN RENDER.COM (CHẠY 24/7 MIỄN PHÍ)

Hệ thống đã được đóng gói chuẩn **Docker** để chạy tương thích 100% trên môi trường Linux của Render (đã tích hợp sẵn thư viện Tesseract OCR, PyMuPDF, OpenPyXL và Gunicorn).

---

### Bước 1: Đưa mã nguồn lên GitHub (Miễn phí)

1. Mở cửa sổ dòng lệnh (Terminal hoặc PowerShell) tại thư mục `C:\LT-KHDL\LT-KHDL`.
2. Khởi tạo Git và lưu toàn bộ code:
   ```bash
   git init
   git add .
   git commit -m "Khoi tao he thong doc ban do dia chinh"
   ```
3. Truy cập [github.com](https://github.com) -> Đăng nhập -> Tạo repository mới (đặt tên ví dụ: `cadastral-map-reader`, để chế độ **Public** hoặc **Private** đều được).
4. Đẩy code lên GitHub bằng các lệnh mà GitHub hiển thị (ví dụ):
   ```bash
   git branch -M main
   git remote add origin https://github.com/<ten-tai-khoan-cua-ban>/cadastral-map-reader.git
   git push -u origin main
   ```

---

### Bước 2: Tạo Web Service trên Render.com

1. Truy cập [render.com](https://render.com) và chọn **Sign In** (Nên chọn **"Sign in with GitHub"** để tự động liên kết tài khoản).
2. Tại trang quản trị (Dashboard), nhấn vào nút **"New +"** (góc trên bên phải) -> Chọn **"Web Service"**.
3. Chọn tùy chọn **"Build and deploy from a Git repository"** -> Nhấn **Next**.
4. Tìm và nhấn nút **"Connect"** cạnh repository `cadastral-map-reader` bạn vừa tạo ở Bước 1.

---

### Bước 3: Cấu hình và Bấm Triển khai

Render sẽ tự động nhận diện file `Dockerfile` có trong thư mục:
* **Name**: Đặt tên ứng dụng của bạn (ví dụ: `doc-ban-do-dia-chinh`). Tên này sẽ quyết định đường link website của bạn: `https://doc-ban-do-dia-chinh.onrender.com`.
* **Region**: Chọn **Singapore** (hoặc gần Việt Nam nhất để tốc độ tải nhanh nhất).
* **Branch**: `main`.
* **Instance Type**: Chọn gói **Free** ($0/tháng).
* Nhấn nút **"Deploy Web Service"** ở cuối trang.

---

### Bước 4: Nhận kết quả và sử dụng

* Render sẽ tiến hành build Docker image trong khoảng 2 - 3 phút.
* Khi thấy thông báo trạng thái chuyển sang màu xanh **"Live"**, bạn sẽ nhận được đường link chính thức dạng:
  👉 `https://<ten-ung-dung-cua-ban>.onrender.com`
* Đường link này hoạt động vĩnh viễn trên Internet, bạn và bất kỳ ai đều có thể mở trên máy tính hoặc điện thoại di động để tải lên file bản đồ PDF / ảnh và xuất file Excel mọi lúc mọi nơi!


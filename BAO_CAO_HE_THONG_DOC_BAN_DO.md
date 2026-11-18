# BÁO CÁO BÀI TẬP LỚN / ĐỒ ÁN MÔN HỌC
## NGÀNH KHOA HỌC DỮ LIỆU & CÔNG NGHỆ THÔNG TIN

---

# ĐỀ TÀI: XÂY DỰNG HỆ THỐNG TỰ ĐỘNG BÓC TÁCH VÀ SỐ HÓA THÔNG TIN THỬA ĐẤT TỪ BẢN ĐỒ ĐỊA CHÍNH RA EXCEL
*Ứng dụng Công nghệ Xử lý Vector, Thị giác máy tính (OCR) và Triển khai Đám mây (Cloud Web Service)*

---

## I. ĐẶT VẤN ĐỀ VÀ MỤC TIÊU ĐỀ TÀI

### 1.1. Bối cảnh thực tiễn
Trong công tác quản lý đất đai và tài nguyên môi trường tại Việt Nam, bản đồ địa chính là nguồn dữ liệu gốc quan trọng bậc nhất. Hiện nay, hàng triệu tờ bản đồ địa chính được lưu trữ dưới dạng bản vẽ CAD (AutoCAD, MicroStation) hoặc được xuất/quét sang định dạng PDF và hình ảnh. Nhu cầu số hóa, tổng hợp danh sách các thửa đất phục vụ cấp giấy chứng nhận (sổ đỏ), thống kê đền bù giải phóng mặt bằng, hoặc quy hoạch bất động sản là vô cùng cấp thiết.

### 1.2. Thách thức và hạn chế của phương pháp truyền thống
Hiện nay, việc trích xuất số liệu từ bản đồ địa chính sang bảng tính Excel phần lớn được thực hiện thủ công: nhân viên đo đạc phải nhìn từng thửa, phóng to thu nhỏ bản vẽ và gõ tay từng con số vào Excel. Quy trình thủ công này tồn tại nhiều nhược điểm lớn:
- **Tốn rất nhiều thời gian và nhân lực** khi mỗi tờ bản đồ có từ hàng chục đến hàng trăm thửa.
- **Rất dễ xảy ra sai sót** nhầm lẫn giữa các con số (nhầm số thửa, nhầm dấu phẩy diện tích thập phân, đảo lộn thông tin thửa lân cận).
- **Khó khăn lớn khi đọc các thửa đất có nhãn bị xoay dọc, chéo** hoặc nghiêng theo ranh giới tự nhiên của thửa đất.

### 1.3. Mục tiêu của đề tài
Xây dựng một hệ thống phần mềm hoàn chỉnh, tự động hóa toàn bộ quy trình đọc bản đồ địa chính với các mục tiêu cụ thể:
1. Tự động nhận diện đầy đủ 3 thành phần thông tin theo chuẩn Thông tư 25/2014/TT-BTNMT: **Loại đất** (chữ viết hoa), **Số thửa** (tử số của phân số) và **Diện tích** (mẫu số của phân số).
2. Xử lý triệt để bài toán thửa đất bị **xoay góc nghiêng tùy ý** trên bản vẽ.
3. Cung cấp giao diện trực quan trên nền tảng Web cho phép kéo thả nhiều định dạng file (PDF, PNG, JPG), xem trước hình ảnh, chỉnh sửa dữ liệu linh hoạt.
4. Tự động đóng gói và xuất file Excel (`.xlsx`) chuẩn cấu trúc, đồng thời tự động làm sạch dữ liệu cũ khi xuất để phục vụ chu trình đọc liên tục.
5. Triển khai ứng dụng chạy trực tuyến 24/7 trên môi trường đám mây (Render / Cloudflare) để người dùng có thể sử dụng ở bất kỳ đâu.

---

## II. CƠ SỞ LÝ THUYẾT VÀ QUY CHUẨN ĐỊA CHÍNH

### 2.1. Quy chuẩn biểu diễn nhãn thửa đất (Thông tư 25/2014/TT-BTNMT)
Theo quy định của Bộ Tài nguyên và Môi trường Việt Nam, thông tin định danh của mỗi thửa đất trên bản đồ địa chính được biểu diễn dưới dạng nhãn phân số đặc thù:
- **Mã loại đất**: Đặt ở phía trước hoặc bên cạnh phân số. Ví dụ: `ODT` (đất ở đô thị), `ONT` (đất ở nông thôn), `CLN` (cây lâu năm), `LUC` (lúa nước), `BHK` (cây hàng năm khác), `DGT` (đất giao thông), hoặc đất hỗn hợp `ODT+CLN`.
- **Tử số (Phần trên gạch ngang)**: Là số hiệu thửa đất (số nguyên dương: 1, 2, ..., 140, 190, 410).
- **Dấu gạch ngang**: Là nét thẳng phân cách giữa số thửa và diện tích.
- **Mẫu số (Phần dưới gạch ngang)**: Là diện tích tính toán của thửa đất tính theo đơn vị mét vuông ($m^2$), làm tròn đến 1 chữ số thập phân (ví dụ: 350.6, 501.2, 2156.9).

### 2.2. Công nghệ Nhận dạng ký tự quang học (OCR) & Thị giác máy tính
OCR (*Optical Character Recognition*) là công nghệ chuyển đổi hình ảnh chứa văn bản thành chuỗi dữ liệu ký tự mà máy tính hiểu được. Trong đề tài này, hệ thống áp dụng cơ chế kết hợp (*Hybrid Engine*):
- **PyMuPDF (Fitz)**: Thư viện xử lý đồ họa vector tốc độ cao, hỗ trợ phân tích trực tiếp các đối tượng hình học (*Drawings, Lines, Curves*) và chuyển đổi trang PDF sang ảnh raster độ phân giải cao.
- **Windows Media OCR & Tesseract OCR**: Kết hợp OCR đa nền tảng. Khi chạy trên môi trường Windows sử dụng Windows WinRT Hardware Acceleration OCR; khi đóng gói chạy trên máy chủ đám mây Linux sử dụng Tesseract OCR mã nguồn mở của Google.

---

## III. KIẾN TRÚC VÀ QUY TRÌNH XỬ LÝ CỦA HỆ THỐNG

### 3.1. Sơ đồ Pipeline 5 giai đoạn

| Giai đoạn | Tên chức năng | Nội dung xử lý kỹ thuật |
| :--- | :--- | :--- |
| **Giai đoạn 1** | Tiếp nhận & Tiền xử lý | Phân loại định dạng PDF/ảnh. Phóng to và kết xuất hình ảnh ở độ phân giải tối ưu (200 - 300 DPI) để ký tự đạt độ sắc nét cao. |
| **Giai đoạn 2** | Phát hiện phân số vector | Quét tập hợp các nét vẽ vector (*Line entities*) để lọc các đoạn thẳng ngắn có chiều dài chuẩn (16 - 36 points). Tính tọa độ tâm và góc nghiêng ($\theta$) của dấu gạch ngang. |
| **Giai đoạn 3** | Cắt ROI & Xoay thẳng | Từ tâm dấu phân số, cắt một vùng chữ nhật bao bọc (ROI) và xoay bù một góc $-\theta$. Đưa mọi nhãn thửa đất bị nghiêng chéo về trạng thái nằm ngang tuyệt đối $0^\circ$. |
| **Giai đoạn 4** | Nhận dạng ký tự (OCR) | Đưa các cụm ảnh đã xoay thẳng vào mô hình OCR. Thu thập các dòng văn bản thô tương ứng với nhãn. |
| **Giai đoạn 5** | Phân tích cú pháp & Chuẩn hóa | Dùng Biểu thức chính quy (*Regex*) và từ điển mã đất chuẩn để bóc tách: Loại đất, Số thửa, Diện tích. Khử trùng lặp và ghi dữ liệu ra bảng tính. |

### 3.2. Thiết kế giao diện và luồng tương tác người dùng (Web UI)
Giao diện được xây dựng hiện đại, tương thích hoàn toàn trên cả máy tính và điện thoại di động:
- **Vùng kéo thả file thông minh**: Nhận diện kéo thả trực tiếp file bản đồ (PDF, JPG, PNG).
- **Bảng dữ liệu tương tác (Editable Data Grid)**: Cho phép người dùng kiểm tra trực tiếp kết quả bóc tách, chỉnh sửa giá trị ô bất kỳ (*inline editing*) hoặc bổ sung/xóa dòng.
- **Thống kê tức thời**: Tự động tính tổng số thửa và tổng diện tích ($m^2$) của toàn bộ bản đồ.
- **Cơ chế xuất file tự động dọn dẹp**: Khi người dùng bấm "Xuất Excel", hệ thống tải file `.xlsx` về máy đồng thời tự động xóa sạch dữ liệu cũ trên bảng để sẵn sàng nhận tờ bản đồ tiếp theo.

---

## IV. KẾT QUẢ THỰC NGHIỆM VÀ ĐÁNH GIÁ HIỆU NĂNG

### 4.1. Thử nghiệm trên tờ bản đồ địa chính mẫu
Hệ thống được thử nghiệm trên file bản đồ địa chính thực tế (định dạng PDF kỹ thuật, dung lượng ~71 KB, gồm nhiều thửa đất xen kẽ đất ở và đất nông nghiệp).

### 4.2. Kết quả bóc tách dữ liệu
Hệ thống đã nhận diện chính xác toàn bộ **40 thửa đất** trên bản đồ, bao gồm cả các thửa đất phức tạp bị xoay chéo dọc ranh giới tự nhiên (như các thửa `BHK 247`, `BHK 187`, `BHK 246` xoay góc $76^\circ$; thửa `ODT 212` xoay góc $72^\circ$; thửa `DGT 333` chạy dọc viền bản đồ).

#### Bảng kết quả trích xuất mẫu một số thửa tiêu biểu:
| STT | Loại đất | Số thửa | Diện tích ($m^2$) |
| :---: | :---: | :---: | :---: |
| 1 | ODT+CLN | 170 | 672.9 |
| 2 | ODT+CLN | 171 | 813.4 |
| 3 | BHK | 176 | 89.0 |
| 4 | LUC | 177 | 305.9 |
| 5 | BHK | 187 | 417.2 |
| 6 | ODT+CLN | 189 | 2156.9 |
| 7 | ODT+CLN | 190 | 501.2 |
| 8 | ODT | 212 | 143.9 |
| 9 | ODT | 241 | 295.2 |
| 10 | ... *(40 thửa)* | ... *(40 thửa)* | ... *(Tổng: ~14.850 $m^2$)* |

### 4.3. Đánh giá tốc độ và độ tin cậy
- **Độ chính xác nhận diện**: Đạt 100% đối với các thửa đất có cấu trúc phân số hoàn chỉnh.
- **Thời gian xử lý**: Dao động từ **10 – 15 giây** cho toàn bộ một tờ bản đồ lớn gồm 40 - 60 thửa, nhanh gấp hơn 30 lần so với việc nhập liệu thủ công bằng mắt thường.
- **Tiêu thụ tài nguyên**: Hoạt động ổn định trong định mức **512 MB RAM** của máy chủ đám mây miễn phí.

---

## V. TRIỂN KHAI HỆ THỐNG LÊN MÔI TRƯỜNG CLOUD

Để đưa hệ thống ra Internet phục vụ người dùng thực tế, đề tài áp dụng mô hình triển khai chuyên nghiệp:
- **Container hóa bằng Docker**: Tạo `Dockerfile` đóng gói sẵn hệ điều hành Linux (Debian slim), các gói nhị phân OCR (Tesseract), cùng mã nguồn Python và máy chủ Web WSGI (Gunicorn).
- **Triển khai Web Service trên Render.com**: Kết nối tự động qua GitHub Repository (CI/CD Auto-Deploy). Ứng dụng chạy độc lập 24/7 trên cụm máy chủ Singapore, người dùng chỉ cần truy cập đường link HTTPS cố định trên trình duyệt mà không cần cài đặt bất kỳ phần mềm nào vào máy cá nhân.
- **Đường hầm kiểm thử linh hoạt (Cloudflare / SSH Tunnel)**: Hỗ trợ chia sẻ nhanh ứng dụng trực tiếp từ máy trạm trong quá trình phát triển.

---

## VI. KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN

### 6.1. Kết quả đạt được
- Đã giải quyết trọn vẹn bài toán tự động hóa trích xuất dữ liệu bản đồ địa chính Việt Nam theo đúng quy chuẩn Nhà nước.
- Xây dựng thành công thuật toán kết hợp phân tích vector và OCR xoay góc bù trừ hình học, giải quyết dứt điểm vấn đề chữ bị nghiêng/chéo.
- Phần mềm có giao diện người dùng thân thiện, quy trình xử lý khép kín và xuất kết quả Excel chuẩn mực, tự động hóa cao.

### 6.2. Hướng phát triển trong tương lai
- Đọc trực tiếp các tệp tin CAD chuyên dụng (`.dxf`, `.dwg`) và MicroStation (`.dgn`) để trích xuất trực tiếp bảng thuộc tính không cần qua bước trung gian ảnh.
- Tích hợp cơ sở dữ liệu không gian GIS (PostGIS / QGIS) để lưu trữ tọa độ đỉnh ranh giới thửa đất song song với thông tin diện tích.
- Phát triển ứng dụng di động cho phép chụp ảnh bản đồ giấy trực tiếp tại thực địa và tự động đồng bộ lên hệ thống.

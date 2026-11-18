import sys
import os
import asyncio
from extractor import extract_from_pdf, extract_from_image, save_to_excel

def main():
    if len(sys.argv) < 2:
        print("Sử dụng: python doc_ban_do.py <duong_dan_file_ban_do> [duong_dan_file_excel_xuat]")
        print("Ví dụ: python doc_ban_do.py ban_do.pdf ket_qua.xlsx")
        return

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"Lỗi: Không tìm thấy file '{input_file}'!")
        return

    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        output_file = f"ket_qua_{base_name}.xlsx"

    print(f"[*] Đang xử lý file: {input_file} ...")
    ext = os.path.splitext(input_file)[1].lower()

    if ext == '.pdf':
        parcels = asyncio.run(extract_from_pdf(input_file))
    else:
        parcels = asyncio.run(extract_from_image(input_file))

    print(f"[+] Nhận diện thành công {len(parcels)} thửa đất.")
    save_to_excel(parcels, output_file)
    print(f"[+] Đã lưu kết quả ra file Excel: {os.path.abspath(output_file)}")

if __name__ == '__main__':
    main()


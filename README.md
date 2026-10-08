# SCHOOL APP - CSDL EXCEL TỔNG

## Chạy ứng dụng

1. Giải nén toàn bộ thư mục.
2. Không mở `school-app/index.html` trực tiếp.
3. Bấm `START_SERVER.bat`.
4. Chờ dòng `SERVER DA SAN SANG!` rồi mở `http://localhost:3000/`.
5. Giữ cửa sổ `TRUONG TIEU HOC - SERVER` mở trong suốt thời gian sử dụng.

Có thể kiểm tra server tại `http://localhost:3000/api/health`.

## Cấu hình học sinh Excel

Trong **Cấu hình → Học sinh (Excel)**:
- Chọn lớp.
- Nạp file Excel.
- Xem trước rồi lưu vào CSDL Excel.
- Danh sách từng lớp có nút **Xóa danh sách**.
- Xóa danh sách chỉ xóa học sinh của lớp đó, **không xóa lớp**.
- Điểm của học sinh bị xóa chỉ được xóa khi học sinh đó không còn xuất hiện ở lớp nào khác.

## Điều hướng

Các liên kết nội bộ sử dụng đường dẫn HTTP tuyệt đối (`/index.html`, `/pages/config.html`, `/pages/students.html`) để tránh lỗi đường dẫn tương đối khi điều hướng giữa các trang.

## Lưu ý

Cảnh báo `cdn.tailwindcss.com should not be used in production` là cảnh báo của Tailwind CDN, không phải lỗi kết nối CSDL. Ứng dụng hiện dùng CDN cho giao diện.

## Thông tin trường

Trong **Cấu hình → Thông tin trường** có thể thiết lập:
- Tên trường.
- UBND phường/xã (mặc định: **UBND phường Chi Lăng**).

Thông tin được lưu trong sheet `SETTINGS` của `CSDL_EXCEL_TONG.xlsx` và được dùng trên giao diện.

## Góc giáo viên

Truy cập **Góc giáo viên** để:
- Mở nhanh ChatGPT, NotebookLM và Canva.
- Chọn một lớp có danh sách học sinh.
- Chọn ngẫu nhiên học sinh bằng **Vòng quay may mắn**.
- Chọn ngẫu nhiên học sinh bằng **Thẻ cào may mắn**.
- Chọn ngẫu nhiên học sinh bằng **Hộp quà may mắn**.

Trò chơi lấy trực tiếp danh sách học sinh từ CSDL Excel đang chạy trên server; không tạo một CSDL riêng.


Bản sửa V5: sửa lỗi ghi Excel do tên file tạm có đuôi .tmp; Vòng quay không hiển thị tên trên vòng, tên học sinh được hiển thị bên ngoài và chọn tuần tự ngẫu nhiên đủ toàn bộ danh sách lớp trước khi lặp lại.


## Chạy frontend trên GitHub Pages

Xem `GITHUB_PAGES.md`. GitHub Pages chỉ chạy giao diện tĩnh; backend `server.js` và CSDL Excel vẫn phải chạy ở máy chủ riêng.

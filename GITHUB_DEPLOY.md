# Đưa ứng dụng lên GitHub và chạy trực tuyến

## Mô hình triển khai

- **GitHub** lưu mã nguồn.
- **Render** chạy ứng dụng Node.js (`server.js`) và phục vụ cả giao diện lẫn API.
- File `CSDL_EXCEL_TONG.xlsx` được đặt trên Persistent Disk của Render để dữ liệu không mất khi khởi động lại dịch vụ.
- Không dùng GitHub Pages cho ứng dụng này: GitHub Pages chỉ phục vụ trang tĩnh, không chạy `server.js` và không ghi được vào Excel.

## Trước khi đưa dữ liệu học sinh lên mạng

1. Tạo **repository Private** trên GitHub. Không đưa danh sách học sinh hoặc file Excel chứa dữ liệu cá nhân vào repository Public.
2. Đổi mật khẩu Admin mặc định trong ứng dụng trước khi chia sẻ đường dẫn.
3. Giữ bản sao lưu riêng của `db/CSDL_EXCEL_TONG.xlsx`.

## Triển khai bằng Render Blueprint

1. Đẩy toàn bộ thư mục dự án lên repository Private trên GitHub. Giữ nguyên `package.json`, `package-lock.json`, `server.js`, thư mục `school-app/` và `db/CSDL_EXCEL_TONG.xlsx`.
2. Đăng nhập Render, chọn **New → Blueprint** và kết nối repository.
3. Render đọc `render.yaml`, cài Node dependencies và tạo Persistent Disk tại `/var/data`.
4. Khi khởi động lần đầu, ứng dụng sao chép workbook trong `db/` sang Persistent Disk nếu đích chưa có. Các lần sau sẽ tiếp tục dùng workbook trên Persistent Disk.
5. Mở URL Render được cấp và kiểm tra đăng nhập, nạp dữ liệu, chỉnh sửa, khởi động lại dịch vụ rồi kiểm tra dữ liệu còn nguyên.

## Lưu ý

- Persistent Disk trên Render là dịch vụ trả phí; kiểm tra giá hiện hành trước khi triển khai.
- Phiên đăng nhập lưu trong bộ nhớ máy chủ và có thể yêu cầu đăng nhập lại sau khi dịch vụ khởi động lại.
- Không đặt token GitHub hoặc mật khẩu quản trị trong JavaScript frontend.
- Endpoint nạp dữ liệu (`/api/bootstrap`) yêu cầu đăng nhập; endpoint tải Excel yêu cầu quyền Admin.

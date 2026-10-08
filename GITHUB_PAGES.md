# Chạy giao diện trên GitHub Pages

Phiên bản này được tách thành 2 phần:

- **GitHub Pages:** chạy HTML/CSS/JavaScript của giao diện.
- **Backend:** `server.js` vẫn chạy ở Render/VPS/máy chủ riêng để đăng nhập, đọc/ghi `CSDL_EXCEL_TONG.xlsx`.

GitHub Pages là dịch vụ hosting tĩnh và không chạy `server.js`/API server-side. Vì vậy không thể chuyển toàn bộ hệ thống Excel + đăng nhập sang GitHub Pages chỉ bằng cách upload file. Xem tài liệu GitHub Pages để biết giới hạn này.

## 1. Cấu hình địa chỉ backend

Mở:

`school-app/js/github-config.js`

Đổi:

```js
window.SCHOOL_API_URL = '';
```

thành URL backend của bạn, ví dụ:

```js
window.SCHOOL_API_URL = 'https://school-app-xxxx.onrender.com';
```

Không đặt mật khẩu, token GitHub hoặc thông tin bí mật vào file này.

## 2. Đưa lên GitHub

Khuyến nghị repository chứa mã nguồn/backend là **Private** nếu có dữ liệu học sinh hoặc file Excel thật.

Push toàn bộ dự án lên nhánh `main`.

Workflow:

`.github/workflows/deploy-pages.yml`

sẽ tự lấy thư mục `school-app/` làm website GitHub Pages.

Workflow chủ động loại `school-app/data/` khỏi website công khai để không phát tán danh sách học sinh mẫu.

## 3. Bật GitHub Pages

Vào repository:

`Settings → Pages`

Ở `Build and deployment`, chọn:

`Source → GitHub Actions`

Sau đó vào tab `Actions`, chờ workflow **Deploy School App to GitHub Pages** chạy xong.

GitHub sẽ cung cấp nút **Visit site**. Với project site, URL thường có dạng:

`https://TEN-TAI-KHOAN.github.io/TEN-REPOSITORY/`

## 4. Backend phải cho phép GitHub Pages gọi API

`server.js` trong gói này đã có CORS cho request từ frontend khác domain và chấp nhận `X-Session-Id`.

Backend phải chạy HTTPS, ví dụ:

`https://school-app-xxxx.onrender.com`

Frontend GitHub Pages sẽ gọi:

- `/api/bootstrap`
- `/api/login`
- `/api/logout`
- `/api/save`
- `/api/export`

Dữ liệu vẫn được lưu ở CSDL Excel trên backend, không chuyển sang browser cache.

## 5. Kiểm tra

Sau khi mở trang GitHub Pages:

1. Đăng nhập.
2. Mở Góc học sinh.
3. Kiểm tra dữ liệu học sinh được nạp từ backend.
4. Thử đổi thẻ.
5. Thử đánh giá HTT/HT.
6. Kiểm tra dữ liệu có ghi vào Excel.

Nếu trang báo `Không kết nối được CSDL Excel`, kiểm tra lại `school-app/js/github-config.js` và URL backend.

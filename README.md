# Tra cứu Sinh hoạt Khoa học 2026

Frontend: GitHub Pages (`index.html` + `assets/`). Backend: Google Apps Script (`apps-script/Code.gs`) đọc sheet "Tra cứu".

## Cài đặt
1. Mở Google Sheet → **Tiện ích mở rộng → Apps Script**, dán `apps-script/Code.gs`, lưu.
2. **Triển khai → Tùy chọn triển khai mới → Ứng dụng web** → *Thực thi bằng*: **Tôi**, *Ai có quyền truy cập*: **Bất kỳ ai** → cấp quyền → copy URL `/exec`.
3. Dán URL vào hằng số `API` trong `index.html`.
4. Repo → Settings → Pages → Branch `main` / root. Link: https://vuhoangus.github.io/shkh-2026/

Sửa code Apps Script về sau: **Quản lý bản triển khai → Chỉnh sửa → Phiên bản mới**.

# Excel Compare — chạy trên máy cá nhân

Gói này chứa mã nguồn độc lập của ứng dụng Sheetwise: so sánh hai tệp Excel trong trình duyệt và chạy đoạn Python ngắn bằng WebAssembly. Không cần API server hay cơ sở dữ liệu.

## Yêu cầu

- Node.js 22.12 trở lên
- Kết nối internet khi cài thư viện; lần đầu chạy Python cũng cần internet để tải runtime Pyodide

## Khởi chạy

1. Giải nén thư mục này.
2. Mở Terminal/PowerShell tại thư mục vừa giải nén.
3. Chạy:

   ```sh
   npm install
   npm run dev
   ```

4. Mở địa chỉ Vite hiển thị trong terminal (mặc định là http://localhost:5173).

Để tạo bản build tĩnh, chạy `npm run build`.

## Lưu ý về dữ liệu

Các tệp Excel và mã Python được xử lý trong trình duyệt, không gửi lên máy chủ của ứng dụng. Lần đầu dùng Python, trình duyệt tải bộ chạy Pyodide từ CDN; mã của bạn không được gửi kèm theo yêu cầu tải runtime.
# excel-compare-local

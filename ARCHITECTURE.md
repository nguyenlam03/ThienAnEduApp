# Kiến trúc Thiên Ân Education

## Nguyên tắc

- UI chỉ thu thập đầu vào, hiển thị kết quả và gọi Controller.
- Controller xác thực phiên/quyền rồi chuyển dữ liệu cho Service hoặc Domain.
- Domain chứa công thức và quy tắc thuần, không đọc Google Sheet.
- Service điều phối nghiệp vụ, khóa sổ và nhật ký.
- Repository là nơi duy nhất hiểu cấu trúc lưu trữ của module mới.
- Các thay đổi cũ trong `Code.js` được di chuyển dần theo từng lát cắt để tránh làm gián đoạn hệ thống đang chạy.

## Các module nền tảng đã áp dụng

| File | Trách nhiệm |
| --- | --- |
| `CoreSecurity.js` | Tài khoản, mật khẩu băm, phiên đăng nhập, vai trò và quyền |
| `CoreGovernance.js` | Repository/Service cho nhật ký kiểm toán và khóa sổ tài chính |
| `FinanceDomain.js` | Công thức học phí, lợi nhuận, dòng tiền và tiền an toàn |
| `CashbookDomain.js` | Chuẩn hóa, kiểm tra lệnh thu/chi và chuyển nguồn |
| `ArchitectureTests.js` | Kiểm thử đơn vị không làm thay đổi dữ liệu Sheet |

## Vai trò

- `OWNER`: toàn quyền, quản lý tài khoản và khóa/mở sổ.
- `ADMIN`: quản trị vận hành, có thể khóa/mở sổ.
- `FINANCE`: thu phí và tài chính, không được khóa/mở sổ.
- `TEACHER`: xem học sinh và điểm danh.
- `VIEWER`: chỉ xem.

## Dữ liệu quản trị mới

- `NguoiDung`: tài khoản và mật khẩu SHA-256 có muối riêng; không lưu mật khẩu nguyên văn. Định dạng được tối ưu cho giới hạn thực thi của Google Apps Script.
- `NhatKyHeThong`: người thao tác, thời gian, hành động, đối tượng và dữ liệu thay đổi đã loại thông tin nhạy cảm.
- `KhoaSoTaiChinh`: trạng thái mở/khóa theo kỳ học và tháng.

## Quy tắc khóa sổ

Khi một tháng đã khóa, máy chủ từ chối thay đổi học phí, sổ thu chi, chuyển nguồn, kế hoạch chi, ngân sách và giao dịch gia đình. Kiểm tra nằm ở máy chủ nên không thể bỏ qua bằng cách thao tác trực tiếp từ trình duyệt.

## Hướng di chuyển tiếp theo

1. Tách Repository cho học sinh, điểm danh, học phí và sổ thu chi khỏi `Code.js`.
2. Chuẩn hóa khóa `MaCoSo` cho mô hình nhiều cơ sở.
3. Thay các sheet học phí theo tháng bằng một bảng giao dịch duy nhất có chỉ mục tháng.
4. Thêm adapter lưu trữ để có thể chuyển từ Google Sheet sang PostgreSQL/Cloud SQL mà không đổi Domain.
5. Đưa kiểm thử vào quy trình kiểm tra tự động trước khi triển khai.

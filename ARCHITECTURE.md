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
| `FinanceDomain.js` | Công thức học phí, lợi nhuận, dòng tiền, tiền an toàn và phân bổ hũ tài chính |
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
- `DanhMucHuTaiChinh`: danh mục hũ dùng chung, cho phép thêm và sửa; chỉ hũ chưa từng được tham chiếu mới được xoá.
- `HuTaiChinhThang`: tỷ lệ phân bổ các hũ tài chính theo kỳ học và tháng; số thực chi được tổng hợp từ sổ thu chi.

## Nguồn dữ liệu tài chính duy nhất

- Sáu hũ chuẩn chỉ là dữ liệu khởi tạo. Danh sách thực tế được đọc từ `DanhMucHuTaiChinh`, không viết cố định trong giao diện hay báo cáo.
- Không cho xoá hũ đang được danh mục thu chi, phiếu chi hoặc cấu hình tháng tham chiếu nhằm giữ nguyên lịch sử và tính đúng của báo cáo.
- `DanhMucThuChi.MaHuMacDinh` xác định hũ mặc định của từng danh mục chi.
- `SoThuChi.MaHuTaiChinh` lưu hũ thực tế được chọn trên phiếu chi; người dùng có thể đổi so với mặc định.
- Không tạo sổ giao dịch hũ riêng. Báo cáo hũ đọc trực tiếp các phiếu đang hoạt động trong `SoThuChi`, nên sửa hoặc huỷ phiếu không tạo dữ liệu trùng.
- Ngân sách hũ của tháng được phân bổ từ tổng học phí phải thu trong sheet tháng. Học phí `THU_HOC_PHI` đã ghi vào `SoThuChi` chỉ dùng để hiển thị tỷ lệ thu và cảnh báo dòng tiền.
- Mỗi tháng là một kế hoạch ngân sách độc lập, không mang số dư hũ từ tháng trước. Hạn mức còn lại bằng ngân sách kế hoạch trừ các phiếu chi đang hoạt động của chính tháng đó.

## Quy tắc khóa sổ

Khi một tháng đã khóa, máy chủ từ chối thay đổi học phí, sổ thu chi, chuyển nguồn, kế hoạch chi, ngân sách và giao dịch gia đình. Kiểm tra nằm ở máy chủ nên không thể bỏ qua bằng cách thao tác trực tiếp từ trình duyệt.

## Hướng di chuyển tiếp theo

1. Tách Repository cho học sinh, điểm danh, học phí và sổ thu chi khỏi `Code.js`.
2. Chuẩn hóa khóa `MaCoSo` cho mô hình nhiều cơ sở.
3. Thay các sheet học phí theo tháng bằng một bảng giao dịch duy nhất có chỉ mục tháng.
4. Thêm adapter lưu trữ để có thể chuyển từ Google Sheet sang PostgreSQL/Cloud SQL mà không đổi Domain.
5. Đưa kiểm thử vào quy trình kiểm tra tự động trước khi triển khai.

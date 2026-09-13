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
- `ChotPhanBoHu`: snapshot học phí nền và trạng thái chốt phương án phân bổ của từng kỳ học/tháng.

## Nguồn dữ liệu tài chính duy nhất

- Sáu hũ chuẩn chỉ là dữ liệu khởi tạo. Danh sách thực tế được đọc từ `DanhMucHuTaiChinh`, không viết cố định trong giao diện hay báo cáo.
- Không cho xoá hũ đang được danh mục thu chi, phiếu chi hoặc cấu hình tháng tham chiếu nhằm giữ nguyên lịch sử và tính đúng của báo cáo.
- `DanhMucThuChi.MaHuMacDinh` xác định hũ mặc định của từng danh mục chi.
- `SoThuChi.MaHuTaiChinh` lưu hũ thực tế được chọn trên phiếu chi; người dùng có thể đổi so với mặc định.
- Không tạo sổ giao dịch hũ riêng. Báo cáo hũ đọc trực tiếp các phiếu đang hoạt động trong `SoThuChi`, nên sửa hoặc huỷ phiếu không tạo dữ liệu trùng.
- Ngân sách hũ của tháng được phân bổ từ tổng học phí phải thu trong sheet tháng. Học phí `THU_HOC_PHI` đã ghi vào `SoThuChi` chỉ dùng để hiển thị tỷ lệ thu và cảnh báo dòng tiền.
- Mỗi tháng là một kế hoạch ngân sách độc lập, không mang số dư hũ từ tháng trước. Hạn mức còn lại bằng ngân sách kế hoạch trừ các phiếu chi đang hoạt động của chính tháng đó.
- Trước khi chốt, ngân sách dùng học phí phải thu hiện tại. Sau khi chốt, học phí nền và tỷ lệ hũ được giữ nguyên; phần tăng sau chốt chỉ hiển thị là doanh thu chưa phân bổ.
- `DanhMucHuTaiChinh.VaiTroHeThong = OWNER_COMPENSATION` xác định duy nhất Hũ lương chủ trung tâm. Không nhận diện hũ bằng tên hiển thị.
- Mỗi đợt rút tiền về gia đình là một phiếu `CHI_GIA_DINH` trong `SoThuChi` gắn với Hũ lương chủ trung tâm. Tài chính gia đình chỉ tổng hợp lại phiếu này, không tạo giao dịch thu trùng lặp.
- Thu nhập kế hoạch gia đình, tiền đã thực nhận và hạn mức còn được rút là ba số riêng biệt. Tiền thực còn lại của gia đình chỉ dùng các đợt rút đã phát sinh và thu nhập khác thực nhận.

## Quy tắc khóa sổ

Khi một tháng đã khóa, máy chủ từ chối thay đổi học phí, sổ thu chi, chuyển nguồn, kế hoạch chi, ngân sách và giao dịch gia đình. Kiểm tra nằm ở máy chủ nên không thể bỏ qua bằng cách thao tác trực tiếp từ trình duyệt.

Chốt phân bổ hũ khác với khóa sổ: chốt phân bổ chỉ khóa học phí nền và tỷ lệ hũ, nhưng vẫn cho phép thu, chi và rút tiền trong tháng. Chỉ `OWNER` được chốt hoặc mở chốt; mở chốt bắt buộc nhập lý do và ghi nhật ký.

## Hướng di chuyển tiếp theo

1. Tách Repository cho học sinh, điểm danh, học phí và sổ thu chi khỏi `Code.js`.
2. Chuẩn hóa khóa `MaCoSo` cho mô hình nhiều cơ sở.
3. Thay các sheet học phí theo tháng bằng một bảng giao dịch duy nhất có chỉ mục tháng.
4. Thêm adapter lưu trữ để có thể chuyển từ Google Sheet sang PostgreSQL/Cloud SQL mà không đổi Domain.
5. Đưa kiểm thử vào quy trình kiểm tra tự động trước khi triển khai.

## Phân tách giao diện tài chính

- `QuanLyTaiChinh.html`: tổng quan hiệu quả trung tâm và mô phỏng học phí; chỉ gọi `getTongQuanTaiChinhData`.
- `KeHoachTaiChinh.html`: hũ tài chính, nhân sự, khoản chi định kỳ và kế hoạch tháng. Mỗi tab được tải theo nhu cầu bằng `getKeHoachTaiChinhData`, không tải các tab chưa mở.
- `TaiChinhGiaDinh.html`: hạn mức Hũ lương chủ trung tâm, các đợt rút tiền và ngân sách gia đình; chỉ gọi `getTaiChinhGiaDinhData`.
- Ba màn hình dùng chung các Sheet, khóa sổ, nhật ký và hàm ghi nghiệp vụ hiện có. Không tạo bản sao giao dịch hoặc nguồn dữ liệu tài chính mới.
- Sau thao tác ghi, giao diện chỉ tải lại module đang sử dụng; dữ liệu danh mục ít thay đổi dùng cache theo phiên bản và bị vô hiệu hóa khi `bumpDataVersion_` được gọi.

## Phạm vi tất cả kỳ học

- Lựa chọn đăng nhập `__ALL_TERMS__` là phạm vi hiển thị, không phải mã kỳ học trong dữ liệu.
- `AllTermsService.js` và `AllTerms.html` hiển thị màn hình hiện tại theo từng kỳ học chưa xóa, kể cả kỳ đã ngừng hoạt động. Mỗi nhóm dùng iframe với phiên gắn đúng kỳ học và cùng tài khoản/quyền; dữ liệu và bộ lọc ngày/tháng giữ riêng từng nhóm. Không cộng gộp số dư hay loại bỏ học sinh xuất hiện ở nhiều kỳ.
- Chọn một kỳ cụ thể vẫn dùng luồng hiện có. Các RPC nghiệp vụ không nhận phiên tổng hợp để tránh ghi dữ liệu dưới mã giả.
- Các phiên con phụ thuộc phiên tổng hợp; đăng xuất phiên tổng hợp hoặc một phiên con thu hồi quyền sử dụng toàn bộ nhóm phiên.
- Kiểm tra cục bộ: `node tests/term-scope.cjs`. Chưa thay thế kiểm thử tích hợp trên Google Apps Script.

## Phân kỳ học sinh

- `PhanKyHocSinh.html` là bảng toàn bộ học sinh × kỳ học, truy cập từ nhóm Học sinh. Màn hình này hiển thị một lần ngay cả khi đăng nhập Tất cả kỳ học.
- `StudentTermService.js` đọc và cập nhật `HocSinhKyHoc`; quyền `student.read` để xem, `student.write` để lưu. Chỉ gửi các ô đã thay đổi và đọc lại dữ liệu dưới khóa trước khi ghi.
- Danh sách học sinh, điểm danh, sắp xếp và danh sách nguồn tạo học phí tháng chỉ nhận học sinh có liên kết chưa xóa trong `HocSinhKyHoc`. Không suy đoán từ `HocSinh.MaKyHoc` hoặc tự lấy học sinh khi kỳ chưa có liên kết.
- Bỏ tick gỡ liên kết bằng trạng thái DELETED; tick lại khôi phục liên kết và giữ học phí/ghi chú cũ. Học sinh có thể chưa thuộc kỳ nào và vẫn hiện trong bảng quản trị. Dữ liệu chỉ có mã kỳ cũ ở hồ sơ, chưa có liên kết, cần được tick trong bảng để xuất hiện trong danh sách theo kỳ.
- Thao tác phân kỳ không sửa snapshot học phí tháng, phiếu thu chi hay lịch sử điểm danh đã phát sinh.

## Học phí theo kỳ và thêm nhanh

- `HocPhiKyHoc.html` / `TermTuitionService.js`: cấu hình `KyHoc.HocPhiCap1` (khối 1–5), `HocPhiCap2` (khối 6–9). Kỳ chưa cấu hình dùng mức khởi đầu cũ 1.800.000 / 2.000.000 đồng.
- `HocSinhKyHoc.HocPhiMode`: AUTO nhận mức theo cấp, CUSTOM giữ mức riêng (bao gồm 0 đồng). Học phí cũ chưa có chế độ nhưng có số tiền được giữ như CUSTOM; có thể chuyển về AUTO trên màn hình học phí.
- Lưu cấu hình áp dụng và lưu lại học phí AUTO trong kỳ đó; không thay đổi CUSTOM, kỳ khác, snapshot tháng hoặc giao dịch đã phát sinh. Sửa hồ sơ chỉ chỉnh mức riêng của kỳ hiện tại, giữ mức riêng của các kỳ còn lại.
- Thêm mới, thêm nhanh, tick phân kỳ và chuyển tiếp sang kỳ khác đều lưu mức AUTO theo cấp và cấu hình kỳ đích.
- Thêm nhanh ở bảng phân kỳ nhận tên, khối/lớp, điện thoại và các kỳ tham gia. Có thể chưa gán kỳ; không tạo snapshot học phí tháng. Request ID chống tạo trùng khi gửi lại cùng yêu cầu trong thời hạn cache.

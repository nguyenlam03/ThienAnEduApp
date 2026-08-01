/**
 * Bộ kiểm thử đơn vị an toàn: chỉ kiểm tra hàm thuần, không thay đổi dữ liệu Sheet.
 * Chạy runArchitectureUnitTests() trong Apps Script để nhận báo cáo JSON.
 */
function runArchitectureUnitTests() {
  var results = [];
  function test(name, callback) {
    try {
      callback();
      results.push({ name: name, passed: true });
    } catch (error) {
      results.push({ name: name, passed: false, message: String(error && error.message || error) });
    }
  }
  function equal(actual, expected, message) {
    if (actual !== expected) throw new Error(message + ': nhận ' + actual + ', cần ' + expected);
  }
  function near(actual, expected, tolerance, message) {
    if (Math.abs(actual - expected) > tolerance) throw new Error(message + ': nhận ' + actual + ', cần gần ' + expected);
  }

  test('Mật khẩu được băm ổn định và không lưu nguyên văn', function () {
    var hash = SecurityService.hashPassword('mat-khau-kiem-thu');
    equal(hash, SecurityService.hashPassword('mat-khau-kiem-thu'), 'Hash không ổn định');
    if (hash === 'mat-khau-kiem-thu' || hash.length !== 64) throw new Error('Hash SHA-256 không hợp lệ');
  });
  test('Phân quyền giáo viên không được sửa tài chính', function () {
    equal(SecurityService.hasPermission('TEACHER', 'attendance.write'), true, 'Thiếu quyền điểm danh');
    equal(SecurityService.hasPermission('TEACHER', 'finance.write'), false, 'Giáo viên có quyền tài chính ngoài dự kiến');
  });
  test('Công thức học phí có xét tỷ lệ thu và biên lợi nhuận', function () {
    var pricing = FinanceDomain.calculatePricing({ studentCount: 86, plannedExpense: 99660000, targetMargin: 20, collectionRate: 100 });
    near(pricing.requiredAverageFee, 1448546.51, 1, 'Học phí bình quân cần thiết sai');
  });
  test('Hiệu quả tài chính tách lợi nhuận và tiền an toàn', function () {
    var result = FinanceDomain.calculatePerformance({ revenueForecast: 168000000, accruedRevenue: 168000000, plannedExpense: 99660000, currentCash: 100000000, remainingObligations: 45000000, reserveTarget: 10000000 });
    equal(result.projectedProfit, 68340000, 'Lợi nhuận dự kiến sai');
    equal(result.safeCash, 45000000, 'Tiền thực có thể sử dụng sai');
  });
  test('Định dạng tháng tài chính hợp lệ', function () {
    equal(GovernanceService.validateMonth('2026-08'), '2026-08', 'Không nhận tháng hợp lệ');
    var failed = false;
    try { GovernanceService.validateMonth('2026-13'); } catch (error) { failed = true; }
    equal(failed, true, 'Không chặn tháng sai');
  });
  test('Lệnh sổ thu chi chặn số tiền và ngày không hợp lệ', function () {
    var command = CashbookDomain.transaction({ loai: 'thu', ngayGiaoDich: '2026-08-01', maDanhMuc: 'THU_HOC_PHI', maNguonTien: 'TIEN_MAT', noiDung: 'Thu học phí', soTien: 1000000 });
    equal(command.loai, 'THU', 'Không chuẩn hóa loại giao dịch');
    var failed = false;
    try { CashbookDomain.transaction({ loai: 'CHI', ngayGiaoDich: '2026-02-31', maDanhMuc: 'CHI_KHAC', maNguonTien: 'TIEN_MAT', noiDung: 'Sai ngày', soTien: 1 }); } catch (error) { failed = true; }
    equal(failed, true, 'Không chặn ngày không tồn tại');
  });

  var failed = results.filter(function (item) { return !item.passed; }).length;
  return jsonResponse_({ passed: results.length - failed, failed: failed, total: results.length, results: results });
}

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
  test('Hũ tài chính phân bổ đúng học phí phải thu và hạn mức sau chi', function () {
    var result = FinanceDomain.calculateJars({
      revenue: 100000000,
      jars: [
        { code: 'VAN_HANH', name: 'Vận hành', ratio: 40, order: 1 },
        { code: 'LUONG', name: 'Lương', ratio: 30, order: 2 },
        { code: 'LOI_NHUAN', name: 'Lợi nhuận', ratio: 30, order: 3 }
      ],
      actualByJar: { VAN_HANH: 15000000, LUONG: 10000000 }
    });
    equal(result.summary.ratioTotal, 100, 'Tổng tỷ lệ hũ sai');
    equal(result.summary.allocatedTotal, 100000000, 'Tổng tiền phân bổ sai');
    equal(result.items[0].remaining, 25000000, 'Số dư hũ vận hành sai');
  });
  test('Hũ tài chính của mỗi tháng không mang số dư tháng trước', function () {
    var result = FinanceDomain.calculateJars({
      revenue: 10000000,
      jars: [{ code: 'DU_PHONG', name: 'Dự phòng', ratio: 10, order: 1 }],
      actualByJar: { DU_PHONG: 200000 }
    });
    equal(result.items[0].allocated, 1000000, 'Ngân sách tháng sai');
    equal(result.items[0].remaining, 800000, 'Hạn mức còn lại sai');
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
  test('Phiếu chi giữ hũ tài chính do người dùng lựa chọn', function () {
    var command = CashbookDomain.transaction({ loai: 'CHI', ngayGiaoDich: '2026-08-01', maDanhMuc: 'CHI_MARKETING', maNhanSu: 'NS_001', maHuTaiChinh: 'dau_tu', maNguonTien: 'TIEN_MAT', noiDung: 'Quảng cáo tuyển sinh', soTien: 2000000 });
    equal(command.maHuTaiChinh, 'DAU_TU', 'Không giữ hũ tài chính trên phiếu chi');
    equal(command.maNhanSu, 'NS_001', 'Không giữ mã nhân sự để đối soát khoản chi');
  });

  test('Nguồn tiền trung tâm và gia đình được tách phạm vi', function () {
    equal(getNguonTienDefinition_('BIDV').phamVi, 'TRUNG_TAM', 'BIDV trung tâm sai phạm vi');
    equal(getNguonTienDefinition_('GD_BIDV').phamVi, 'GIA_DINH', 'BIDV cá nhân sai phạm vi');
    equal(normalizeFinanceScope_('gia_dinh'), 'GIA_DINH', 'Không chuẩn hóa phạm vi gia đình');
  });

  test('Phiếu rút lương chủ sở hữu cũ được đối soát với phiếu chi trung tâm', function () {
    var matched = isLegacyOwnerWithdrawalMatch_(
      { loai: 'CHI', ngayGiaoDich: '2026-08-01', soTien: 6000000, noiDung: 'Rút tiền lương thầy Lãm đợt 1' },
      { loai: 'CHI', maDanhMuc: 'CHI_GIA_DINH', phamVi: 'TRUNG_TAM', trangThai: 'HOAT_DONG', ngayGiaoDich: '2026-08-01', soTien: 6000000 }
    );
    equal(matched, true, 'Không nhận diện phiếu rút lương trung tâm');
    var familyExpense = isLegacyOwnerWithdrawalMatch_(
      { loai: 'CHI', ngayGiaoDich: '2026-08-01', soTien: 6000000, noiDung: 'Đóng tiền học cho Pin' },
      { loai: 'CHI', maDanhMuc: 'CHI_GIA_DINH', phamVi: 'TRUNG_TAM', trangThai: 'HOAT_DONG', ngayGiaoDich: '2026-08-01', soTien: 6000000 }
    );
    equal(familyExpense, false, 'Nhầm khoản chi gia đình thực tế với phiếu rút lương');
  });

  test('Mã hũ UUID cũ có dấu gạch ngang vẫn được nhận diện', function () {
    equal(normalizeMaHuTaiChinh_('HU_2D4703A1-9'), 'HU_2D4703A1-9', 'Hũ cũ bị lọc khỏi danh sách');
  });
  test('Tài chính gia đình tách hạn mức kế hoạch khỏi tiền đã thực nhận', function () {
    var result = FinanceDomain.calculateFamilySource({
      plannedOwnerIncome: 25822500, ownerReceived: 6000000, otherIncome: 1000000,
      totalExpense: 2500000, totalSaving: 500000
    });
    equal(result.remainingOwnerAllowance, 19822500, 'Sai hạn mức chủ cơ sở còn được rút');
    equal(result.totalIncome, 7000000, 'Đã cộng tiền kế hoạch vào tiền thực nhận');
    equal(result.remaining, 4000000, 'Sai số tiền gia đình thực còn lại');
  });

  var failed = results.filter(function (item) { return !item.passed; }).length;
  return jsonResponse_({ passed: results.length - failed, failed: failed, total: results.length, results: results });
}

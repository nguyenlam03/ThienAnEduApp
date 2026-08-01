/**
 * Repository: chỉ lớp này biết cấu trúc Sheet nhật ký và khóa sổ.
 */
var GovernanceRepository = (function () {
  var AUDIT_SHEET = 'NhatKyHeThong';
  var PERIOD_SHEET = 'KhoaSoTaiChinh';

  function auditHeaders() {
    return [
      'MaNhatKy', 'ThoiGian', 'MaNguoiDung', 'TenDangNhap', 'VaiTro', 'MaKyHoc',
      'HanhDong', 'DoiTuong', 'MaDoiTuong', 'DuLieuTruocJson', 'DuLieuSauJson',
      'MetadataJson'
    ];
  }

  function periodHeaders() {
    return [
      'MaKhoaSo', 'MaKyHoc', 'Thang', 'TrangThai', 'LyDo',
      'KhoaLuc', 'KhoaBoi', 'MoKhoaLuc', 'MoKhoaBoi', 'UpdatedAt'
    ];
  }

  function ensureSchema() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheet_(ss, AUDIT_SHEET, auditHeaders());
    ensureSheet_(ss, PERIOD_SHEET, periodHeaders());
  }

  function appendAudit(entry) {
    var sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), AUDIT_SHEET, auditHeaders());
    appendObjectsToSheet_(sheet, [entry], auditHeaders());
  }

  function listAudit() {
    ensureSchema();
    return readObjectsNoCache_(AUDIT_SHEET);
  }

  function findPeriod(maKyHoc, month) {
    ensureSchema();
    return readObjectsNoCache_(PERIOD_SHEET).find(function (row) {
      return String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
        String(row.Thang || '').trim() === String(month || '').trim();
    }) || null;
  }

  function savePeriod(maKyHoc, month, values) {
    var sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), PERIOD_SHEET, periodHeaders());
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0].map(function (item) { return String(item || '').trim(); });
    var index = buildHeaderIndex_(headers);
    var target = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][index.MaKyHoc] || '').trim() === maKyHoc && String(rows[i][index.Thang] || '').trim() === month) {
        target = i;
        break;
      }
    }
    var row = target >= 0 ? rows[target].slice() : new Array(headers.length).fill('');
    row[index.MaKhoaSo] = row[index.MaKhoaSo] || ('KSO_' + Utilities.getUuid().slice(0, 10).toUpperCase());
    row[index.MaKyHoc] = maKyHoc;
    row[index.Thang] = month;
    Object.keys(values).forEach(function (key) { if (index[key] !== undefined) row[index[key]] = values[key]; });
    row[index.UpdatedAt] = new Date();
    if (target >= 0) sheet.getRange(target + 1, 1, 1, headers.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
    return row.reduce(function (object, value, column) {
      object[headers[column]] = value;
      return object;
    }, {});
  }

  return {
    ensureSchema: ensureSchema,
    appendAudit: appendAudit,
    listAudit: listAudit,
    findPeriod: findPeriod,
    savePeriod: savePeriod
  };
})();

/**
 * Service: quy tắc kiểm toán và khóa sổ tài chính.
 */
var GovernanceService = (function () {
  function validateMonth(month) {
    month = String(month || '').trim();
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Tháng tài chính không hợp lệ.');
    return month;
  }

  function compactJson(value) {
    if (value === undefined || value === null) return '';
    var text = JSON.stringify(value, function (key, item) {
      if (/matkhau|password|token|dataurl/i.test(key)) return '[REDACTED]';
      return item;
    });
    return text.length > 20000 ? text.slice(0, 20000) + '…' : text;
  }

  function audit(session, action, entity, entityId, before, after, metadata) {
    if (!session || !session.valid) return;
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) throw new Error('Không thể ghi nhật ký kiểm toán lúc này.');
    try {
      GovernanceRepository.appendAudit({
        MaNhatKy: 'AUD_' + Utilities.getUuid().slice(0, 12).toUpperCase(),
        ThoiGian: new Date(),
        MaNguoiDung: session.maNguoiDung || '',
        TenDangNhap: session.tenDangNhap || '',
        VaiTro: session.vaiTro || '',
        MaKyHoc: session.maKyHoc || '',
        HanhDong: String(action || '').trim().toUpperCase(),
        DoiTuong: String(entity || '').trim().toUpperCase(),
        MaDoiTuong: String(entityId || '').trim(),
        DuLieuTruocJson: compactJson(before),
        DuLieuSauJson: compactJson(after),
        MetadataJson: compactJson(metadata)
      });
    } finally { lock.releaseLock(); }
  }

  function periodStatus(session, month) {
    month = validateMonth(month);
    var row = GovernanceRepository.findPeriod(session.maKyHoc, month) || {};
    var locked = String(row.TrangThai || 'OPEN').trim().toUpperCase() === 'LOCKED';
    return {
      thang: month,
      locked: locked,
      trangThai: locked ? 'LOCKED' : 'OPEN',
      lyDo: String(row.LyDo || '').trim(),
      khoaLuc: formatDateDisplay_(row.KhoaLuc),
      khoaBoi: String(row.KhoaBoi || '').trim(),
      moKhoaLuc: formatDateDisplay_(row.MoKhoaLuc),
      moKhoaBoi: String(row.MoKhoaBoi || '').trim()
    };
  }

  function assertOpen(session, month) {
    var status = periodStatus(session, month);
    if (status.locked) {
      throw new Error('Tháng ' + month.slice(5, 7) + '/' + month.slice(0, 4) + ' đã khóa sổ. Không thể thêm, sửa, xóa hoặc đồng bộ dữ liệu tài chính.');
    }
  }

  function lockPeriod(session, month, reason) {
    month = validateMonth(month);
    var existing = periodStatus(session, month);
    if (existing.locked) return existing;
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật trạng thái khóa sổ.');
    var saved;
    try {
      saved = GovernanceRepository.savePeriod(session.maKyHoc, month, {
        TrangThai: 'LOCKED', LyDo: String(reason || '').trim(),
        KhoaLuc: new Date(), KhoaBoi: session.tenDangNhap || session.hoTen || session.maNguoiDung,
        MoKhoaLuc: '', MoKhoaBoi: ''
      });
    } finally { lock.releaseLock(); }
    audit(session, 'LOCK', 'KY_TAI_CHINH', month, existing, saved, null);
    return periodStatus(session, month);
  }

  function unlockPeriod(session, month, reason) {
    month = validateMonth(month);
    var existing = periodStatus(session, month);
    if (!existing.locked) return existing;
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật trạng thái khóa sổ.');
    var saved;
    try {
      saved = GovernanceRepository.savePeriod(session.maKyHoc, month, {
        TrangThai: 'OPEN', LyDo: String(reason || existing.lyDo || '').trim(),
        MoKhoaLuc: new Date(), MoKhoaBoi: session.tenDangNhap || session.hoTen || session.maNguoiDung
      });
    } finally { lock.releaseLock(); }
    audit(session, 'UNLOCK', 'KY_TAI_CHINH', month, existing, saved, { lyDoMoKhoa: String(reason || '').trim() });
    return periodStatus(session, month);
  }

  function listAudit(session, filters) {
    filters = filters || {};
    var from = toDateOnly_(filters.fromDate);
    var to = toDateOnly_(filters.toDate);
    if (to) to = new Date(to.getFullYear(), to.getMonth(), to.getDate(), 23, 59, 59, 999);
    return GovernanceRepository.listAudit().filter(function (row) {
      var time = row.ThoiGian instanceof Date ? row.ThoiGian : new Date(row.ThoiGian);
      return (!filters.maKyHoc || String(row.MaKyHoc || '').trim() === String(filters.maKyHoc).trim()) &&
        (!filters.doiTuong || String(row.DoiTuong || '').trim() === String(filters.doiTuong).trim().toUpperCase()) &&
        (!from || time >= from) && (!to || time <= to);
    }).slice(-1000).reverse().map(function (row) {
      return {
        maNhatKy: String(row.MaNhatKy || '').trim(),
        thoiGian: row.ThoiGian,
        tenDangNhap: String(row.TenDangNhap || '').trim(),
        vaiTro: String(row.VaiTro || '').trim(),
        maKyHoc: String(row.MaKyHoc || '').trim(),
        hanhDong: String(row.HanhDong || '').trim(),
        doiTuong: String(row.DoiTuong || '').trim(),
        maDoiTuong: String(row.MaDoiTuong || '').trim(),
        duLieuTruocJson: String(row.DuLieuTruocJson || ''),
        duLieuSauJson: String(row.DuLieuSauJson || ''),
        metadataJson: String(row.MetadataJson || '')
      };
    });
  }

  return {
    audit: audit,
    periodStatus: periodStatus,
    assertOpen: assertOpen,
    lockPeriod: lockPeriod,
    unlockPeriod: unlockPeriod,
    listAudit: listAudit,
    validateMonth: validateMonth
  };
})();

function ensureArchitectureSheets_() {
  SecurityService.ensureSchema();
  GovernanceRepository.ensureSchema();
}

function safeWriteAuditLog_(session, action, entity, entityId, before, after, metadata) {
  try { GovernanceService.audit(session, action, entity, entityId, before, after, metadata); }
  catch (error) { console.error('Audit log failed: ' + error); }
}

function assertFinancePeriodOpen_(session, month) {
  GovernanceService.assertOpen(session, month);
}

function getFinancePeriodStatus(token, month) {
  var session = SecurityService.requireSession(token, 'finance.read');
  return jsonResponse_(GovernanceService.periodStatus(session, month));
}

function lockFinancePeriod(token, month, reason) {
  var session = SecurityService.requireSession(token, 'finance.close');
  return jsonResponse_({ success: true, status: GovernanceService.lockPeriod(session, month, reason), message: 'Đã khóa sổ tài chính tháng.' });
}

function unlockFinancePeriod(token, month, reason) {
  var session = SecurityService.requireSession(token, 'finance.close');
  return jsonResponse_({ success: true, status: GovernanceService.unlockPeriod(session, month, reason), message: 'Đã mở khóa sổ tài chính tháng.' });
}

function getAuditLog(token, filters) {
  var session = SecurityService.requireSession(token, 'system.admin');
  return jsonResponse_(GovernanceService.listAudit(session, filters));
}

/** Global student/term assignment. Only explicit HocSinh_KyHoc links define membership. */
function requireStudentTermSession_(token, permission) {
  const session = SecurityService.requireSession(token);
  if (!SecurityService.hasPermission(session.vaiTro, permission)) throw new Error('Tài khoản không có quyền quản lý phân kỳ học sinh.');
  return session;
}

function getStudentTermMatrix(token) {
  const session = requireStudentTermSession_(token, 'student.read');
  const relations = readObjectsNoCache_(SHEET_HOCSINH_KYHOC);
  const terms = getAllTerms_();
  const validTerms = new Set(terms.map(term => term.maKyHoc));
  const membership = {};
  relations.forEach(row => {
    const id = String(row.MaHocSinh || '').trim(), term = String(row.MaKyHoc || '').trim();
    if (!id || !validTerms.has(term) || String(row.TrangThai || '').toUpperCase().trim() === 'DELETED') return;
    if (!membership[id]) membership[id] = [];
    if (membership[id].indexOf(term) === -1) membership[id].push(term);
  });
  const students = readObjectsNoCache_(SHEET_HOCSINH)
    .filter(row => row.MaHocSinh && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED')
    .map(row => ({ maHocSinh: String(row.MaHocSinh).trim(), hoTen: String(row.HoTen || ''),
      khoi: String(row.Khoi || ''), lop: String(row.Lop || ''), sapXep: number_(row.SapXep),
      kyHocIds: membership[String(row.MaHocSinh).trim()] || [] }))
    .sort(compareStudentSort_);
  return jsonResponse_({ students: students, terms: terms, khoiList: getKhoiList_(), lopList: getLopList_(), canWrite: SecurityService.hasPermission(session.vaiTro, 'student.write') });
}

// Pure update: retain fee, notes and history on every untouched link and on reactivation.
function applyStudentTermChanges_(rows, changes, studentIds, termIds, now) {
  if (!Array.isArray(changes) || changes.length > 10000) throw new Error('Danh sách thay đổi không hợp lệ.');
  const result = rows.map(row => Object.assign({}, row));
  const seen = new Set();
  changes.forEach(change => {
    if (!change || typeof change.selected !== 'boolean' || typeof change.expected !== 'boolean') throw new Error('Trạng thái chọn kỳ học không hợp lệ.');
    const id = String(change.maHocSinh || '').trim(), term = String(change.maKyHoc || '').trim();
    if (!studentIds.has(id) || !termIds.has(term)) throw new Error('Học sinh hoặc kỳ học không còn tồn tại. Vui lòng tải lại bảng.');
    const key = JSON.stringify([id, term]);
    if (seen.has(key)) throw new Error('Một ô kỳ học bị gửi trùng.');
    seen.add(key);
    const matches = result.filter(row => String(row.MaHocSinh || '').trim() === id && String(row.MaKyHoc || '').trim() === term);
    const active = matches.filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
    const current = active.length > 0;
    if (current !== change.expected && current !== change.selected) throw new Error('Phân kỳ học đã được người khác cập nhật. Vui lòng tải lại bảng.');
    if (current === change.selected) return;
    if (!change.selected) {
      active.forEach(row => { row.TrangThai = 'DELETED'; row.UpdatedAt = now; });
    } else if (matches.length) {
      const row = matches[matches.length - 1];
      row.TrangThai = 'ACTIVE'; row.UpdatedAt = now;
    } else {
      result.push({ MaHocSinh: id, MaKyHoc: term, HocPhi: '', TrangThaiHocPhi: '', GhiChuHocPhi: '',
        TrangThai: 'ACTIVE', CreatedAt: now, UpdatedAt: now });
    }
  });
  return result;
}

function saveStudentTermMatrix(token, changes) {
  const session = requireStudentTermSession_(token, 'student.write');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Có người đang cập nhật học sinh. Vui lòng thử lại.');
  try {
    const students = readObjectsNoCache_(SHEET_HOCSINH).filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
    const terms = readObjectsNoCache_(SHEET_KYHOC).filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
    const rows = readObjectsNoCache_(SHEET_HOCSINH_KYHOC);
    let next = applyStudentTermChanges_(rows, changes,
      new Set(students.map(row => String(row.MaHocSinh || '').trim()).filter(Boolean)),
      new Set(terms.map(row => String(row.MaKyHoc || '').trim()).filter(id => id && id !== ALL_TERMS_CODE_)), new Date());
    const changedIds = new Set(changes.map(change => change.maHocSinh));
    next = fillAutomaticTuition_(next, students.filter(row => changedIds.has(String(row.MaHocSinh || '').trim())), terms);
    if (changes.length) {
      writeObjectsToSheet_(SHEET_HOCSINH_KYHOC, next, getHocSinhKyHocHeaders_());
      bumpDataVersion_();
      safeWriteAuditLog_(session, 'UPDATE', 'PHAN_KY_HOC_SINH', '', null, { changes: changes });
    }
    return jsonResponse_({ success: true, message: 'Đã lưu phân kỳ học cho học sinh.' });
  } finally { lock.releaseLock(); }
}
function quickAddStudent(token, data) {
  const session = requireStudentTermSession_(token, 'student.write');
  data = data || {};
  const requestId = String(data.requestId || '');
  if (!/^[a-zA-Z0-9_-]{8,80}$/.test(requestId)) throw new Error('Vui lòng mở lại form thêm nhanh.');
  const requestKey = 'QUICK_STUDENT_' + session.maNguoiDung + '_' + requestId;
  const name = String(data.hoTen || '').trim(), grade = String(data.khoi || '').trim(), classId = String(data.lop || '').trim();
  const phone = String(data.sdtPhuHuynh || '').trim();
  if (!name || name.length > 200) throw new Error('Vui lòng nhập tên học sinh, tối đa 200 ký tự.');
  if (phone && !/^0\d{9}$/.test(phone)) throw new Error('Số điện thoại phải gồm 10 chữ số và bắt đầu bằng 0.');
  if (!Array.isArray(data.kyHocIds)) throw new Error('Danh sách kỳ học không hợp lệ.');
  const selected = Array.from(new Set(data.kyHocIds.map(id => String(id || '').trim())));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Có người đang cập nhật học sinh. Vui lòng thử lại.');
  try {
    const previous = CacheService.getScriptCache().get(requestKey);
    if (previous) { const result = JSON.parse(previous); if (result.error) throw new Error(result.error); return previous; }
    const classRow = getLopList_().find(row => row.maLop === classId && row.khoi === grade);
    if (!classRow) throw new Error('Vui lòng chọn đúng khối và lớp.');
    const terms = readObjectsNoCache_(SHEET_KYHOC).filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
    const termIds = new Set(terms.map(row => String(row.MaKyHoc || '').trim()));
    if (selected.some(id => !id || id === ALL_TERMS_CODE_ || !termIds.has(id))) throw new Error('Kỳ học không hợp lệ.');
    tuitionForGrade_(grade, { cap1: 0, cap2: 0 });
    const classOrder = Number(classRow.thuTu);
    if (!Number.isInteger(classOrder) || classOrder < 1 || classOrder > 9) throw new Error('Lớp chưa có thứ tự hợp lệ.');
    const students = readObjectsNoCache_(SHEET_HOCSINH);
    const used = new Set(students.filter(row => String(row.Lop || '').trim() === classId && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED').map(row => number_(row.SapXep)));
    let order = classOrder * 100;
    while (used.has(order) && order <= classOrder * 100 + 99) order++;
    if (order > classOrder * 100 + 99) throw new Error('Lớp đã dùng hết dải thứ tự học sinh.');
    const id = 'HS_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase(), now = new Date();
    const student = { MaHocSinh: id, MaKyHoc: selected[0] || '', HoTen: name, Khoi: grade, Lop: classId, SapXep: order,
      Truong: Number(grade) <= 5 ? 'TH Tam Thiện' : 'THCS Phước Thái', SDTPhuHuynh: phone,
      KhongThuPhi: 'Không', TrangThai: 'ACTIVE', CreatedAt: now, UpdatedAt: now };
    const links = selected.map(term => ({ MaHocSinh: id, MaKyHoc: term, HocPhi: '', HocPhiMode: 'AUTO', TrangThaiHocPhi: '', GhiChuHocPhi: '', TrangThai: 'ACTIVE', CreatedAt: now, UpdatedAt: now }));
    const priced = fillAutomaticTuition_(links, [student], terms);
    // Validate both records before the first write. This does not create monthly fee snapshots.
    appendObjectsToSheet_(ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_HOCSINH, getHocSinhHeaders_()), [student], getHocSinhHeaders_());
    try {
      if (priced.length) appendObjectsToSheet_(ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_HOCSINH_KYHOC, getHocSinhKyHocHeaders_()), priced, getHocSinhKyHocHeaders_());
    } catch (error) {
      const message = 'Đã tạo hồ sơ ' + id + ' nhưng chưa lưu đủ kỳ học. Hãy tải lại bảng và tick kỳ học cho hồ sơ này; không thêm lại học sinh.';
      CacheService.getScriptCache().put(requestKey, JSON.stringify({ error: message }), 21600);
      throw new Error(message);
    } finally { bumpDataVersion_(); }
    safeWriteAuditLog_(session, 'CREATE', 'HOC_SINH', id, null, { hoTen: name, khoi: grade, lop: classId, kyHocIds: selected });
    const response = jsonResponse_({ success: true, student: { maHocSinh: id, hoTen: name, khoi: grade, lop: classId, sapXep: order, kyHocIds: selected }, message: 'Đã thêm học sinh và lưu học phí theo kỳ.' });
    CacheService.getScriptCache().put(requestKey, response, 21600);
    return response;
  } finally { lock.releaseLock(); }
}

function tuitionAmount_(value) {
  if (value === '' || value === null || value === undefined) throw new Error('Vui lòng nhập học phí.');
  const amount = Number(value);
  if (!Number.isSafeInteger(amount) || amount < 0 || amount > 1000000000) throw new Error('Học phí phải là số nguyên từ 0 đến 1 tỷ đồng.');
  return amount;
}
function tuitionMode_(row) {
  if (row.HocPhiMode === 'AUTO' || row.HocPhiMode === 'CUSTOM') return row.HocPhiMode;
  // Without historical mode metadata, preserve an existing amount as an individual rate.
  return row.HocPhi !== '' && row.HocPhi !== null && row.HocPhi !== undefined ? 'CUSTOM' : 'AUTO';
}
function termTuitionConfig_(term) {
  term = term || {};
  return { cap1: term.HocPhiCap1 === '' || term.HocPhiCap1 == null ? 1800000 : tuitionAmount_(term.HocPhiCap1),
    cap2: term.HocPhiCap2 === '' || term.HocPhiCap2 == null ? 2000000 : tuitionAmount_(term.HocPhiCap2) };
}
function tuitionForGrade_(khoi, config) {
  const grade = Number(khoi);
  if (!Number.isInteger(grade) || grade < 1 || grade > 9) throw new Error('Khối học sinh phải từ 1 đến 9 để xác định cấp học.');
  return grade <= 5 ? config.cap1 : config.cap2;
}
function getTermDefaultTuition_(termId, khoi) {
  const term = readObjects_(SHEET_KYHOC).find(row => String(row.MaKyHoc || '').trim() === termId);
  if (!term) throw new Error('Không tìm thấy kỳ học để lấy học phí.');
  return tuitionForGrade_(khoi, termTuitionConfig_(term));
}
function resolveStudentTermTuition_(relation, student, term) {
  return tuitionMode_(relation) === 'CUSTOM' ? tuitionAmount_(relation.HocPhi) : tuitionForGrade_(student.Khoi, termTuitionConfig_(term));
}
function fillAutomaticTuition_(relations, students, terms) {
  const studentMap = new Map(students.map(row => [String(row.MaHocSinh || '').trim(), row]));
  const termMap = new Map(terms.map(row => [String(row.MaKyHoc || '').trim(), row]));
  return relations.map(row => {
    if (String(row.TrangThai || '').trim().toUpperCase() === 'DELETED') return row;
    const student = studentMap.get(String(row.MaHocSinh || '').trim()), term = termMap.get(String(row.MaKyHoc || '').trim());
    if (!student || !term) return row;
    const mode = tuitionMode_(row);
    return Object.assign({}, row, { HocPhiMode: mode, HocPhi: resolveStudentTermTuition_(row, student, term) });
  });
}
function getTermTuitionData(token, termId) {
  const session = requireStudentTermSession_(token, 'student.read');
  const terms = getAllTerms_();
  const selected = String(termId || (session.maKyHoc === ALL_TERMS_CODE_ ? (terms[0] || {}).maKyHoc : session.maKyHoc) || '');
  const term = readObjectsNoCache_(SHEET_KYHOC).find(row => String(row.MaKyHoc || '').trim() === selected && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
  if (!term) return jsonResponse_({ terms: terms, maKyHoc: '', students: [], canWrite: SecurityService.hasPermission(session.vaiTro, 'student.write') });
  const relationMap = new Map(readObjectsNoCache_(SHEET_HOCSINH_KYHOC).filter(row => String(row.MaKyHoc || '').trim() === selected && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED').map(row => [String(row.MaHocSinh || '').trim(), row]));
  const students = readObjectsNoCache_(SHEET_HOCSINH).filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED' && relationMap.has(String(row.MaHocSinh || '').trim()))
    .map(row => { const relation = relationMap.get(String(row.MaHocSinh).trim()); return { maHocSinh: String(row.MaHocSinh).trim(), hoTen: String(row.HoTen || ''), khoi: String(row.Khoi || ''), lop: String(row.Lop || ''),
      mode: tuitionMode_(relation), amount: resolveStudentTermTuition_(relation, row, term) }; });
  return jsonResponse_({ terms: terms, maKyHoc: selected, config: termTuitionConfig_(term), students: students, canWrite: SecurityService.hasPermission(session.vaiTro, 'student.write') });
}
function saveTermTuition(token, data) {
  const session = requireStudentTermSession_(token, 'student.write');
  data = data || {};
  const cap1 = tuitionAmount_(data.cap1), cap2 = tuitionAmount_(data.cap2);
  if (!Array.isArray(data.overrides) || data.overrides.length > 10000) throw new Error('Danh sách học phí riêng không hợp lệ.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Có người đang cập nhật học phí. Vui lòng thử lại.');
  let changed = false;
  try {
    const terms = readObjectsNoCache_(SHEET_KYHOC).map(row => Object.assign({}, row));
    const term = terms.find(row => String(row.MaKyHoc || '').trim() === String(data.maKyHoc || '').trim() && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
    if (!term) throw new Error('Kỳ học không tồn tại.');
    term.HocPhiCap1 = cap1; term.HocPhiCap2 = cap2;
    const students = readObjectsNoCache_(SHEET_HOCSINH).filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
    const ids = new Set(students.map(row => String(row.MaHocSinh || '').trim()));
    const relations = readObjectsNoCache_(SHEET_HOCSINH_KYHOC).map(row => Object.assign({}, row));
    const seen = new Set();
    data.overrides.forEach(item => {
      if (!item || !ids.has(item.maHocSinh) || seen.has(item.maHocSinh) || !['AUTO','CUSTOM'].includes(item.mode)) throw new Error('Học sinh hoặc chế độ học phí không hợp lệ.');
      seen.add(item.maHocSinh);
      const links = relations.filter(row => String(row.MaHocSinh || '').trim() === item.maHocSinh && String(row.MaKyHoc || '').trim() === String(term.MaKyHoc).trim() && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
      if (!links.length) throw new Error('Học sinh không còn thuộc kỳ học này. Vui lòng tải lại.');
      const amount = item.mode === 'CUSTOM' ? tuitionAmount_(item.amount) : '';
      links.forEach(row => { row.HocPhiMode = item.mode; row.HocPhi = amount; row.UpdatedAt = new Date(); });
    });
    const selectedLinks = relations.filter(row => String(row.MaKyHoc || '').trim() === String(term.MaKyHoc).trim());
    const updated = fillAutomaticTuition_(selectedLinks, students, [term]);
    let index = 0;
    const next = relations.map(row => String(row.MaKyHoc || '').trim() === String(term.MaKyHoc).trim() ? updated[index++] : row);
    writeObjectsToSheet_(SHEET_KYHOC, terms, ['MaKyHoc','TenKyHoc','TrangThai','MacDinh','HocPhiCap1','HocPhiCap2']);
    changed = true;
    writeObjectsToSheet_(SHEET_HOCSINH_KYHOC, next, getHocSinhKyHocHeaders_());
    safeWriteAuditLog_(session, 'UPDATE', 'HOC_PHI_KY_HOC', String(term.MaKyHoc), null, data);
    return jsonResponse_({ success: true, message: 'Đã lưu học phí kỳ học và học phí từng học sinh.' });
  } finally { if (changed) bumpDataVersion_(); lock.releaseLock(); }
}

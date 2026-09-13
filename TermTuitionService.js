// Accept explicit grade labels, never extract arbitrary digits from class names.
function normalizeTuitionGrade_(value) {
  const text = String(value == null ? '' : value).trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const match = text.match(/^(?:(?:khoi|k|lop)[ _-]*)?0?([1-9])$/);
  return match ? Number(match[1]) : null;
}
function tuitionGradeContext_() {
  return { grades: readObjects_(SHEET_KHOI), classes: readObjects_(SHEET_LOP) };
}
function resolveTuitionGrade_(student, context) {
  context = context || { grades: [], classes: [] };
  const code = String(student.Khoi || '').trim();
  function fromCode(value) {
    const direct = normalizeTuitionGrade_(value);
    if (direct !== null) return direct;
    const definition = context.grades.find(row => String(row.Khoi || '').trim() === String(value || '').trim());
    return definition ? normalizeTuitionGrade_(definition.TenKhoi) : null;
  }
  if (code) return fromCode(code);
  const classRow = context.classes.find(row => String(row.MaLop || '').trim() === String(student.Lop || '').trim());
  if (classRow) return String(classRow.Khoi || '').trim() ? fromCode(classRow.Khoi) : normalizeTuitionGrade_(classRow.TenLop);
  return normalizeTuitionGrade_(student.Lop);
}
function studentTuitionState_(relation, student, term, context) {
  const grade = resolveTuitionGrade_(student, context);
  const mode = tuitionMode_(relation);
  const issue = grade === null ? 'Chưa xác định khối 1–9. Vui lòng sửa khối/lớp trong hồ sơ học sinh.' : '';
  if (mode === 'CUSTOM') return { grade: grade, amount: tuitionAmount_(relation.HocPhi), issue: issue };
  if (grade !== null) return { grade: grade, amount: tuitionForGrade_(grade, termTuitionConfig_(term)), issue: '' };
  // An unresolved automatic rate is not zero and must not be assigned to either tier.
  const stored = relation.HocPhi;
  const amount = stored !== '' && stored != null && Number.isSafeInteger(Number(stored)) && Number(stored) >= 0 && Number(stored) <= 1000000000 ? Number(stored) : null;
  return { grade: null, amount: amount, issue: issue };
}
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
  const grade = normalizeTuitionGrade_(khoi);
  if (!Number.isInteger(grade) || grade < 1 || grade > 9) throw new Error('Khối học sinh phải từ 1 đến 9 để xác định cấp học.');
  return grade <= 5 ? config.cap1 : config.cap2;
}
function getTermDefaultTuition_(termId, khoi) {
  const term = readObjects_(SHEET_KYHOC).find(row => String(row.MaKyHoc || '').trim() === termId);
  if (!term) throw new Error('Không tìm thấy kỳ học để lấy học phí.');
  return tuitionForGrade_(resolveTuitionGrade_({ Khoi: khoi }, tuitionGradeContext_()), termTuitionConfig_(term));
}
function resolveStudentTermTuition_(relation, student, term) {
  return tuitionMode_(relation) === 'CUSTOM' ? tuitionAmount_(relation.HocPhi) : tuitionForGrade_(student.Khoi, termTuitionConfig_(term));
}
function fillAutomaticTuition_(relations, students, terms) {
  const context = tuitionGradeContext_();
  const studentMap = new Map(students.map(row => [String(row.MaHocSinh || '').trim(), row]));
  const termMap = new Map(terms.map(row => [String(row.MaKyHoc || '').trim(), row]));
  return relations.map(row => {
    if (String(row.TrangThai || '').trim().toUpperCase() === 'DELETED') return row;
    const student = studentMap.get(String(row.MaHocSinh || '').trim()), term = termMap.get(String(row.MaKyHoc || '').trim());
    if (!student || !term) return row;
    const mode = tuitionMode_(row);
    const state = studentTuitionState_(row, student, term, context);
    if (mode === 'AUTO' && state.grade === null) return row;
    return Object.assign({}, row, { HocPhiMode: mode, HocPhi: state.amount });
  });
}
function getTermTuitionData(token, termId) {
  const session = requireStudentTermSession_(token, 'student.read');
  const terms = getAllTerms_();
  const selected = String(termId || (session.maKyHoc === ALL_TERMS_CODE_ ? (terms[0] || {}).maKyHoc : session.maKyHoc) || '');
  const term = readObjectsNoCache_(SHEET_KYHOC).find(row => String(row.MaKyHoc || '').trim() === selected && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
  if (!term) return jsonResponse_({ terms: terms, maKyHoc: '', students: [], canWrite: SecurityService.hasPermission(session.vaiTro, 'student.write') });
  const relationMap = new Map(readObjectsNoCache_(SHEET_HOCSINH_KYHOC).filter(row => String(row.MaKyHoc || '').trim() === selected && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED').map(row => [String(row.MaHocSinh || '').trim(), row]));
  const context = tuitionGradeContext_();
  const students = readObjectsNoCache_(SHEET_HOCSINH).filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED' && relationMap.has(String(row.MaHocSinh || '').trim()))
    .map(row => { const relation = relationMap.get(String(row.MaHocSinh).trim()); const state = studentTuitionState_(relation, row, term, context); return { maHocSinh: String(row.MaHocSinh).trim(), hoTen: String(row.HoTen || ''), khoi: String(row.Khoi || ''), lop: String(row.Lop || ''),
      mode: tuitionMode_(relation), amount: state.amount, capHoc: state.grade === null ? null : (state.grade <= 5 ? 1 : 2), issue: state.issue }; });
  return jsonResponse_({ terms: terms, maKyHoc: selected, config: termTuitionConfig_(term), students: students, unresolvedCount: students.filter(student => student.issue).length, canWrite: SecurityService.hasPermission(session.vaiTro, 'student.write') });
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
    const context = tuitionGradeContext_();
    const ids = new Set(students.map(row => String(row.MaHocSinh || '').trim()));
    const relations = readObjectsNoCache_(SHEET_HOCSINH_KYHOC).map(row => Object.assign({}, row));
    const seen = new Set();
    data.overrides.forEach(item => {
      if (!item || !ids.has(item.maHocSinh) || seen.has(item.maHocSinh) || !['AUTO','CUSTOM'].includes(item.mode)) throw new Error('Học sinh hoặc chế độ học phí không hợp lệ.');
      seen.add(item.maHocSinh);
      const links = relations.filter(row => String(row.MaHocSinh || '').trim() === item.maHocSinh && String(row.MaKyHoc || '').trim() === String(term.MaKyHoc).trim() && String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED');
      if (!links.length) throw new Error('Học sinh không còn thuộc kỳ học này. Vui lòng tải lại.');
      if (item.mode === 'AUTO') {
        const student = students.find(row => String(row.MaHocSinh || '').trim() === item.maHocSinh);
        if (resolveTuitionGrade_(student, context) === null) throw new Error('Chưa xác định cấp của ' + (student.HoTen || item.maHocSinh) + '. Hãy sửa khối/lớp hoặc giữ mức riêng.');
      }
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

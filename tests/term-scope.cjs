// Local regression tests. No Google account, network or Sheet writes.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
process.chdir(path.join(__dirname, '..'));
const cache = new Map();
let serial = 0;
const rows = {
  KyHoc: [
    { MaKyHoc: 'A', TenKyHoc: 'Kỳ A', TrangThai: 'ACTIVE' },
    { MaKyHoc: 'B', TenKyHoc: 'Kỳ B', TrangThai: 'INACTIVE' },
    { MaKyHoc: 'D', TenKyHoc: 'Đã xóa', TrangThai: 'DELETED' }
  ],
  NguoiDung: [{ MaNguoiDung: 'u1', TenDangNhap: 'teacher', HoTen: 'Teacher', VaiTro: 'TEACHER', TrangThai: 'ACTIVE' }],
  HocSinhKyHoc: [{ MaHocSinh: 's1', MaKyHoc: 'A', TrangThai: 'ACTIVE', HocPhi: 1200000, GhiChuHocPhi: 'Giữ ghi chú' }, { MaHocSinh: 's2', MaKyHoc: 'B', TrangThai: 'ACTIVE' }],
  HocSinh: [{ MaHocSinh: 's1', HoTen: 'Student A', MaKyHoc: 'A', Khoi: '3' }, { MaHocSinh: 's2', HoTen: 'Student B', MaKyHoc: 'B', Khoi: '7' }]
};
const ctx = {
  console,
  Utilities: {
    getUuid: () => 'test-token-' + (++serial),
    DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (algorithm, value) => Array.from(crypto.createHash('sha256').update(value).digest()),
    formatDate: (date, timeZone, format) => {
      const parts = new Intl.DateTimeFormat('en-CA', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(date);
      const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
      return format.replace('yyyy', values.year).replace('MM', values.month).replace('dd', values.day);
    }
  },
  Session: { getScriptTimeZone: () => 'Asia/Ho_Chi_Minh' },
  CacheService: { getScriptCache: () => ({ get: key => cache.get(key) || null, put: (key, value) => cache.set(key, value), remove: key => cache.delete(key) }) },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({}) }
};
vm.createContext(ctx);
// Apps Script names are shared by server scripts and HTML templates.
const scriptNames = new Map();
for (const file of fs.readdirSync('.').filter(file => /\.(js|gs|html)$/.test(file))) {
  const name = file.replace(/\.[^.]+$/, '');
  assert.ok(!scriptNames.has(name), 'Duplicate Apps Script name: ' + name + ' (' + scriptNames.get(name) + ', ' + file + ')');
  scriptNames.set(name, file);
}
let syntaxBlocks = 0;
for (const file of fs.readdirSync('.').filter(file => /\.(js|html)$/.test(file))) {
  const source = fs.readFileSync(file, 'utf8');
  if (file.endsWith('.js')) { vm.runInContext(source, ctx, { filename: file }); syntaxBlocks++; }
  else for (const match of source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) {
    new vm.Script(match[1].replace(/<\?[\s\S]*?\?>/g, 'null'), { filename: file }); syntaxBlocks++;
  }
}
const feeFormatter = fs.readFileSync('QuanLyHocSinh.html','utf8').match(/function formatMoneyInputValue\(value\) \{[\s\S]*?\n    \}/)[0];
assert.equal(vm.runInNewContext(feeFormatter + ';formatMoneyInputValue(0)'), '0', 'Student editor preserves a zero rate');
const baseline = JSON.parse(ctx.runArchitectureUnitTests());
assert.equal(baseline.failed, 0, JSON.stringify(baseline.results.filter(item => !item.passed)));
Object.assign(ctx, {
  readObjects_: name => rows[name] || [], readObjectsNoCache_: name => rows[name] || [],
  ensureSheet_: () => ({}), cacheGetString_: () => null, cachePutString_: () => {},
  buildCacheKey_: value => value,
  safeWriteAuditLog_: () => {}
});
const options = JSON.parse(ctx.getKyHocList());
assert.equal(options[0].maKyHoc, '__ALL_TERMS__');
assert.deepEqual(options.slice(1).map(item => item.maKyHoc), ['A']);
assert.deepEqual(Array.from(ctx.getAllTerms_(), item => item.maKyHoc), ['A', 'B']);
function parentSession() {
  cache.set('LOGIN_TOKEN_parent', JSON.stringify({ maKyHoc: '__ALL_TERMS__', maNguoiDung: 'u1' }));
}
parentSession();
assert.equal(ctx.SecurityService.getSession('parent').valid, true);
const a = ctx.SecurityService.createTermSession('parent', 'A');
const b = ctx.SecurityService.createTermSession('parent', 'B');
assert.equal(ctx.SecurityService.getSession(a).maKyHoc, 'A');
assert.equal(ctx.SecurityService.getSession(b).maKyHoc, 'B');
assert.equal(ctx.SecurityService.getSession(a).vaiTro, 'TEACHER');
assert.throws(() => ctx.SecurityService.requireSession(a, 'finance.write'));
assert.throws(() => ctx.SecurityService.requireSession('parent', 'attendance.write'));
assert.throws(() => ctx.SecurityService.createTermSession(a, 'B'));
assert.throws(() => ctx.SecurityService.createTermSession('parent', 'D'));
assert.deepEqual(JSON.parse(ctx.getHocSinhList(a, {})).map(item => item.maHocSinh), ['s1']);
assert.deepEqual(JSON.parse(ctx.getHocSinhList(b, {})).map(item => item.maHocSinh), ['s2']);
// The matrix reads all students, but only writers can alter membership.
assert.equal(JSON.parse(ctx.getStudentTermMatrix('parent')).students.length, 2);
assert.equal(JSON.parse(ctx.getStudentTermMatrix(a)).canWrite, false);
assert.throws(() => ctx.saveStudentTermMatrix('parent', []));
rows.NguoiDung[0].VaiTro = 'ADMIN';
let released = 0, writes = 0;
ctx.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => released++ }) };
ctx.writeObjectsToSheet_ = (name, value) => { rows[name] = value; writes++; };
ctx.bumpDataVersion_ = () => {};
ctx.saveStudentTermMatrix('parent', [{ maHocSinh: 's1', maKyHoc: 'A', selected: false, expected: true }]);
assert.equal(JSON.parse(ctx.getHocSinhList(a, {})).length, 0, 'Legacy MaKyHoc must not restore an unchecked student');
assert.equal(ctx.getHocSinhTheoKyHocForThuPhi_('A').length, 0, 'Empty membership must not include unassigned students');
assert.deepEqual(JSON.parse(ctx.getHocSinhList(b, {})).map(item => item.maHocSinh), ['s2']);
ctx.saveStudentTermMatrix('parent', [{ maHocSinh: 's1', maKyHoc: 'A', selected: true, expected: false }]);
assert.equal(rows.HocSinhKyHoc[0].HocPhi, 1200000);
assert.equal(rows.HocSinhKyHoc[0].GhiChuHocPhi, 'Giữ ghi chú');
assert.equal(rows.HocSinhKyHoc.length, 2, 'Reactivation must not duplicate links');
const beforeInvalid = JSON.stringify(rows.HocSinhKyHoc);
assert.throws(() => ctx.saveStudentTermMatrix('parent', [
  { maHocSinh: 's1', maKyHoc: 'A', selected: false, expected: true },
  { maHocSinh: 's1', maKyHoc: 'D', selected: true, expected: false }
]));
assert.equal(JSON.stringify(rows.HocSinhKyHoc), beforeInvalid, 'Validate complete batch before writing');
assert.equal(writes, 2); assert.equal(released, 3, 'Always release the lock');
rows.HocSinh.push({ MaHocSinh: 's3', HoTen: 'Unassigned', MaKyHoc: 'A', Khoi: '5' });
assert.deepEqual(JSON.parse(ctx.getHocSinhList(a, {})).map(item => item.maHocSinh), ['s1']);
assert.deepEqual(JSON.parse(ctx.getStudentTermMatrix(a)).students.find(item => item.maHocSinh === 's3').kyHocIds, []);
console.log('Matrix tests passed: read/write roles, check/uncheck, strict membership, fee preservation, validation and locking.');
// Tuition is stored per term, with explicit automatic or individual mode.
rows.HocSinhKyHoc.push({ MaHocSinh: 's2', MaKyHoc: 'A', TrangThai: 'ACTIVE', HocPhiMode: 'AUTO', HocPhi: '' });
rows.HocSinhKyHoc.push({ MaHocSinh: 's1', MaKyHoc: 'B', TrangThai: 'ACTIVE', HocPhiMode: 'AUTO', HocPhi: '' });
ctx.saveTermTuition('parent', { maKyHoc: 'A', cap1: 100, cap2: 200, overrides: [] });
const link = (id, term) => rows.HocSinhKyHoc.find(r => r.MaHocSinh === id && r.MaKyHoc === term && r.TrangThai !== 'DELETED');
assert.equal(link('s1','A').HocPhi, 1200000, 'Keep legacy individual fees');
assert.equal(link('s2','A').HocPhi, 200, 'Secondary school rate');
ctx.saveTermTuition('parent', { maKyHoc: 'B', cap1: 300, cap2: 400, overrides: [] });
assert.equal(link('s1','B').HocPhi, 300, 'Primary school rate in another term');
assert.equal(link('s2','A').HocPhi, 200, 'Other terms unchanged');
ctx.saveTermTuition('parent', { maKyHoc: 'A', cap1: 111, cap2: 222, overrides: [{ maHocSinh:'s2', mode:'CUSTOM', amount:0 }] });
assert.equal(link('s2','A').HocPhi, 0);
ctx.saveTermTuition('parent', { maKyHoc: 'A', cap1: 111, cap2: 999, overrides: [] });
assert.equal(link('s2','A').HocPhi, 0, 'Zero custom fee survives default rate changes');
assert.equal(ctx.getHocSinhTheoKyHocForThuPhi_('A').find(s => s.maHocSinh === 's2').hocPhi, 0, 'Monthly roster must preserve zero');
ctx.saveTermTuition('parent', { maKyHoc: 'A', cap1: 111, cap2: 222, overrides: [{ maHocSinh:'s2', mode:'AUTO' }] });
assert.equal(link('s2','A').HocPhi, 222, 'Return to automatic rate');
ctx.saveHocSinhKyHoc_('s1', ['A','B'], 333, { maKyHoc:'B', khoi:'3', mode:'CUSTOM' });
assert.equal(link('s1','A').HocPhi,1200000,'Editing one term must preserve other term fees');
assert.equal(link('s1','B').HocPhi,333);
assert.throws(() => ctx.saveTermTuition('parent', {maKyHoc:'A',cap1:-1,cap2:200,overrides:[]}));
assert.throws(() => ctx.tuitionAmount_(Infinity));
assert.throws(() => ctx.tuitionAmount_(''));
assert.equal(ctx.tuitionForGrade_(5,{cap1:10,cap2:20}),10);
assert.equal(ctx.tuitionForGrade_(6,{cap1:10,cap2:20}),20);
// Quick add validates class, creates automatic fees in each selected term, and is retry-safe.
rows.Khoi=[{Khoi:'3',TenKhoi:'Khối 3',TrangThai:'ACTIVE',ThuTu:3}];
rows.Lop=[{MaLop:'L3',TenLop:'Lớp 3',Khoi:'3',TrangThai:'ACTIVE',ThuTu:3}];
ctx.ensureSheet_=(ss,name)=>({name});
ctx.appendObjectsToSheet_=(sheet,items)=>{if(!rows[sheet.name])rows[sheet.name]=[];rows[sheet.name].push(...items);};
const quick={requestId:'quick_test_001',hoTen:'New student',khoi:'3',lop:'L3',sdtPhuHuynh:'',kyHocIds:['A','B']};
const added=JSON.parse(ctx.quickAddStudent('parent',quick));
assert.equal(link(added.student.maHocSinh,'A').HocPhi,111);
assert.equal(link(added.student.maHocSinh,'B').HocPhi,300);
assert.equal(link(added.student.maHocSinh,'B').HocPhiMode,'AUTO');
const count=rows.HocSinh.length;
assert.equal(JSON.parse(ctx.quickAddStudent('parent',quick)).student.maHocSinh,added.student.maHocSinh);
assert.equal(rows.HocSinh.length,count,'Retry must not add a duplicate student');
assert.throws(()=>ctx.quickAddStudent('parent',{...quick,requestId:'quick_test_002',lop:'WRONG'}));
assert.equal(rows.HocSinh.length,count);
const unassigned=JSON.parse(ctx.quickAddStudent('parent',{...quick,requestId:'quick_test_003',kyHocIds:[]}));
assert.equal(unassigned.student.kyHocIds.length,0);
console.log('Tuition and quick-add tests passed: tier boundaries, term isolation, legacy/custom/zero rates, automatic restoration, validation and retry safety.');
// One unresolved grade must not take down either the tuition page or student editor.
const badRows=[
  {MaHocSinh:'missing',HoTen:'Missing grade',Khoi:'',Lop:''},
  {MaHocSinh:'label',HoTen:'Grade label',Khoi:'Khối 3',Lop:''},
  {MaHocSinh:'unsupported',HoTen:'Grade 10',Khoi:'10',Lop:'L3'},
  {MaHocSinh:'classgrade',HoTen:'Class mapping',Khoi:'',Lop:'L3'}
];
rows.HocSinh.push(...badRows);
rows.HocSinhKyHoc.push(...badRows.map(row=>({MaHocSinh:row.MaHocSinh,MaKyHoc:'A',HocPhiMode:'AUTO',HocPhi:row.MaHocSinh==='missing'?777:'',TrangThai:'ACTIVE'})));
const tuitionPage=JSON.parse(ctx.getTermTuitionData('parent','A'));
assert.equal(tuitionPage.maKyHoc,'A');
assert.equal(tuitionPage.students.find(s=>s.maHocSinh==='missing').capHoc,null);
assert.equal(tuitionPage.students.find(s=>s.maHocSinh==='missing').amount,777);
assert.equal(tuitionPage.students.find(s=>s.maHocSinh==='unsupported').amount,null,'Unknown is not zero');
assert.equal(tuitionPage.students.find(s=>s.maHocSinh==='unsupported').capHoc,null,'Do not override invalid explicit grade using class');
assert.equal(tuitionPage.students.find(s=>s.maHocSinh==='label').capHoc,1);
assert.equal(tuitionPage.students.find(s=>s.maHocSinh==='classgrade').capHoc,1);
assert.equal(tuitionPage.unresolvedCount,2);
assert.ok(JSON.parse(ctx.getHocSinhList(a,{})).find(s=>s.maHocSinh==='missing'),'Student can still be loaded for repair');
ctx.saveTermTuition('parent',{maKyHoc:'A',cap1:555,cap2:666,overrides:[]});
assert.equal(link('missing','A').HocPhi,777,'Keep stored amount when grade cannot be resolved');
assert.equal(link('unsupported','A').HocPhi,'','Do not invent a fee');
assert.equal(link('label','A').HocPhi,555);
assert.throws(()=>ctx.saveTermTuition('parent',{maKyHoc:'A',cap1:555,cap2:666,overrides:[{maHocSinh:'missing',mode:'AUTO'}]}),/Missing grade/);
ctx.saveTermTuition('parent',{maKyHoc:'A',cap1:555,cap2:666,overrides:[{maHocSinh:'missing',mode:'CUSTOM',amount:0}]});
assert.equal(link('missing','A').HocPhi,0,'Explicit custom rate remains available without grade');
for(const value of ['',null,'10','3A','Lớp 3A','-1']) assert.equal(ctx.normalizeTuitionGrade_(value),null);
assert.equal(ctx.normalizeTuitionGrade_('Khối 9'),9);
console.log('Invalid-grade regression passed: page availability, explicit normalization, class lookup, warnings and no guessed tuition.');
ctx.logout('parent');
assert.equal(ctx.SecurityService.getSession(a).valid, false);
assert.equal(ctx.SecurityService.getSession(b).valid, false);
parentSession();
ctx.logout(a);
assert.equal(ctx.SecurityService.getSession('parent').valid, false);
assert.equal(ctx.SecurityService.getSession(b).valid, false);
console.log(`${baseline.passed}/${baseline.total} architecture tests passed; ${syntaxBlocks} syntax blocks parsed.`);
console.log('Term scope checks passed: options, inactive/deleted terms, scoped sessions, permissions, student isolation, and parent/child logout.');

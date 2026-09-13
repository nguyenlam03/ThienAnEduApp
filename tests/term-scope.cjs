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
  HocSinh: [{ MaHocSinh: 's1', HoTen: 'Student A', MaKyHoc: 'A' }, { MaHocSinh: 's2', HoTen: 'Student B', MaKyHoc: 'B' }]
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
let syntaxBlocks = 0;
for (const file of fs.readdirSync('.').filter(file => /\.(js|html)$/.test(file))) {
  const source = fs.readFileSync(file, 'utf8');
  if (file.endsWith('.js')) { vm.runInContext(source, ctx, { filename: file }); syntaxBlocks++; }
  else for (const match of source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) {
    new vm.Script(match[1].replace(/<\?[\s\S]*?\?>/g, 'null'), { filename: file }); syntaxBlocks++;
  }
}
const baseline = JSON.parse(ctx.runArchitectureUnitTests());
assert.equal(baseline.failed, 0, JSON.stringify(baseline.results.filter(item => !item.passed)));
Object.assign(ctx, {
  readObjects_: name => rows[name] || [], readObjectsNoCache_: name => rows[name] || [],
  ensureSheet_: () => ({}), cacheGetString_: () => null, cachePutString_: () => {},
  buildCacheKey_: value => value, getHocSinhKyHocMap_: () => ({}), hasMappingForKyHoc_: () => true,
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
ctx.logout('parent');
assert.equal(ctx.SecurityService.getSession(a).valid, false);
assert.equal(ctx.SecurityService.getSession(b).valid, false);
parentSession();
ctx.logout(a);
assert.equal(ctx.SecurityService.getSession('parent').valid, false);
assert.equal(ctx.SecurityService.getSession(b).valid, false);
console.log(`${baseline.passed}/${baseline.total} architecture tests passed; ${syntaxBlocks} syntax blocks parsed.`);
console.log('Term scope checks passed: options, inactive/deleted terms, scoped sessions, permissions, student isolation, and parent/child logout.');

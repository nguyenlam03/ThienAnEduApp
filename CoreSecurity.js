/**
 * Security Service
 * Quản lý người dùng, mật khẩu băm, phiên đăng nhập và phân quyền.
 * Controller/UI không được tự kiểm tra vai trò; mọi quyền đi qua service này.
 */
var SecurityService = (function () {
  var USER_SHEET = 'NguoiDung';
  var PASSWORD_HASH_PROPERTY = 'APP_OWNER_PASSWORD_SHA256';
  var LEGACY_OWNER_PASSWORD_HASH = 'f77219d0a97a8606f6a57c1eb11deb68702fe680621a81dee7ce28698de19cb2';
  var SESSION_SECONDS = 21600;
  var LOGIN_ATTEMPT_PREFIX = 'LOGIN_ATTEMPT_';
  var MAX_LOGIN_ATTEMPTS = 5;
  var LOGIN_LOCK_SECONDS = 900;
  var ALLOWED_ROLES = ['OWNER', 'ADMIN', 'FINANCE', 'TEACHER', 'VIEWER'];
  var ROLE_PERMISSIONS = {
    OWNER: ['*'],
    ADMIN: ['system.admin', 'student.read', 'student.write', 'attendance.read', 'attendance.write', 'tuition.read', 'tuition.write', 'finance.read', 'finance.write', 'finance.close'],
    FINANCE: ['student.read', 'tuition.read', 'tuition.write', 'finance.read', 'finance.write'],
    TEACHER: ['student.read', 'attendance.read', 'attendance.write'],
    VIEWER: ['student.read', 'attendance.read', 'tuition.read', 'finance.read']
  };

  function headers() {
    return [
      'MaNguoiDung', 'TenDangNhap', 'HoTen', 'VaiTro', 'MatKhauHash',
      'TrangThai', 'LastLoginAt', 'CreatedAt', 'UpdatedAt'
    ];
  }

  function hashPassword(password) {
    return bytesToHex(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(password || ''),
      Utilities.Charset.UTF_8
    ));
  }

  function bytesToHex(bytes) {
    return bytes.map(function (byte) {
      var value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    }).join('');
  }

  function derivePassword(password, salt, iterations) {
    var value = String(salt) + ':' + String(password || '');
    for (var i = 0; i < iterations; i++) {
      value = bytesToHex(Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        value,
        Utilities.Charset.UTF_8
      ));
    }
    return value;
  }

  function createPasswordHash(password) {
    var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    return ['v3', salt, hashPassword(salt + ':' + String(password || ''))].join('$');
  }

  function verifyPassword(password, storedHash) {
    storedHash = String(storedHash || '');
    var parts = storedHash.split('$');
    if (parts.length === 3 && parts[0] === 'v3') {
      return safeEqual(hashPassword(parts[1] + ':' + String(password || '')), parts[2]);
    }
    if (parts.length === 4 && parts[0] === 'v2') {
      var iterations = Number(parts[1]);
      if (!iterations || iterations > 20000) return false;
      return safeEqual(derivePassword(password, parts[2], iterations), parts[3]);
    }
    return safeEqual(hashPassword(password), storedHash);
  }

  function safeEqual(left, right) {
    left = String(left || '');
    right = String(right || '');
    var mismatch = left.length ^ right.length;
    var length = Math.max(left.length, right.length);
    for (var i = 0; i < length; i++) {
      mismatch |= (left.charCodeAt(i % Math.max(1, left.length)) || 0) ^
        (right.charCodeAt(i % Math.max(1, right.length)) || 0);
    }
    return mismatch === 0;
  }

  function ensureSchema() {
    var sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), USER_SHEET, headers());
    var users = readObjectsNoCache_(USER_SHEET).filter(function (row) {
      return String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED';
    });
    if (users.length) return sheet;

    var props = PropertiesService.getScriptProperties();
    var initialHash = props.getProperty(PASSWORD_HASH_PROPERTY) || LEGACY_OWNER_PASSWORD_HASH;
    props.setProperty(PASSWORD_HASH_PROPERTY, initialHash);
    var now = new Date();
    appendObjectsToSheet_(sheet, [{
      MaNguoiDung: 'USER_OWNER',
      TenDangNhap: 'admin',
      HoTen: 'Chủ cơ sở',
      VaiTro: 'OWNER',
      MatKhauHash: initialHash,
      TrangThai: 'ACTIVE',
      LastLoginAt: '',
      CreatedAt: now,
      UpdatedAt: now
    }], headers());
    return sheet;
  }

  function getUsers() {
    ensureSchema();
    return readObjectsNoCache_(USER_SHEET).map(function (row) {
      return {
        maNguoiDung: String(row.MaNguoiDung || '').trim(),
        tenDangNhap: String(row.TenDangNhap || '').trim().toLowerCase(),
        hoTen: String(row.HoTen || '').trim(),
        vaiTro: String(row.VaiTro || 'VIEWER').trim().toUpperCase(),
        matKhauHash: String(row.MatKhauHash || '').trim(),
        trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase(),
        lastLoginAt: row.LastLoginAt || ''
      };
    }).filter(function (user) {
      return user.maNguoiDung && user.tenDangNhap && user.trangThai !== 'DELETED';
    });
  }

  function authenticate(username, password, maKyHoc) {
    username = String(username || 'admin').trim().toLowerCase();
    var attemptKey = LOGIN_ATTEMPT_PREFIX + username;
    var cache = CacheService.getScriptCache();
    var attempts = Number(cache.get(attemptKey) || 0);
    if (attempts >= MAX_LOGIN_ATTEMPTS) {
      return { success: false, message: 'Tài khoản tạm khóa 15 phút do đăng nhập sai nhiều lần.' };
    }
    var user = getUsers().find(function (item) {
      return item.tenDangNhap === username && item.trangThai === 'ACTIVE';
    });
    var ownerPropertyHash = user && user.vaiTro === 'OWNER'
      ? String(PropertiesService.getScriptProperties().getProperty(PASSWORD_HASH_PROPERTY) || '')
      : '';
    var ownerLegacyMatch = !!user && user.vaiTro === 'OWNER' && ownerPropertyHash.indexOf('v2$') !== 0 &&
      ownerPropertyHash.indexOf('v3$') !== 0 && safeEqual(hashPassword(password), ownerPropertyHash);
    if (!user || (!ownerLegacyMatch && !verifyPassword(password, user.matKhauHash))) {
      cache.put(attemptKey, String(attempts + 1), LOGIN_LOCK_SECONDS);
      return { success: false, message: 'Tên đăng nhập hoặc mật khẩu không đúng.' };
    }

    var term = getKyHocArray_().find(function (item) { return item.maKyHoc === maKyHoc; });
    if (!term) return { success: false, message: 'Kỳ học không hợp lệ hoặc chưa được kích hoạt.' };

    var token = createSession(user, term);
    cache.remove(attemptKey);
    var loginUpdates = { LastLoginAt: new Date() };
    if (user.matKhauHash.indexOf('v3$') !== 0) {
      loginUpdates.MatKhauHash = createPasswordHash(password);
      if (user.vaiTro === 'OWNER') {
        PropertiesService.getScriptProperties().setProperty(PASSWORD_HASH_PROPERTY, loginUpdates.MatKhauHash);
      }
    }
    try { updateUserRow(user.maNguoiDung, loginUpdates); } catch (error) { console.warn(error); }
    return {
      success: true,
      token: token,
      user: stripPrivateUser(user),
      maKyHoc: term.maKyHoc,
      tenKyHoc: term.tenKyHoc
    };
  }

  function createSession(user, term) {
    var token = Utilities.getUuid();
    var payload = {
      maKyHoc: term.maKyHoc,
      tenKyHoc: term.tenKyHoc,
      maNguoiDung: user.maNguoiDung,
      tenDangNhap: user.tenDangNhap,
      hoTen: user.hoTen,
      vaiTro: user.vaiTro,
      issuedAt: Date.now()
    };
    CacheService.getScriptCache().put(CACHE_LOGIN_PREFIX + token, JSON.stringify(payload), SESSION_SECONDS);
    return token;
  }

  function createLegacyOwnerSession(term) {
    var owner = getUsers().find(function (item) {
      return item.vaiTro === 'OWNER' && item.trangThai === 'ACTIVE';
    });
    if (!owner) throw new Error('Chưa cấu hình tài khoản chủ cơ sở.');
    return createSession(owner, term);
  }

  function getSession(token) {
    if (!token) return { valid: false };
    var cached = CacheService.getScriptCache().get(CACHE_LOGIN_PREFIX + token);
    if (!cached) return { valid: false };

    var payload = null;
    try { payload = JSON.parse(cached); } catch (error) {
      payload = {
        maKyHoc: String(cached || '').trim(),
        maNguoiDung: 'USER_OWNER', tenDangNhap: 'admin', hoTen: 'Chủ cơ sở', vaiTro: 'OWNER'
      };
    }
    var term = getKyHocArray_().find(function (item) { return item.maKyHoc === payload.maKyHoc; });
    if (!term) return { valid: false };

    var user = getUsers().find(function (item) {
      return item.maNguoiDung === payload.maNguoiDung && item.trangThai === 'ACTIVE';
    });
    if (!user && payload.maNguoiDung === 'USER_OWNER') {
      user = { maNguoiDung: 'USER_OWNER', tenDangNhap: 'admin', hoTen: 'Chủ cơ sở', vaiTro: 'OWNER', trangThai: 'ACTIVE' };
    }
    if (!user) return { valid: false };
    return {
      valid: true,
      maKyHoc: term.maKyHoc,
      tenKyHoc: term.tenKyHoc,
      maNguoiDung: user.maNguoiDung,
      tenDangNhap: user.tenDangNhap,
      hoTen: user.hoTen,
      vaiTro: user.vaiTro
    };
  }

  function hasPermission(role, permission) {
    if (!permission) return true;
    var permissions = ROLE_PERMISSIONS[String(role || 'VIEWER').toUpperCase()] || [];
    return permissions.indexOf('*') !== -1 || permissions.indexOf(permission) !== -1;
  }

  function requireSession(token, permission) {
    var session = getSession(token);
    if (!session.valid) throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
    if (!hasPermission(session.vaiTro, permission)) {
      throw new Error('Tài khoản không có quyền thực hiện thao tác này.');
    }
    return session;
  }

  function changePassword(token, currentPassword, newPassword) {
    var session = requireSession(token);
    newPassword = String(newPassword || '');
    if (newPassword.length < 8) throw new Error('Mật khẩu mới phải có ít nhất 8 ký tự.');
    var user = getUsers().find(function (item) { return item.maNguoiDung === session.maNguoiDung; });
    if (!user || !verifyPassword(currentPassword, user.matKhauHash)) {
      throw new Error('Mật khẩu hiện tại không đúng.');
    }
    var newStoredHash = createPasswordHash(newPassword);
    updateUserRow(user.maNguoiDung, { MatKhauHash: newStoredHash });
    if (user.vaiTro === 'OWNER') {
      PropertiesService.getScriptProperties().setProperty(PASSWORD_HASH_PROPERTY, newStoredHash);
    }
    safeWriteAuditLog_(session, 'CHANGE_PASSWORD', 'NGUOI_DUNG', user.maNguoiDung, null, null);
    CacheService.getScriptCache().remove(CACHE_LOGIN_PREFIX + token);
    return { success: true, message: 'Đã đổi mật khẩu. Vui lòng đăng nhập lại.' };
  }

  function saveUser(token, data) {
    var session = requireSession(token, 'system.admin');
    data = data || {};
    var id = String(data.maNguoiDung || '').trim() || ('USER_' + Utilities.getUuid().slice(0, 10).toUpperCase());
    var username = String(data.tenDangNhap || '').trim().toLowerCase();
    var fullName = String(data.hoTen || '').trim();
    var role = String(data.vaiTro || 'VIEWER').trim().toUpperCase();
    if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw new Error('Tên đăng nhập chỉ gồm chữ thường, số, dấu chấm, gạch dưới hoặc gạch ngang.');
    if (!fullName) throw new Error('Vui lòng nhập họ tên người dùng.');
    if (ALLOWED_ROLES.indexOf(role) === -1) throw new Error('Vai trò không hợp lệ.');
    var users = getUsers();
    var existing = users.find(function (item) { return item.maNguoiDung === id; });
    if ((role === 'OWNER' || (existing && existing.vaiTro === 'OWNER')) && session.vaiTro !== 'OWNER') {
      throw new Error('Chỉ chủ cơ sở mới được thay đổi tài khoản chủ cơ sở.');
    }
    if (id === session.maNguoiDung && String(data.trangThai || 'ACTIVE').trim().toUpperCase() === 'INACTIVE') {
      throw new Error('Không thể tự khóa tài khoản đang đăng nhập.');
    }
    if (users.some(function (item) { return item.maNguoiDung !== id && item.tenDangNhap === username; })) {
      throw new Error('Tên đăng nhập đã tồn tại.');
    }
    var password = String(data.matKhau || '');
    if (!existing && password.length < 8) throw new Error('Người dùng mới phải có mật khẩu ít nhất 8 ký tự.');
    var values = {
      TenDangNhap: username, HoTen: fullName, VaiTro: role,
      TrangThai: String(data.trangThai || 'ACTIVE').trim().toUpperCase() === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE'
    };
    if (password) {
      if (password.length < 8) throw new Error('Mật khẩu phải có ít nhất 8 ký tự.');
      values.MatKhauHash = createPasswordHash(password);
    }
    updateUserRow(id, values);
    safeWriteAuditLog_(session, existing ? 'UPDATE' : 'CREATE', 'NGUOI_DUNG', id, null, {
      tenDangNhap: username, hoTen: fullName, vaiTro: role, trangThai: values.TrangThai
    });
    return { success: true, maNguoiDung: id, message: 'Đã lưu người dùng.' };
  }

  function updateOrInsertUser(id, values) {
    var sheet = ensureSchema();
    var rows = sheet.getDataRange().getValues();
    var index = buildHeaderIndex_(rows[0].map(function (item) { return String(item || '').trim(); }));
    var target = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][index.MaNguoiDung] || '').trim() === id) { target = i; break; }
    }
    var now = new Date();
    var row = target >= 0 ? rows[target].slice() : new Array(rows[0].length).fill('');
    row[index.MaNguoiDung] = id;
    Object.keys(values).forEach(function (key) { if (index[key] !== undefined) row[index[key]] = values[key]; });
    row[index.CreatedAt] = target >= 0 ? (row[index.CreatedAt] || now) : now;
    row[index.UpdatedAt] = now;
    if (target >= 0) sheet.getRange(target + 1, 1, 1, row.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  }

  function updateUserRow(id, values) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật người dùng.');
    try { updateOrInsertUser(id, values); } finally { lock.releaseLock(); }
  }

  function touchLastLogin(id) {
    try { updateUserRow(id, { LastLoginAt: new Date() }); } catch (error) { console.warn(error); }
  }

  function stripPrivateUser(user) {
    return {
      maNguoiDung: user.maNguoiDung,
      tenDangNhap: user.tenDangNhap,
      hoTen: user.hoTen,
      vaiTro: user.vaiTro,
      trangThai: user.trangThai,
      lastLoginAt: user.lastLoginAt
    };
  }

  return {
    ensureSchema: ensureSchema,
    authenticate: authenticate,
    getSession: getSession,
    createLegacyOwnerSession: createLegacyOwnerSession,
    requireSession: requireSession,
    hasPermission: hasPermission,
    changePassword: changePassword,
    saveUser: saveUser,
    getUsers: function (token) {
      requireSession(token, 'system.admin');
      return getUsers().map(stripPrivateUser);
    },
    hashPassword: hashPassword
  };
})();

function getNguoiDungList(token) {
  return jsonResponse_(SecurityService.getUsers(token));
}

function saveNguoiDung(token, data) {
  return jsonResponse_(SecurityService.saveUser(token, data));
}

function changeMyPassword(token, currentPassword, newPassword) {
  return jsonResponse_(SecurityService.changePassword(token, currentPassword, newPassword));
}

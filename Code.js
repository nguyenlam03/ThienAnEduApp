const SHEET_KYHOC = 'KyHoc';
const SHEET_KHOI = 'Khoi';
const SHEET_LOP = 'Lop';
const SHEET_HOCSINH = 'HocSinh';
const SHEET_HOCSINH_KYHOC = 'HocSinhKyHoc';
const SHEET_THUPHI = 'ThuPhi';
const SHEET_DANHMUC_THUCHI = 'DanhMucThuChi';
const SHEET_SOTHUCHI = 'SoThuChi';
const SHEET_NGUONTIEN = 'NguonTien';
const SHEET_DOITUONG_THUCHI = 'DoiTuongThuChi';
const SHEET_KEHOACH_CHI_THANG = 'KeHoachChiThang';
const SHEET_CAUHINH_TAICHINH_THANG = 'CauHinhTaiChinhThang';
const SHEET_NHANSU_TAICHINH = 'NhanSuTaiChinh';
const SHEET_KHOANCHI_DINHKY = 'KhoanChiDinhKy';
const SHEET_DANHMUC_HU_TAICHINH = 'DanhMucHuTaiChinh';
const SHEET_HU_TAICHINH_THANG = 'HuTaiChinhThang';
const SHEET_CHOT_PHANBO_HU = 'ChotPhanBoHu';
const SHEET_DANHMUC_GIADINH = 'DanhMucTaiChinhGiaDinh';
const SHEET_GIAODICH_GIADINH = 'GiaoDichTaiChinhGiaDinh';
const SHEET_CAUHINH_GIADINH_THANG = 'CauHinhGiaDinhThang';
const SHEET_DIEMDANH = 'DiemDanh';
const SHEET_CAUHINH = 'CauHinh';
const SHEET_DANHMUC_KHOANTHU_PHI = 'DanhMucKhoanThuPhi';
const CONFIG_NOTIFICATION_DURATION = 'NOTIFICATION_DURATION_SECONDS';
const CONFIG_BRAND_NAME = 'APP_BRAND_NAME';
const CONFIG_BRAND_TAGLINE = 'APP_BRAND_TAGLINE';
const CONFIG_BRAND_LOGO_URL = 'APP_BRAND_LOGO_URL';

const CACHE_LOGIN_PREFIX = 'LOGIN_TOKEN_';
const CACHE_PREFIX = 'TA_CACHE_';
const CACHE_SECONDS = 1800;
const CACHE_CHUNK_SIZE = 90000;
const DATA_CACHE_ENABLED_PROPERTY = 'DATA_CACHE_ENABLED';
let DATA_VERSION_MEMO_ = '';
let DATA_CACHE_ENABLED_MEMO_ = null;

/**
 * Include file HTML dùng chung.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Mã hóa dữ liệu trước khi chèn vào một thẻ script của HTML template.
 * JSON.stringify đơn thuần vẫn có thể chứa chuỗi đóng thẻ hoặc ký tự phân
 * cách dòng khiến trình duyệt dừng phân tích toàn bộ JavaScript của trang.
 */
function inlineScriptJson_(value) {
  const json = JSON.stringify(value);
  return String(json === undefined ? 'null' : json)
    .replace(/&/g, '\\u0026')
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

/**
 * Điều hướng trang.
 */
function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : 'Login';
  const notificationDurationMs = getNotificationDurationMs_();
  const brand = getAppBrandConfig_();

  const protectedPages = [
    'Index',
    'QuanLyHocSinh',
    'QuanLyDiemDanh',
    'QuanLyThuPhu',
    'QuanLyThuChi',
    'QuanLyTaiChinh',
    'KeHoachTaiChinh',
    'TaiChinhGiaDinh'
  ];

  if (protectedPages.indexOf(page) !== -1) {
    const token = e.parameter.token || '';
    const session = getSessionFromToken_(token);

    if (!session.valid) {
      const expiredLoginTemplate = HtmlService.createTemplateFromFile('Login');
      expiredLoginTemplate.notificationDurationMs = notificationDurationMs;
      applyBrandToTemplate_(expiredLoginTemplate, brand);
      return expiredLoginTemplate.evaluate()
        .setTitle('Đăng nhập')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const template = HtmlService.createTemplateFromFile(page);

    template.token = token;
    template.appUrl = ScriptApp.getService().getUrl();
    template.maKyHoc = session.maKyHoc;
    template.kyHocId = session.maKyHoc;
    template.tenKyHoc = session.tenKyHoc;
    template.kyHocName = session.tenKyHoc;
    template.currentUserName = session.hoTen || session.tenDangNhap || '';
    template.currentUserRole = session.vaiTro || '';
    template.cacheEnabled = isDataCacheEnabled_();
    template.notificationDurationMs = notificationDurationMs;
    applyBrandToTemplate_(template, brand);

    return template.evaluate()
      .setTitle(brand.name)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const loginTemplate = HtmlService.createTemplateFromFile('Login');
  loginTemplate.notificationDurationMs = notificationDurationMs;
  applyBrandToTemplate_(loginTemplate, brand);

  return loginTemplate.evaluate()
    .setTitle('Đăng nhập')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
   SETUP DATABASE
========================================================= */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureArchitectureSheets_();

  const kyHocSheet = ensureSheet_(ss, SHEET_KYHOC, [
    'MaKyHoc', 'TenKyHoc', 'TrangThai', 'MacDinh'
  ]);

  const khoiSheet = ensureSheet_(ss, SHEET_KHOI, [
    'Khoi', 'TenKhoi', 'ThuTu', 'TrangThai'
  ]);

  const lopSheet = ensureSheet_(ss, SHEET_LOP, [
    'MaLop', 'TenLop', 'Khoi', 'ThuTu', 'TrangThai'
  ]);

  ensureSheet_(ss, SHEET_HOCSINH, getHocSinhHeaders_());
  ensureSheet_(ss, SHEET_HOCSINH_KYHOC, getHocSinhKyHocHeaders_());
  ensureSheet_(ss, SHEET_DANHMUC_THUCHI, getDanhMucThuChiHeaders_());
  ensureSheet_(ss, SHEET_SOTHUCHI, getSoThuChiHeaders_());
  ensureSheet_(ss, SHEET_NGUONTIEN, getNguonTienHeaders_());
  ensureSheet_(ss, SHEET_DOITUONG_THUCHI, getDoiTuongThuChiHeaders_());
  ensureSheet_(ss, SHEET_KEHOACH_CHI_THANG, getKeHoachChiThangHeaders_());
  ensureSheet_(ss, SHEET_CAUHINH_TAICHINH_THANG, getCauHinhTaiChinhThangHeaders_());
  ensureQuanLyTaiChinhSheets_();
  ensureSheet_(ss, SHEET_DIEMDANH, getDiemDanhHeaders_());
  ensureDanhMucKhoanThuPhiSheet_();
  ensureAppConfigSheet_();

  // Dữ liệu thu phí được lưu theo từng sheet tháng, ví dụ: Thang07.2026.
  // Không tạo thêm sheet ThuPhi tổng hợp để tránh dữ liệu trùng lặp và nặng file.

  if (kyHocSheet.getLastRow() < 2) {
    kyHocSheet.getRange(2, 1, 1, 4).setValues([[
      'HE_2026', 'Khoá học hè 2026', 'ACTIVE', true
    ]]);
  }

  if (khoiSheet.getLastRow() < 2) {
    khoiSheet.getRange(2, 1, 9, 4).setValues([
      ['1', 'Khối 1', 1, 'ACTIVE'],
      ['2', 'Khối 2', 2, 'ACTIVE'],
      ['3', 'Khối 3', 3, 'ACTIVE'],
      ['4', 'Khối 4', 4, 'ACTIVE'],
      ['5', 'Khối 5', 5, 'ACTIVE'],
      ['6', 'Khối 6', 6, 'ACTIVE'],
      ['7', 'Khối 7', 7, 'ACTIVE'],
      ['8', 'Khối 8', 8, 'ACTIVE'],
      ['9', 'Khối 9', 9, 'ACTIVE']
    ]);
  }

  if (lopSheet.getLastRow() < 2) {
    lopSheet.getRange(2, 1, 9, 5).setValues([
      ['L1', 'Lớp 1', '1', 1, 'ACTIVE'],
      ['L2', 'Lớp 2', '2', 2, 'ACTIVE'],
      ['L3', 'Lớp 3', '3', 3, 'ACTIVE'],
      ['L4', 'Lớp 4', '4', 4, 'ACTIVE'],
      ['L5', 'Lớp 5', '5', 5, 'ACTIVE'],
      ['L6', 'Lớp 6', '6', 6, 'ACTIVE'],
      ['L7', 'Lớp 7', '7', 7, 'ACTIVE'],
      ['L8', 'Lớp 8', '8', 8, 'ACTIVE'],
      ['L9', 'Lớp 9', '9', 9, 'ACTIVE']
    ]);
  }

  ensureThuChiSheets_();

  ss.getSheets().forEach(sheet => {
    formatHeader_(sheet);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  bumpDataVersion_();

  return 'Đã setup xong database.';
}

function ensureSheet_(ss, sheetName, requiredHeaders) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const values = sheet.getDataRange().getValues();
  const currentHeaders = values && values.length
    ? values[0].map(h => String(h || '').trim()).filter(h => h)
    : [];

  let headers = currentHeaders.length ? currentHeaders.slice() : [];

  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
    }
  });

  if (headers.length === 0) {
    headers = requiredHeaders.slice();
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader_(sheet);

  return sheet;
}

function formatHeader_(sheet) {
  const lastCol = sheet.getLastColumn();

  if (lastCol <= 0) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0284c7');

  sheet.setFrozenRows(1);
}

function applyBrandToTemplate_(template, brand) {
  template.brandName = brand.name;
  template.brandTagline = brand.tagline;
  template.brandLogoUrl = brand.logoUrl;
}

function ensureAppConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CAUHINH);
  const headers = ['MaCauHinh', 'GiaTri', 'MoTa'];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CAUHINH);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(sheet);
  }

  let values = sheet.getDataRange().getValues();
  let currentHeaders = values[0].map(value => String(value || '').trim());
  if (headers.some(header => currentHeaders.indexOf(header) === -1)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(sheet);
    values = sheet.getDataRange().getValues();
    currentHeaders = headers.slice();
  }
  const keyIndex = currentHeaders.indexOf('MaCauHinh');
  const hasNotificationConfig = keyIndex !== -1 && values.slice(1).some(row => {
    return String(row[keyIndex] || '').trim().toUpperCase() === CONFIG_NOTIFICATION_DURATION;
  });

  if (!hasNotificationConfig) {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([[
      CONFIG_NOTIFICATION_DURATION,
      5,
      'Thá»i gian hiá»ƒn thá»‹ notification, tÃ­nh báº±ng giÃ¢y (1-300).'
    ]]);
  }

  const configRows = [
    [CONFIG_BRAND_NAME, 'Thiên Ân Education', 'Tên cơ sở dạy học hiển thị trong toàn bộ phần mềm.'],
    [CONFIG_BRAND_TAGLINE, 'Khơi nguồn tri thức', 'Khẩu hiệu của cơ sở dạy học.'],
    [CONFIG_BRAND_LOGO_URL, '', 'URL ảnh logo PNG/JPG/SVG dùng trong toàn bộ phần mềm.']
  ];
  const existingKeys = values.slice(1).map(row => String(row[keyIndex] || '').trim().toUpperCase());
  configRows.forEach(row => {
    if (existingKeys.indexOf(row[0]) === -1) {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
    }
  });

  return sheet;
}

function getAppBrandConfig_() {
  const sheet = ensureAppConfigSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value || '').trim());
  const keyIndex = headers.indexOf('MaCauHinh');
  const valueIndex = headers.indexOf('GiaTri');
  const map = {};
  if (keyIndex !== -1 && valueIndex !== -1) {
    values.slice(1).forEach(row => {
      map[String(row[keyIndex] || '').trim().toUpperCase()] = String(row[valueIndex] || '').trim();
    });
  }
  return {
    name: map[CONFIG_BRAND_NAME] || 'Thiên Ân Education',
    tagline: map[CONFIG_BRAND_TAGLINE] || 'Khơi nguồn tri thức',
    logoUrl: map[CONFIG_BRAND_LOGO_URL] || ''
  };
}

function getNotificationDurationMs_() {
  const sheet = ensureAppConfigSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 5000;

  const headers = values[0].map(value => String(value || '').trim());
  const keyIndex = headers.indexOf('MaCauHinh');
  const valueIndex = headers.indexOf('GiaTri');
  if (keyIndex === -1 || valueIndex === -1) return 5000;

  const row = values.slice(1).find(item => {
    return String(item[keyIndex] || '').trim().toUpperCase() === CONFIG_NOTIFICATION_DURATION;
  });
  const seconds = row ? number_(row[valueIndex]) : 5;
  return Math.min(300, Math.max(1, seconds || 5)) * 1000;
}

/* =========================================================
   LOGIN / SESSION / CACHE CONTROL
========================================================= */

function getKyHocList() {
  return jsonResponse_(getKyHocArray_());
}

function getKyHocArray_() {
  const rows = readObjects_(SHEET_KYHOC);

  return rows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() === 'ACTIVE')
    .map(row => ({
      maKyHoc: String(row.MaKyHoc || '').trim(),
      tenKyHoc: String(row.TenKyHoc || '').trim(),
      macDinh: row.MacDinh === true || String(row.MacDinh).toUpperCase() === 'TRUE'
    }));
}

function checkLogin(password, maKyHoc, username) {
  const inputPassword = String(password || '').trim();
  const inputKyHoc = String(maKyHoc || '').trim();
  const inputUsername = String(username || 'admin').trim().toLowerCase();

  if (!inputKyHoc) {
    return jsonResponse_({
      success: false,
      message: 'Vui lòng chọn kỳ học.'
    });
  }

  if (!inputPassword) {
    return jsonResponse_({
      success: false,
      message: 'Vui lòng nhập mật khẩu.'
    });
  }

  const auth = SecurityService.authenticate(inputUsername, inputPassword, inputKyHoc);
  if (!auth.success) return jsonResponse_(auth);
  const webAppUrl = ScriptApp.getService().getUrl();
  const session = {
    valid: true,
    maKyHoc: auth.maKyHoc,
    tenKyHoc: auth.tenKyHoc,
    maNguoiDung: auth.user.maNguoiDung,
    tenDangNhap: auth.user.tenDangNhap,
    hoTen: auth.user.hoTen,
    vaiTro: auth.user.vaiTro
  };
  safeWriteAuditLog_(session, 'LOGIN', 'PHIEN_DANG_NHAP', auth.token.slice(0, 8), null, {
    maKyHoc: auth.maKyHoc
  });

  return jsonResponse_({
    success: true,
    message: 'Đăng nhập thành công.',
    maKyHoc: auth.maKyHoc,
    tenKyHoc: auth.tenKyHoc,
    user: auth.user,
    token: auth.token,
    redirectUrl: webAppUrl + '?page=Index&token=' + encodeURIComponent(auth.token)
  });
}

function logout(token) {
  const session = getSessionFromToken_(token);
  if (token) {
    CacheService.getScriptCache().remove(CACHE_LOGIN_PREFIX + token);
  }
  safeWriteAuditLog_(session, 'LOGOUT', 'PHIEN_DANG_NHAP', String(token || '').slice(0, 8), null, null);

  return jsonResponse_({
    success: true,
    message: 'Đã đăng xuất.'
  });
}

/**
 * Bật/tắt cache dữ liệu nghiệp vụ. Mặc định bật để tối ưu tốc độ đọc.
 * Cache token đăng nhập vẫn hoạt động độc lập để duy trì phiên đăng nhập.
 */
function setDataCacheEnabled(token, enabled) {
  requireSession_(token, 'system.admin');

  const isEnabled = enabled === true || String(enabled).toLowerCase() === 'true';
  PropertiesService.getScriptProperties().setProperty(
    DATA_CACHE_ENABLED_PROPERTY,
    isEnabled ? 'TRUE' : 'FALSE'
  );
  DATA_CACHE_ENABLED_MEMO_ = isEnabled;

  // Đổi phiên bản để mọi cache cũ không còn được dùng lại.
  bumpDataVersion_();

  return jsonResponse_({
    success: true,
    enabled: isEnabled,
    message: isEnabled
      ? 'Đã bật cache dữ liệu.'
      : 'Đã tắt cache dữ liệu. Hệ thống sẽ đọc trực tiếp từ Google Sheet.'
  });
}

function isDataCacheEnabled_() {
  if (DATA_CACHE_ENABLED_MEMO_ !== null) return DATA_CACHE_ENABLED_MEMO_;

  DATA_CACHE_ENABLED_MEMO_ = String(
    PropertiesService.getScriptProperties().getProperty(DATA_CACHE_ENABLED_PROPERTY) || 'TRUE'
  ).toUpperCase() === 'TRUE';

  return DATA_CACHE_ENABLED_MEMO_;
}

/**
 * Xoá bộ nhớ tạm bằng cách đổi DATA_VERSION.
 * Cache cũ sẽ tự hết hạn, cache mới sẽ được tạo lại từ Google Sheets.
 */
function clearCacheAndReload(token) {
  const session = requireSession_(token, 'system.admin');

  bumpDataVersion_();

  if (isDataCacheEnabled_()) {
    warmCacheForSession_(session.maKyHoc);
  }

  return jsonResponse_({
    success: true,
    cacheEnabled: isDataCacheEnabled_(),
    message: isDataCacheEnabled_()
      ? 'Đã xoá cache cũ và nạp lại dữ liệu từ Google Sheet.'
      : 'Cache đang tắt. Dữ liệu sẽ được đọc trực tiếp từ Google Sheet.'
  });
}

function warmCacheForSession_(maKyHoc) {
  readObjects_(SHEET_KYHOC);
  readObjects_(SHEET_KHOI);
  readObjects_(SHEET_LOP);
  readObjects_(SHEET_HOCSINH);
  readObjects_(SHEET_HOCSINH_KYHOC);
  readObjects_(SHEET_DANHMUC_THUCHI);
  readObjects_(SHEET_SOTHUCHI);

  // Không tự khởi tạo tháng thu phí khi chỉ xoá cache.
  // Sheet tháng chỉ được tạo khi người dùng thật sự chọn tháng tại trang thu phí.
}

function createLoginToken_(maKyHoc) {
  const term = getKyHocArray_().find(item => item.maKyHoc === maKyHoc);
  if (!term) throw new Error('Kỳ học không hợp lệ hoặc chưa được kích hoạt.');
  return SecurityService.createLegacyOwnerSession(term);
}

function getSessionFromToken_(token) {
  return SecurityService.getSession(token);
}

function requireSession_(token, permission) {
  return SecurityService.requireSession(token, permission);
}

/* =========================================================
   QUẢN LÝ HỌC SINH
========================================================= */

function getInitialHocSinhData(token) {
  const session = requireSession_(token, 'student.read');
  const cacheKey = buildCacheKey_('initial_hocsinh_' + session.maKyHoc);

  const cached = cacheGetString_(cacheKey);
  if (cached) return cached;

  const data = {
    session: session,
    kyHocList: getKyHocArray_(),
    khoiList: getKhoiList_(),
    lopList: getLopList_()
  };

  const json = jsonResponse_(data);
  cachePutString_(cacheKey, json, CACHE_SECONDS);

  return json;
}

function getKhoiList_() {
  const rows = readObjects_(SHEET_KHOI);

  return rows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() === 'ACTIVE')
    .map(row => ({
      khoi: String(row.Khoi || '').trim(),
      tenKhoi: String(row.TenKhoi || '').trim(),
      thuTu: Number(row.ThuTu) || 999
    }))
    .sort((a, b) => a.thuTu - b.thuTu);
}

function getLopList_() {
  const rows = readObjects_(SHEET_LOP);
  const khoiOrderMap = getKhoiList_().reduce((map, item, index) => {
    map[item.khoi] = number_(item.thuTu) || index + 1;
    return map;
  }, {});

  return rows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() === 'ACTIVE')
    .map(row => ({
      maLop: String(row.MaLop || '').trim(),
      tenLop: String(row.TenLop || '').trim(),
      khoi: String(row.Khoi || '').trim(),
      thuTu: Number(row.ThuTu) || 999
    }))
    .sort((a, b) => {
      if (a.khoi !== b.khoi) {
        return number_(khoiOrderMap[a.khoi] || 999) - number_(khoiOrderMap[b.khoi] || 999);
      }
      return a.thuTu - b.thuTu;
    });
}

function getHocSinhList(token, filters) {
  const session = requireSession_(token, 'student.read');

  filters = filters || {};

  const cacheKey = buildCacheKey_(
    'hocsinh_list_' +
    session.maKyHoc + '_' +
    hashString_(JSON.stringify(filters))
  );

  const cached = cacheGetString_(cacheKey);
  if (cached) return cached;

  const filterKhoi = String(filters.khoi || '').trim();
  const filterLop = String(filters.lop || '').trim();
  const keyword = normalizeText_(filters.keyword || '');

  const hocSinhRows = readObjects_(SHEET_HOCSINH);
  const hocSinhKyHocMap = getHocSinhKyHocMap_();
  const hasCurrentKyHocMapping = hasMappingForKyHoc_(session.maKyHoc);

  let result = hocSinhRows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED')
    .map(row => {
      const maHocSinh = String(row.MaHocSinh || '').trim();

      if (!maHocSinh) return null;

      const mapping = hocSinhKyHocMap[maHocSinh] || {
        kyHocIds: [],
        kyHocNames: [],
        byKyHoc: {}
      };

      const rowMaKyHoc = String(row.MaKyHoc || '').trim();

      const belongsToCurrentKyHoc =
        mapping.kyHocIds.indexOf(session.maKyHoc) !== -1 ||
        rowMaKyHoc === session.maKyHoc ||
        (!hasCurrentKyHocMapping && !rowMaKyHoc);

      if (!belongsToCurrentKyHoc) return null;

      const currentFee = mapping.byKyHoc[session.maKyHoc] || {};
      const ngayVao = row.NgayVao || row.NgaySinh || '';

      return {
        maHocSinh: maHocSinh,
        sapXep: number_(row.SapXep),
        maKyHoc: rowMaKyHoc,
        khoi: String(row.Khoi || '').trim(),
        lop: String(row.Lop || '').trim(),
        hoTen: String(row.HoTen || '').trim(),
        truong: String(row.Truong || '').trim() || 'THCS Long Phước',
        ngaySinh: formatDateForInput_(ngayVao),
        ngaySinhDisplay: formatDateDisplay_(ngayVao),
        gioiTinh: String(row.GioiTinh || '').trim(),
        khongThuPhi: toBoolean_(row.KhongThuPhi),
        sdtPhuHuynh: String(row.SDTPhuHuynh || '').trim(),
        diaChi: String(row.DiaChi || '').trim(),
        ghiChu: String(row.GhiChu || '').trim(),

        kyHocIds: mapping.kyHocIds,
        kyHocNames: mapping.kyHocNames,
        hocPhi: currentFee.hocPhi || '',
        trangThaiHocPhi: currentFee.trangThaiHocPhi || '',
        ghiChuHocPhi: currentFee.ghiChuHocPhi || ''
      };
    })
    .filter(item => item !== null);

  if (filterKhoi) {
    result = result.filter(item => item.khoi === filterKhoi);
  }

  if (filterLop) {
    result = result.filter(item => item.lop === filterLop);
  }

  if (keyword) {
    result = result.filter(item => {
      const haystack = normalizeText_(
        item.hoTen + ' ' +
        item.truong + ' ' +
        item.sdtPhuHuynh + ' ' +
        item.ghiChu
      );

      return haystack.indexOf(keyword) !== -1;
    });
  }

  result.sort(compareStudentSort_);

  const json = jsonResponse_(result);
  cachePutString_(cacheKey, json, CACHE_SECONDS);

  return json;
}

function getDiemDanhData(token, dateText) {
  const session = requireSession_(token, 'attendance.read');
  const ngay = String(dateText || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) throw new Error('Ngày điểm danh không hợp lệ.');
  ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_DIEMDANH, getDiemDanhHeaders_());

  const students = JSON.parse(getHocSinhList(token, {}));
  const attendanceMap = readObjectsNoCache_(SHEET_DIEMDANH)
    .filter(row => String(row.MaKyHoc || '').trim() === session.maKyHoc && formatDateForInput_(row.NgayDiemDanh) === ngay)
    .reduce((map, row) => {
      map[String(row.MaHocSinh || '').trim()] = {
        trangThai: String(row.TrangThai || 'CHUA_DIEM_DANH').trim().toUpperCase(),
        ghiChu: String(row.GhiChu || '').trim()
      };
      return map;
    }, {});

  const result = students.map(student => {
    const saved = attendanceMap[student.maHocSinh] || {};
    return Object.assign({}, student, {
      trangThaiDiemDanh: saved.trangThai || 'CHUA_DIEM_DANH',
      ghiChuDiemDanh: saved.ghiChu || ''
    });
  });
  return jsonResponse_({
    date: ngay,
    students: result,
    khoiList: getKhoiList_(),
    lopList: getLopList_()
  });
}

function getTheoDoiDiemDanhData(token, fromDateText, toDateText) {
  const session = requireSession_(token, 'attendance.read');
  const fromText = String(fromDateText || '').trim();
  const toText = String(toDateText || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fromText) || !/^\d{4}-\d{2}-\d{2}$/.test(toText)) {
    throw new Error('Khoảng ngày theo dõi không hợp lệ.');
  }
  const fromDate = parseInputDate_(fromText);
  const toDate = parseInputDate_(toText);
  const dayCount = Math.floor((toDate.getTime() - fromDate.getTime()) / 86400000) + 1;
  if (dayCount < 1) throw new Error('Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.');
  if (dayCount > 93) throw new Error('Khoảng theo dõi tối đa là 93 ngày.');

  ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_DIEMDANH, getDiemDanhHeaders_());
  const students = JSON.parse(getHocSinhList(token, {}));
  const studentIds = new Set(students.map(item => item.maHocSinh));
  const records = readObjectsNoCache_(SHEET_DIEMDANH)
    .filter(row => {
      const date = formatDateForInput_(row.NgayDiemDanh);
      const id = String(row.MaHocSinh || '').trim();
      return String(row.MaKyHoc || '').trim() === session.maKyHoc &&
        date >= fromText && date <= toText && studentIds.has(id);
    })
    .map(row => ({
      ngay: formatDateForInput_(row.NgayDiemDanh),
      maHocSinh: String(row.MaHocSinh || '').trim(),
      trangThai: String(row.TrangThai || 'CHUA_DIEM_DANH').trim().toUpperCase(),
      ghiChu: String(row.GhiChu || '').trim()
    }));

  return jsonResponse_({
    fromDate: fromText,
    toDate: toText,
    dayCount: dayCount,
    students: students,
    records: records,
    khoiList: getKhoiList_(),
    lopList: getLopList_()
  });
}

function saveDiemDanh(token, dateText, records) {
  const session = requireSession_(token, 'attendance.write');
  const ngay = String(dateText || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) throw new Error('Ngày điểm danh không hợp lệ.');
  records = Array.isArray(records) ? records : [];
  if (!records.length) throw new Error('Chưa có dữ liệu điểm danh để lưu.');

  const validStatuses = ['CHUA_DIEM_DANH', 'CO_MAT', 'VANG_CO_PHEP', 'VANG_KHONG_PHEP', 'DI_MUON'];
  const studentIds = new Set(JSON.parse(getHocSinhList(token, {})).map(item => item.maHocSinh));
  const cleanRecords = records.map(item => ({
    maHocSinh: String(item.maHocSinh || '').trim(),
    trangThai: String(item.trangThai || '').trim().toUpperCase(),
    ghiChu: String(item.ghiChu || '').trim().slice(0, 500)
  })).filter(item => studentIds.has(item.maHocSinh) && validStatuses.indexOf(item.trangThai) !== -1);
  if (!cleanRecords.length) throw new Error('Không có học sinh hợp lệ để lưu điểm danh.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang có người lưu điểm danh. Vui lòng thử lại.');
  try {
    const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_DIEMDANH, getDiemDanhHeaders_());
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value || '').trim());
    const values = sheet.getDataRange().getValues();
    const index = buildHeaderIndex_(headers);
    const rowMap = {};
    values.slice(1).forEach((row, offset) => {
      const key = String(row[index.MaKyHoc] || '').trim() + '|' + formatDateForInput_(row[index.NgayDiemDanh]) + '|' + String(row[index.MaHocSinh] || '').trim();
      rowMap[key] = offset + 1;
    });
    const now = new Date();
    const dirtyRows = [];
    cleanRecords.forEach(item => {
      const key = session.maKyHoc + '|' + ngay + '|' + item.maHocSinh;
      let rowIndex = rowMap[key];
      if (rowIndex === undefined) {
        rowIndex = values.length;
        values.push(new Array(headers.length).fill(''));
        rowMap[key] = rowIndex;
        values[rowIndex][index.MaDiemDanh] = 'DD_' + Utilities.getUuid().slice(0, 12).toUpperCase();
        values[rowIndex][index.MaKyHoc] = session.maKyHoc;
        values[rowIndex][index.NgayDiemDanh] = parseInputDate_(ngay);
        values[rowIndex][index.MaHocSinh] = item.maHocSinh;
      }
      values[rowIndex][index.TrangThai] = item.trangThai;
      values[rowIndex][index.GhiChu] = item.ghiChu;
      values[rowIndex][index.UpdatedAt] = now;
      dirtyRows.push(rowIndex);
    });
    const sortedDirtyRows = Array.from(new Set(dirtyRows)).sort((a, b) => a - b);
    let groupStart = sortedDirtyRows[0];
    let groupEnd = groupStart;
    for (let i = 1; i <= sortedDirtyRows.length; i++) {
      const current = sortedDirtyRows[i];
      if (current === groupEnd + 1) {
        groupEnd = current;
        continue;
      }
      sheet.getRange(groupStart + 1, 1, groupEnd - groupStart + 1, headers.length)
        .setValues(values.slice(groupStart, groupEnd + 1));
      groupStart = current;
      groupEnd = current;
    }
  } finally {
    lock.releaseLock();
  }
  bumpDataVersion_();
  safeWriteAuditLog_(session, 'UPSERT', 'DIEM_DANH', ngay, null, { soHocSinh: cleanRecords.length });
  return jsonResponse_({ success: true, savedCount: cleanRecords.length, message: 'Đã lưu điểm danh ' + cleanRecords.length + ' học sinh.' });
}

function saveHocSinhOrder(token, orderedStudentIds) {
  const session = requireSession_(token, 'student.write');
  const orderedIds = Array.isArray(orderedStudentIds)
    ? orderedStudentIds.map(id => String(id || '').trim()).filter(id => id)
    : [];

  if (!orderedIds.length) {
    throw new Error('Danh sách sắp xếp không có học sinh.');
  }

  if (new Set(orderedIds).size !== orderedIds.length) {
    throw new Error('Danh sách sắp xếp có học sinh bị trùng.');
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang có người cập nhật danh sách học sinh. Vui lòng thao tác lại.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_HOCSINH);

    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('Chưa có dữ liệu học sinh để sắp xếp.');
    }

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);

    if (index.MaHocSinh === undefined || index.SapXep === undefined) {
      throw new Error('Sheet học sinh thiếu cột MaHocSinh hoặc SapXep.');
    }

    const relationRows = readObjectsNoCache_(SHEET_HOCSINH_KYHOC);
    const mappedIds = {};

    relationRows.forEach(row => {
      if (
        String(row.MaKyHoc || '').trim() === session.maKyHoc &&
        String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED'
      ) {
        const id = String(row.MaHocSinh || '').trim();
        if (id) mappedIds[id] = true;
      }
    });

    const hasCurrentMapping = Object.keys(mappedIds).length > 0;
    const currentStudentIds = values.slice(1).reduce((ids, row) => {
      const id = String(row[index.MaHocSinh] || '').trim();
      const status = index.TrangThai === undefined
        ? 'ACTIVE'
        : String(row[index.TrangThai] || '').trim().toUpperCase();
      const rowKyHoc = index.MaKyHoc === undefined
        ? ''
        : String(row[index.MaKyHoc] || '').trim();

      if (!id || status === 'DELETED') return ids;

      const belongsToCurrentKyHoc =
        !!mappedIds[id] ||
        rowKyHoc === session.maKyHoc ||
        (!hasCurrentMapping && !rowKyHoc);

      if (belongsToCurrentKyHoc) ids.push(id);
      return ids;
    }, []);

    const currentIdMap = currentStudentIds.reduce((map, id) => {
      map[id] = true;
      return map;
    }, {});

    if (
      orderedIds.length !== currentStudentIds.length ||
      orderedIds.some(id => !currentIdMap[id])
    ) {
      throw new Error('Danh sách học sinh đã thay đổi. Vui lòng đóng cửa sổ sắp xếp và tải lại dữ liệu.');
    }

    const lopOrderMap = getLopList_().reduce((map, item) => {
      map[String(item.khoi || '') + '|' + String(item.maLop || '')] = number_(item.thuTu);
      return map;
    }, {});
    const studentGroupMap = values.slice(1).reduce((map, row) => {
      const id = String(row[index.MaHocSinh] || '').trim();
      if (!id || !currentIdMap[id]) return map;

      const khoi = index.Khoi === undefined ? '' : String(row[index.Khoi] || '').trim();
      const lop = index.Lop === undefined ? '' : String(row[index.Lop] || '').trim();
      const groupKey = khoi + '|' + lop;
      map[id] = {
        groupKey: groupKey,
        classNumber: number_(lopOrderMap[groupKey])
      };
      return map;
    }, {});
    const nextOrderByGroup = {};
    const orderMap = orderedIds.reduce((map, id) => {
      const group = studentGroupMap[id];

      if (!group || group.classNumber < 1 || group.classNumber > 9) {
        throw new Error('Lớp học thiếu thứ tự hợp lệ. Cột ThuTu trong sheet Lop phải từ 1 đến 9.');
      }

      const position = number_(nextOrderByGroup[group.groupKey]);

      if (position > 99) {
        throw new Error('Mỗi lớp chỉ hỗ trợ tối đa 100 vị trí sắp xếp.');
      }

      map[id] = group.classNumber * 100 + position;
      nextOrderByGroup[group.groupKey] = position + 1;
      return map;
    }, {});
    const now = new Date();

    for (let i = 1; i < values.length; i++) {
      const id = String(values[i][index.MaHocSinh] || '').trim();
      if (!orderMap[id]) continue;

      values[i][index.SapXep] = orderMap[id];
      if (index.UpdatedAt !== undefined) values[i][index.UpdatedAt] = now;
    }

    sheet
      .getRange(2, 1, values.length - 1, headers.length)
      .setValues(values.slice(1));

    ss.getSheets().forEach(monthSheet => {
      if (!/^Thang\d{2}\.\d{4}$/.test(monthSheet.getName()) || monthSheet.getLastRow() < 2) {
        return;
      }

      const monthValues = monthSheet.getDataRange().getValues();
      const monthHeaders = monthValues[0].map(header => String(header || '').trim());
      const monthIndex = buildHeaderIndex_(monthHeaders);

      if (
        monthIndex.MaHocSinh === undefined ||
        monthIndex.MaKyHoc === undefined ||
        monthIndex.SapXep === undefined
      ) {
        return;
      }

      let monthChanged = false;

      for (let i = 1; i < monthValues.length; i++) {
        const id = String(monthValues[i][monthIndex.MaHocSinh] || '').trim();
        const maKyHoc = String(monthValues[i][monthIndex.MaKyHoc] || '').trim();

        if (maKyHoc !== session.maKyHoc || !orderMap[id]) continue;
        if (number_(monthValues[i][monthIndex.SapXep]) === orderMap[id]) continue;

        monthValues[i][monthIndex.SapXep] = orderMap[id];
        monthChanged = true;
      }

      if (monthChanged) {
        monthSheet
          .getRange(2, 1, monthValues.length - 1, monthHeaders.length)
          .setValues(monthValues.slice(1));
      }
    });
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();

  return jsonResponse_({
    success: true,
    updatedCount: orderedIds.length,
    message: 'Đã cập nhật SapXep theo dải lớp cho ' + orderedIds.length + ' học sinh.'
  });
}

function saveHocSinh(token, hocSinh) {
  const session = requireSession_(token, 'student.write');

  hocSinh = hocSinh || {};

  const maHocSinh = String(hocSinh.maHocSinh || '').trim();
  const kyHocIds = Array.isArray(hocSinh.kyHocIds) ? hocSinh.kyHocIds : [];
  let sapXep = String(hocSinh.sapXep || '').trim();
  const khoi = String(hocSinh.khoi || '').trim();
  const lop = String(hocSinh.lop || '').trim();
  const hoTen = String(hocSinh.hoTen || '').trim();
  const truong = String(hocSinh.truong || '').trim() || 'THCS Long Phước';
  const ngayVao = String(hocSinh.ngaySinh || '').trim();
  const gioiTinh = String(hocSinh.gioiTinh || '').trim();
  const sdtPhuHuynh = String(hocSinh.sdtPhuHuynh || '').trim();
  const diaChi = String(hocSinh.diaChi || '').trim();
  const ghiChu = String(hocSinh.ghiChu || '').trim();
  const hocPhi = String(hocSinh.hocPhi || '').replace(/[^\d]/g, '');
  const khongThuPhi = toBoolean_(hocSinh.khongThuPhi);
  const capNhatThuPhi = !khongThuPhi && toBoolean_(hocSinh.capNhatThuPhi);
  const thuPhiYearMonth = String(hocSinh.thuPhiYearMonth || '').trim();

  if (kyHocIds.length === 0) {
    throw new Error('Vui lòng chọn ít nhất một kỳ học.');
  }

  if (!khoi) {
    throw new Error('Vui lòng chọn khối.');
  }

  if (!lop) {
    throw new Error('Vui lòng chọn lớp.');
  }

  if (!hoTen) {
    throw new Error('Vui lòng nhập họ tên học sinh.');
  }

  if (['TH Tam Thiện', 'THCS Phước Thái', 'THCS Long Phước'].indexOf(truong) === -1) {
    throw new Error('Trường học không hợp lệ.');
  }

  if (sdtPhuHuynh && !/^0\d{9}$/.test(sdtPhuHuynh)) {
    throw new Error('Số điện thoại phụ huynh phải gồm 10 số và bắt đầu bằng số 0.');
  }

  if (capNhatThuPhi && kyHocIds.indexOf(session.maKyHoc) === -1) {
    throw new Error('Muốn cập nhật thu phí tháng hiện tại, học sinh phải thuộc kỳ học đang đăng nhập.');
  }

  if (capNhatThuPhi && !/^\d{4}-\d{2}$/.test(thuPhiYearMonth)) {
    throw new Error('Tháng cập nhật thu phí không hợp lệ.');
  }
  if (/^\d{4}-\d{2}$/.test(thuPhiYearMonth) && (capNhatThuPhi || maHocSinh)) {
    assertFinancePeriodOpen_(session, thuPhiYearMonth);
  }

  const validLop = getLopList_().find(item => item.maLop === lop && item.khoi === khoi);

  if (!validLop) {
    throw new Error('Lớp không thuộc khối đã chọn.');
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang có người cập nhật học sinh. Vui lòng thao tác lại.');
  }

  let newId = maHocSinh;
  let existed = false;
  let savedStudentRow = null;

  try {
  const now = new Date();
  const rows = readObjectsNoCache_(SHEET_HOCSINH);

  const classNumber = number_(validLop.thuTu);
  const sortRangeStart = classNumber * 100;
  const sortRangeEnd = sortRangeStart + 99;
  const requestedSort = number_(sapXep);
  const requestedSortUsed = rows.some(row => {
    return String(row.MaHocSinh || '').trim() !== maHocSinh &&
      String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED' &&
      String(row.Khoi || '').trim() === khoi &&
      String(row.Lop || '').trim() === lop &&
      number_(row.SapXep) === requestedSort;
  });

  if (classNumber < 1 || classNumber > 9) {
    throw new Error('Lớp học thiếu thứ tự hợp lệ. Cột ThuTu trong sheet Lop phải từ 1 đến 9.');
  }

  if (requestedSort < sortRangeStart || requestedSort > sortRangeEnd || requestedSortUsed) {
    const maxSortInClass = rows.reduce((maxValue, row) => {
      const rowId = String(row.MaHocSinh || '').trim();
      const rowStatus = String(row.TrangThai || '').trim().toUpperCase();

      if (
        rowId !== maHocSinh &&
        rowStatus !== 'DELETED' &&
        String(row.Khoi || '').trim() === khoi &&
        String(row.Lop || '').trim() === lop
      ) {
        const value = number_(row.SapXep);
        if (value >= sortRangeStart && value <= sortRangeEnd) {
          return Math.max(maxValue, value);
        }
      }

      return maxValue;
    }, sortRangeStart - 1);
    const availableSort = maxSortInClass + 1;

    if (availableSort > sortRangeEnd) {
      throw new Error('Lớp đã sử dụng hết dải SapXep từ ' + sortRangeStart + ' đến ' + sortRangeEnd + '.');
    }

    sapXep = String(availableSort);
  }

  const existingRow = rows.find(row => {
    return String(row.MaHocSinh || '').trim() === maHocSinh &&
      String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
  });

  existed = !!existingRow;

  if (!existed) {
    newId = 'HS_' + Utilities.getUuid().slice(0, 8).toUpperCase();
  }

  const parsedNgayVao = parseInputDate_(ngayVao);
  savedStudentRow = Object.assign({}, existingRow || {}, {
    MaHocSinh: newId,
    MaKyHoc: session.maKyHoc,
    SapXep: sapXep,
    Khoi: khoi,
    Lop: lop,
    HoTen: hoTen,
    Truong: truong,
    NgayVao: parsedNgayVao,
    NgaySinh: parsedNgayVao,
    GioiTinh: gioiTinh,
    KhongThuPhi: khongThuPhi ? 'Có' : 'Không',
    SDTPhuHuynh: sdtPhuHuynh,
    DiaChi: diaChi,
    GhiChu: ghiChu,
    TrangThai: 'ACTIVE',
    CreatedAt: existingRow && existingRow.CreatedAt ? existingRow.CreatedAt : now,
    UpdatedAt: now
  });

  if (existed) {
    updateObjectRowById_(SHEET_HOCSINH, 'MaHocSinh', newId, savedStudentRow, getHocSinhHeaders_());
  } else {
    const studentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ensureSheet_(studentSpreadsheet, SHEET_HOCSINH, getHocSinhHeaders_());
    appendObjectsToSheet_(studentSheet, [savedStudentRow], getHocSinhHeaders_());
  }
  saveHocSinhKyHoc_(newId, kyHocIds, hocPhi);
  } finally {
    lock.releaseLock();
  }

  let thuPhiSyncResult = null;
  if (existed && /^\d{4}-\d{2}$/.test(thuPhiYearMonth)) {
    thuPhiSyncResult = syncHocSinhToThuPhiMonth_(
      session.maKyHoc,
      thuPhiYearMonth,
      savedStudentRow,
      hocPhi
    );
  }

  // Đổi phiên bản trước khi đọc lại dữ liệu để bảo đảm không dùng cache cũ.
  bumpDataVersion_();

  let thuPhiResult = null;

  if (
    (!existed && capNhatThuPhi) ||
    (existed && !khongThuPhi && thuPhiSyncResult && !thuPhiSyncResult.updated)
  ) {
    thuPhiResult = addHocSinhToThuPhiMonth_(
      session.maKyHoc,
      thuPhiYearMonth,
      newId
    );

    bumpDataVersion_();
  }

  let message = existed
    ? 'Đã cập nhật học sinh.'
    : 'Đã thêm học sinh mới.';

  if (thuPhiResult && thuPhiResult.added) {
    message += ' Đồng thời đã thêm học sinh vào ' + thuPhiResult.sheetName + '.';
  } else if (thuPhiResult && thuPhiResult.alreadyExists) {
    message += ' Học sinh đã có trong ' + thuPhiResult.sheetName + '.';
  }

  if (thuPhiSyncResult && thuPhiSyncResult.updated) {
    message += ' Đã đồng bộ thông tin sang ' + thuPhiSyncResult.sheetName + '.';
  }
  safeWriteAuditLog_(session, existed ? 'UPDATE' : 'CREATE', 'HOC_SINH', newId, null, {
    hoTen: hoTen, khoi: khoi, lop: lop, sapXep: number_(sapXep), thuPhiYearMonth: thuPhiYearMonth
  });

  return jsonResponse_({
    success: true,
    maHocSinh: newId,
    sapXep: number_(sapXep),
    thuPhiUpdated: !!(
      (thuPhiResult && (thuPhiResult.added || thuPhiResult.alreadyExists)) ||
      (thuPhiSyncResult && thuPhiSyncResult.updated)
    ),
    thuPhiSheetName: thuPhiResult
      ? thuPhiResult.sheetName
      : (thuPhiSyncResult ? thuPhiSyncResult.sheetName : ''),
    message: message
  });
}

function deleteHocSinh(token, maHocSinh) {
  const session = requireSession_(token, 'student.write');

  const id = String(maHocSinh || '').trim();

  if (!id) {
    throw new Error('Thiếu mã học sinh cần xoá.');
  }

  const now = new Date();

  const rows = readObjectsNoCache_(SHEET_HOCSINH).map(row => {
    if (String(row.MaHocSinh || '').trim() === id) {
      return Object.assign({}, row, {
        TrangThai: 'DELETED',
        UpdatedAt: now
      });
    }

    return row;
  });

  writeObjectsToSheet_(SHEET_HOCSINH, rows, getHocSinhHeaders_());
  markHocSinhKyHocDeleted_(id);

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'DELETE', 'HOC_SINH', id, null, null);

  return jsonResponse_({
    success: true,
    message: 'Đã xoá học sinh và xoá bộ nhớ tạm.'
  });
}

function getDiemDanhHeaders_() {
  return ['MaDiemDanh', 'MaKyHoc', 'NgayDiemDanh', 'MaHocSinh', 'TrangThai', 'GhiChu', 'UpdatedAt'];
}

function getHocSinhHeaders_() {
  return [
    'MaHocSinh',
    'MaKyHoc',
    'SapXep',
    'Khoi',
    'Lop',
    'HoTen',
    'Truong',
    'NgayVao',
    'NgaySinh',
    'GioiTinh',
    'KhongThuPhi',
    'SDTPhuHuynh',
    'DiaChi',
    'GhiChu',
    'TrangThai',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function saveHocSinhKyHoc_(maHocSinh, kyHocIds, hocPhi) {
  const now = new Date();
  const oldRows = readObjectsNoCache_(SHEET_HOCSINH_KYHOC);

  const activeRows = oldRows.filter(row => {
    return String(row.MaHocSinh || '').trim() === maHocSinh &&
      String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
  });
  const activeKyHocIds = activeRows
    .map(row => String(row.MaKyHoc || '').trim())
    .filter(id => id)
    .sort();
  const requestedKyHocIds = kyHocIds.map(id => String(id || '').trim()).filter(id => id).sort();
  const mappingUnchanged = activeKyHocIds.length === requestedKyHocIds.length &&
    activeKyHocIds.every((id, index) => id === requestedKyHocIds[index]) &&
    activeRows.every(row => String(number_(row.HocPhi)) === String(number_(hocPhi)));

  if (mappingUnchanged) return;

  const keptRows = oldRows.map(row => {
    if (
      String(row.MaHocSinh || '').trim() === maHocSinh &&
      String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED'
    ) {
      return Object.assign({}, row, {
        TrangThai: 'DELETED',
        UpdatedAt: now
      });
    }

    return row;
  });

  const newRows = kyHocIds.map(maKyHoc => ({
    MaHocSinh: maHocSinh,
    MaKyHoc: maKyHoc,
    HocPhi: hocPhi,
    TrangThaiHocPhi: '',
    GhiChuHocPhi: '',
    TrangThai: 'ACTIVE',
    CreatedAt: now,
    UpdatedAt: now
  }));

  writeObjectsToSheet_(
    SHEET_HOCSINH_KYHOC,
    keptRows.concat(newRows),
    getHocSinhKyHocHeaders_()
  );
}

function markHocSinhKyHocDeleted_(maHocSinh) {
  const now = new Date();

  const rows = readObjectsNoCache_(SHEET_HOCSINH_KYHOC).map(row => {
    if (String(row.MaHocSinh || '').trim() === maHocSinh) {
      return Object.assign({}, row, {
        TrangThai: 'DELETED',
        UpdatedAt: now
      });
    }

    return row;
  });

  writeObjectsToSheet_(SHEET_HOCSINH_KYHOC, rows, getHocSinhKyHocHeaders_());
}

function getHocSinhKyHocHeaders_() {
  return [
    'MaHocSinh',
    'MaKyHoc',
    'HocPhi',
    'TrangThaiHocPhi',
    'GhiChuHocPhi',
    'TrangThai',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function getHocSinhKyHocMap_() {
  const rows = readObjects_(SHEET_HOCSINH_KYHOC);
  const kyHocNameMap = {};

  getKyHocArray_().forEach(item => {
    kyHocNameMap[item.maKyHoc] = item.tenKyHoc;
  });

  return rows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED')
    .reduce((map, row) => {
      const maHocSinh = String(row.MaHocSinh || '').trim();
      const maKyHoc = String(row.MaKyHoc || '').trim();

      if (!maHocSinh || !maKyHoc) return map;

      if (!map[maHocSinh]) {
        map[maHocSinh] = {
          kyHocIds: [],
          kyHocNames: [],
          byKyHoc: {}
        };
      }

      map[maHocSinh].kyHocIds.push(maKyHoc);
      map[maHocSinh].kyHocNames.push(kyHocNameMap[maKyHoc] || maKyHoc);
      map[maHocSinh].byKyHoc[maKyHoc] = {
        hocPhi: row.HocPhi || '',
        trangThaiHocPhi: row.TrangThaiHocPhi || '',
        ghiChuHocPhi: row.GhiChuHocPhi || ''
      };

      return map;
    }, {});
}

function hasMappingForKyHoc_(maKyHoc) {
  return readObjects_(SHEET_HOCSINH_KYHOC)
    .some(row => {
      return String(row.MaKyHoc || '').trim() === maKyHoc &&
        String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
    });
}

/* =========================================================
   QUẢN LÝ THU PHÍ THEO SHEET THÁNG
========================================================= */

function getDanhMucKhoanThuPhiHeaders_() {
  return [
    'MaKhoanThu',
    'TenKhoanThu',
    'SoTienMacDinh',
    'ThuTu',
    'TrangThai',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function ensureDanhMucKhoanThuPhiSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, SHEET_DANHMUC_KHOANTHU_PHI, getDanhMucKhoanThuPhiHeaders_());
  if (sheet.getLastRow() < 2) {
    const now = new Date();
    sheet.getRange(2, 1, 2, getDanhMucKhoanThuPhiHeaders_().length).setValues([
      ['KTP_DUNG_CU', 'Dụng cụ học tập', 0, 1, 'ACTIVE', now, now],
      ['KTP_CSVC', 'Phụ phí CSVC', 0, 2, 'ACTIVE', now, now]
    ]);
  }
  return sheet;
}

function getDanhMucKhoanThuPhiList_() {
  const sheet = ensureDanhMucKhoanThuPhiSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(header => String(header || '').trim());
  const index = buildHeaderIndex_(headers);
  return values.slice(1).map(row => ({
    maKhoanThu: String(row[index.MaKhoanThu] || '').trim(),
    tenKhoanThu: String(row[index.TenKhoanThu] || '').trim(),
    soTienMacDinh: number_(row[index.SoTienMacDinh]),
    thuTu: number_(row[index.ThuTu]),
    trangThai: String(row[index.TrangThai] || 'ACTIVE').trim().toUpperCase()
  })).filter(item => item.maKhoanThu && item.tenKhoanThu)
    .sort((a, b) => a.thuTu - b.thuTu || a.tenKhoanThu.localeCompare(b.tenKhoanThu, 'vi'));
}

function getDanhMucKhoanThuPhi(token) {
  requireSession_(token, 'tuition.read');
  return jsonResponse_({ categories: getDanhMucKhoanThuPhiList_() });
}

function saveDanhMucKhoanThuPhi(token, data) {
  requireSession_(token, 'tuition.write');
  data = data || {};
  const inputId = String(data.maKhoanThu || '').trim().toUpperCase();
  const tenKhoanThu = String(data.tenKhoanThu || '').trim();
  const soTienMacDinh = number_(data.soTienMacDinh);
  const thuTu = Math.max(0, number_(data.thuTu));
  const trangThai = String(data.trangThai || 'ACTIVE').trim().toUpperCase() === 'INACTIVE'
    ? 'INACTIVE'
    : 'ACTIVE';
  if (!tenKhoanThu) throw new Error('Vui lòng nhập tên khoản thu.');
  if (soTienMacDinh < 0) throw new Error('Số tiền mặc định không được nhỏ hơn 0.');

  const sheet = ensureDanhMucKhoanThuPhiSheet_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật danh mục khoản thu. Vui lòng thử lại.');
  let savedId = inputId;
  try {
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    let targetIndex = -1;
    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const rowId = String(values[rowIndex][index.MaKhoanThu] || '').trim().toUpperCase();
      const rowName = String(values[rowIndex][index.TenKhoanThu] || '').trim().toLocaleLowerCase('vi');
      if (inputId && rowId === inputId) targetIndex = rowIndex;
      if (rowName === tenKhoanThu.toLocaleLowerCase('vi') && rowId !== inputId) {
        throw new Error('Tên khoản thu đã tồn tại.');
      }
    }
    savedId = inputId || ('KTP_' + Utilities.getUuid().slice(0, 10).toUpperCase());
    const now = new Date();
    const row = targetIndex >= 0 ? values[targetIndex].slice() : new Array(headers.length).fill('');
    row[index.MaKhoanThu] = savedId;
    row[index.TenKhoanThu] = tenKhoanThu;
    row[index.SoTienMacDinh] = soTienMacDinh;
    row[index.ThuTu] = thuTu;
    row[index.TrangThai] = trangThai;
    row[index.CreatedAt] = targetIndex >= 0 ? (row[index.CreatedAt] || now) : now;
    row[index.UpdatedAt] = now;
    if (targetIndex >= 0) sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  bumpDataVersion_();
  return jsonResponse_({
    success: true,
    maKhoanThu: savedId,
    categories: getDanhMucKhoanThuPhiList_(),
    message: inputId ? 'Đã cập nhật khoản thu.' : 'Đã thêm khoản thu mới.'
  });
}

function parseKhoanThuThem_(value) {
  if (!value) return [];
  try {
    const parsed = Array.isArray(value) ? value : JSON.parse(String(value));
    if (!Array.isArray(parsed)) return [];
    return parsed.map(item => ({
      maKhoanThu: String(item.maKhoanThu || '').trim(),
      tenKhoanThu: String(item.tenKhoanThu || '').trim(),
      soTien: Math.max(0, number_(item.soTien))
    })).filter(item => item.maKhoanThu && item.tenKhoanThu && item.soTien > 0);
  } catch (error) {
    return [];
  }
}

function normalizeKhoanThuThemInput_(value) {
  const input = Array.isArray(value) ? value : parseKhoanThuThem_(value);
  const categoryMap = getDanhMucKhoanThuPhiList_().reduce((map, item) => {
    map[item.maKhoanThu] = item;
    return map;
  }, {});
  const seen = {};
  return input.map(item => {
    const code = String(item.maKhoanThu || '').trim();
    const category = categoryMap[code];
    const amount = number_(item.soTien);
    if (!category) throw new Error('Khoản thu không tồn tại trong danh mục: ' + code + '.');
    if (amount < 0) throw new Error('Số tiền khoản thu không được nhỏ hơn 0.');
    if (seen[code]) throw new Error('Khoản thu bị chọn trùng: ' + category.tenKhoanThu + '.');
    seen[code] = true;
    return { maKhoanThu: code, tenKhoanThu: category.tenKhoanThu, soTien: amount };
  }).filter(item => item.soTien > 0);
}

/**
 * Trả dữ liệu thu phí của một tháng.
 *
 * Nguyên tắc quan trọng:
 * - Lần đầu mở tháng: chụp cố định danh sách học sinh của kỳ học.
 * - Các lần mở sau: chỉ đọc dữ liệu đã chụp, tuyệt đối không đồng bộ lại
 *   theo sheet HocSinh.
 * - Khi thu phí: chỉ cập nhật các cột thu phí của đúng một học sinh.
 */
function getQuanLyThuPhiData(token, yearMonth) {
  const session = requireSession_(token, 'tuition.read');
  const ym = parseYearMonth_(yearMonth);

  const cacheKey = buildCacheKey_(
    'thuphi_snapshot_v5_' + session.maKyHoc + '_' + ym.year + '_' + ym.month
  );

  const cached = cacheGetString_(cacheKey);
  if (cached) return cached;

  const snapshot = ensureThuPhiMonthSnapshot_(
    session.maKyHoc,
    ym.year,
    ym.month
  );

  const sources = getNguonTienList_(session.maKyHoc)
    .filter(item => item.trangThai === 'ACTIVE');
  const feeCategories = getDanhMucKhoanThuPhiList_();

  const rows = snapshot.rows
    .filter(row => {
      return String(row.MaKyHoc || '').trim() === session.maKyHoc &&
        String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
    });

  const studentMetaMap = readObjects_(SHEET_HOCSINH).reduce((map, row) => {
    const id = String(row.MaHocSinh || '').trim();
    if (!id || String(row.TrangThai || '').trim().toUpperCase() === 'DELETED') return map;
    map[id] = {
      truong: String(row.Truong || '').trim() || 'THCS Long Phước',
      gioiTinh: String(row.GioiTinh || '').trim(),
      khongThuPhi: toBoolean_(row.KhongThuPhi)
    };
    return map;
  }, {});

  const resultStudents = rows
    .map(row => {
      const maHocSinh = String(row.MaHocSinh || '').trim();
      const studentMeta = studentMetaMap[maHocSinh] || {};
      const tamNghi = toBoolean_(row.TamNghi) ||
        String(row.TrangThaiThu || '').trim() === 'Tạm nghỉ';
      const rowHocPhi = number_(row.HocPhi);
      const hocPhiGoc = number_(row.HocPhiGoc) || rowHocPhi;
      const khoanThuThem = parseKhoanThuThem_(row.KhoanThuThemJson);
      const tongKhoanThuThem = khoanThuThem.reduce((sum, item) => sum + number_(item.soTien), 0);
      const hasStoredBaseFee = row.HocPhiCoBan !== undefined && String(row.HocPhiCoBan).trim() !== '';
      const hocPhiCoBan = hasStoredBaseFee
        ? number_(row.HocPhiCoBan)
        : Math.max(hocPhiGoc - tongKhoanThuThem, 0);
      const hocPhi = tamNghi ? 0 : rowHocPhi;
      const daThu = tamNghi ? 0 : number_(row.SoTienDaThu);
      const conLai = tamNghi ? 0 : Math.max(hocPhi - daThu, 0);
      const savedSource = String(row.NguonTienThu || '').trim().toUpperCase();
      const inferredSource = savedSource || inferNguonTienFromLegacy_(
        row.HinhThucThu,
        ''
      );

      return {
        maHocSinh: maHocSinh,
        sapXep: number_(row.SapXep),

        hoTen: String(row.HoTen || '').trim(),
        khoi: String(row.Khoi || '').trim(),
        tenKhoi: String(row.TenKhoi || row.Khoi || '').trim(),
        lop: String(row.Lop || '').trim(),
        tenLop: String(row.TenLop || row.Lop || '').trim(),
        truong: String(row.Truong || studentMeta.truong || '').trim(),
        gioiTinh: String(row.GioiTinh || studentMeta.gioiTinh || '').trim(),
        sdtPhuHuynh: String(row.SDTPhuHuynh || '').trim(),

        hocPhiGoc: hocPhiGoc,
        hocPhiCoBan: hocPhiCoBan,
        khoanThuThem: khoanThuThem,
        tongKhoanThuThem: tongKhoanThuThem,
        hocPhi: hocPhi,
        soTienDaThu: daThu,
        conLai: conLai,
        tamNghi: tamNghi,
        trangThai: tamNghi ? 'Tạm nghỉ' : getTrangThaiThuPhi_(hocPhi, daThu),

        soPhieu: String(row.SoPhieu || '').trim(),
        ngayThu: tamNghi ? '' : formatDateForInput_(row.NgayDong),
        ngayThuDisplay: tamNghi ? '' : formatDateDisplay_(row.NgayDong),
        hinhThucThu: tamNghi ? '' : String(row.HinhThucThu || '').trim(),
        maNguonTien: tamNghi ? '' : inferredSource,
        tenNguonTien: tamNghi
          ? ''
          : (getNguonTienName_(inferredSource) || (daThu > 0 ? 'Chưa phân loại' : '')),
        ghiChu: String(row.GhiChu || '').trim()
      };
    })
    .filter(item => {
      const studentMeta = studentMetaMap[item.maHocSinh] || {};
      return !studentMeta.khongThuPhi;
    })
    .sort(compareStudentSort_);

  const totalStudents = resultStudents.length;
  const temporaryLeaveCount = resultStudents.filter(item => item.tamNghi).length;
  const paidCount = resultStudents.filter(item => !item.tamNghi && number_(item.soTienDaThu) > 0).length;
  const fullyPaidCount = resultStudents.filter(item => {
    return !item.tamNghi &&
      number_(item.hocPhi) > 0 &&
      number_(item.soTienDaThu) >= number_(item.hocPhi);
  }).length;

  const totalExpected = resultStudents.reduce((sum, item) => {
    return sum + number_(item.hocPhi);
  }, 0);

  const totalCollected = resultStudents.reduce((sum, item) => {
    return sum + number_(item.soTienDaThu);
  }, 0);

  const totalRemaining = resultStudents.reduce((sum, item) => {
    return sum + number_(item.conLai);
  }, 0);

  const unassignedCount = resultStudents.filter(item => {
    return !item.tamNghi &&
      number_(item.soTienDaThu) > 0 &&
      !item.maNguonTien;
  }).length;

  const json = jsonResponse_({
    session: session,
    month: ym.month,
    year: ym.year,
    monthText: String(ym.month).padStart(2, '0') + '/' + ym.year,
    sheetName: snapshot.sheet.getName(),
    snapshotCreated: snapshot.created,
    snapshotLocked: true,
    sources: sources,
    feeCategories: feeCategories,

    summary: {
      totalStudents: totalStudents,
      temporaryLeaveCount: temporaryLeaveCount,
      paidCount: paidCount,
      fullyPaidCount: fullyPaidCount,
      totalExpected: totalExpected,
      totalCollected: totalCollected,
      totalRemaining: totalRemaining,
      unassignedCount: unassignedCount
    },

    students: resultStudents
  });

  cachePutString_(cacheKey, json, CACHE_SECONDS);

  return json;
}

function getThuPhiQrImage(token, amount, addInfo) {
  requireSession_(token, 'tuition.read');
  const paymentAmount = Math.floor(number_(amount));
  if (paymentAmount <= 0 || paymentAmount > 9999999999999) {
    throw new Error('Số tiền tạo mã QR không hợp lệ.');
  }

  const paymentInfo = String(addInfo || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 50);
  const query = [
    'amount=' + encodeURIComponent(paymentAmount),
    'addInfo=' + encodeURIComponent(paymentInfo || 'DONG HOC PHI'),
    'accountName=' + encodeURIComponent('HO KINH DOANH TRAN NGUYEN LAM')
  ].join('&');
  const url = 'https://img.vietqr.io/image/BIDV-8838077136-qr_only.png?' + query;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error('Không thể tạo mã QR thanh toán. Vui lòng thử lại.');
  }
  const blob = response.getBlob();
  return jsonResponse_({
    dataUrl: 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes()),
    logoDataUrl: getAppBrandLogoDataUrl_()
  });
}

function getAppBrandLogoImage(token) {
  requireSession_(token, 'tuition.read');
  return jsonResponse_({ dataUrl: getAppBrandLogoDataUrl_() });
}

function getAppBrandLogoDataUrl_() {
  const logoUrl = getAppBrandConfig_().logoUrl;
  if (!logoUrl) return '';
  let fetchUrl = logoUrl;
  const driveMatch = logoUrl.match(/(?:\/d\/|[?&]id=)([-\w]{20,})/);
  if (driveMatch) fetchUrl = 'https://drive.google.com/uc?export=download&id=' + driveMatch[1];
  try {
    const response = UrlFetchApp.fetch(fetchUrl, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return '';
    const blob = response.getBlob();
    if (blob.getBytes().length > 2 * 1024 * 1024) return '';
    const contentType = blob.getContentType() || 'image/png';
    if (contentType.indexOf('image/') !== 0) return '';
    return 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (error) {
    return '';
  }
}

/**
 * Ghi nhận hoặc chỉnh sửa thông tin thu phí của một học sinh.
 * Chỉ cập nhật đúng một dòng trong sheet tháng; không ghi lại toàn bộ sheet.
 */
function saveThuPhiHocSinh(token, data) {
  const session = requireSession_(token, 'tuition.write');

  data = data || {};

  const maHocSinh = String(data.maHocSinh || '').trim();
  const yearMonth = String(data.yearMonth || '').trim();
  const cheDoLuu = String(data.cheDoLuu || 'THU_PHI').trim().toUpperCase();
  const chiCapNhatKhoanThu = cheDoLuu === 'CAP_NHAT_KHOAN_THU';
  const hocPhiCoBanInput = number_(data.hocPhiCoBan !== undefined ? data.hocPhiCoBan : data.hocPhi);
  const khoanThuThemInput = normalizeKhoanThuThemInput_(data.khoanThuThem || []);
  const tongKhoanThuThemInput = khoanThuThemInput.reduce((sum, item) => sum + number_(item.soTien), 0);
  const hocPhiInput = hocPhiCoBanInput + tongKhoanThuThemInput;
  const soTienDaThuInput = number_(data.soTienDaThu);
  const tamNghi = toBoolean_(data.tamNghi);
  const ngayThuText = String(data.ngayThu || '').trim();
  const maNguonTien = String(data.maNguonTien || '').trim().toUpperCase();
  const ghiChu = String(data.ghiChu || '').trim();

  if (!maHocSinh) {
    throw new Error('Thiếu mã học sinh.');
  }

  if (!yearMonth) {
    throw new Error('Vui lòng chọn tháng thu phí.');
  }
  assertFinancePeriodOpen_(session, GovernanceService.validateMonth(yearMonth));

  if (!tamNghi && hocPhiCoBanInput < 0) {
    throw new Error('Học phí cơ bản không được nhỏ hơn 0.');
  }

  if (!chiCapNhatKhoanThu && !tamNghi && soTienDaThuInput < 0) {
    throw new Error('Số tiền đã thu không được nhỏ hơn 0.');
  }

  if (!chiCapNhatKhoanThu && !tamNghi && soTienDaThuInput <= 0) {
    throw new Error('Vui lòng nhập số tiền cần ghi nhận thu.');
  }

  if (!chiCapNhatKhoanThu && !tamNghi && soTienDaThuInput > 0 && !ngayThuText) {
    throw new Error('Vui lòng nhập ngày đóng học phí.');
  }

  if (!chiCapNhatKhoanThu && !tamNghi && soTienDaThuInput > 0 && !getNguonTienDefinition_(maNguonTien)) {
    throw new Error('Vui lòng chọn nguồn nhận tiền học phí.');
  }

  const ym = parseYearMonth_(yearMonth);
  const snapshot = ensureThuPhiMonthSnapshot_(
    session.maKyHoc,
    ym.year,
    ym.month,
    { includeRows: false }
  );

  ensureThuChiSheets_(session.maKyHoc);

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang có người cập nhật thu phí. Vui lòng thao tác lại.');
  }

  let savedSoPhieu = '';

  try {
    const sheet = snapshot.sheet;
    const values = sheet.getDataRange().getValues();

    if (!values || values.length < 2) {
      throw new Error('Danh sách thu phí tháng này chưa có học sinh.');
    }

    const headers = values[0].map(header => String(header || '').trim());
    const headerIndex = headers.reduce((map, header, index) => {
      if (header) map[header] = index;
      return map;
    }, {});

    const requiredHeaders = [
      'MaHocSinh',
      'MaKyHoc',
      'HocPhi',
      'HocPhiGoc',
      'HocPhiCoBan',
      'KhoanThuThemJson',
      'TongKhoanThuThem',
      'TamNghi',
      'SoTienDaThu',
      'DaDong',
      'ConLai',
      'SoPhieu',
      'NgayDong',
      'HinhThucThu',
      'NguonTienThu',
      'GhiChu',
      'TrangThaiThu',
      'TrangThai',
      'UpdatedAt'
    ];

    requiredHeaders.forEach(header => {
      if (headerIndex[header] === undefined) {
        throw new Error('Sheet thu phí thiếu cột: ' + header + '.');
      }
    });

    let targetArrayIndex = -1;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowMaHocSinh = String(row[headerIndex.MaHocSinh] || '').trim();
      const rowMaKyHoc = String(row[headerIndex.MaKyHoc] || '').trim();
      const rowTrangThai = String(row[headerIndex.TrangThai] || '').trim().toUpperCase();

      if (
        rowMaHocSinh === maHocSinh &&
        rowMaKyHoc === session.maKyHoc &&
        rowTrangThai !== 'DELETED'
      ) {
        targetArrayIndex = i;
        break;
      }
    }

    if (targetArrayIndex === -1) {
      throw new Error('Không tìm thấy học sinh trong danh sách cố định của tháng này.');
    }

    const row = values[targetArrayIndex].slice();
    const oldHocPhi = number_(row[headerIndex.HocPhi]);
    const oldHocPhiGoc = number_(row[headerIndex.HocPhiGoc]) || oldHocPhi;
    const existingSoPhieu = String(row[headerIndex.SoPhieu] || '').trim();

    if (chiCapNhatKhoanThu && !tamNghi) {
      const soTienDaThuHienTai = number_(row[headerIndex.SoTienDaThu]);
      const conLai = Math.max(hocPhiInput - soTienDaThuHienTai, 0);

      row[headerIndex.HocPhiGoc] = hocPhiInput;
      row[headerIndex.HocPhiCoBan] = hocPhiCoBanInput;
      row[headerIndex.KhoanThuThemJson] = JSON.stringify(khoanThuThemInput);
      row[headerIndex.TongKhoanThuThem] = tongKhoanThuThemInput;
      row[headerIndex.HocPhi] = hocPhiInput;
      row[headerIndex.TamNghi] = 'Không';
      row[headerIndex.ConLai] = conLai;
      row[headerIndex.GhiChu] = ghiChu;
      row[headerIndex.TrangThaiThu] = getTrangThaiThuPhi_(hocPhiInput, soTienDaThuHienTai);
    } else if (tamNghi) {
      row[headerIndex.HocPhiGoc] = hocPhiInput || oldHocPhiGoc;
      row[headerIndex.HocPhiCoBan] = hocPhiCoBanInput;
      row[headerIndex.KhoanThuThemJson] = JSON.stringify(khoanThuThemInput);
      row[headerIndex.TongKhoanThuThem] = tongKhoanThuThemInput;
      row[headerIndex.HocPhi] = 0;
      row[headerIndex.TamNghi] = 'Có';
      row[headerIndex.SoTienDaThu] = 0;
      row[headerIndex.DaDong] = 'Không';
      row[headerIndex.ConLai] = 0;
      row[headerIndex.NgayDong] = '';
      row[headerIndex.HinhThucThu] = '';
      row[headerIndex.NguonTienThu] = '';
      row[headerIndex.GhiChu] = ghiChu;
      row[headerIndex.TrangThaiThu] = 'Tạm nghỉ';
    } else {
      const hocPhi = hocPhiInput;
      const soTienDaThu = soTienDaThuInput;
      const conLai = Math.max(hocPhi - soTienDaThu, 0);

      row[headerIndex.HocPhiGoc] = hocPhi;
      row[headerIndex.HocPhiCoBan] = hocPhiCoBanInput;
      row[headerIndex.KhoanThuThemJson] = JSON.stringify(khoanThuThemInput);
      row[headerIndex.TongKhoanThuThem] = tongKhoanThuThemInput;
      row[headerIndex.HocPhi] = hocPhi;
      row[headerIndex.TamNghi] = 'Không';
      row[headerIndex.SoTienDaThu] = soTienDaThu;
      row[headerIndex.DaDong] = soTienDaThu > 0 ? 'Có' : 'Không';
      row[headerIndex.ConLai] = conLai;
      row[headerIndex.NgayDong] = soTienDaThu > 0 ? parseInputDate_(ngayThuText) : '';
      row[headerIndex.HinhThucThu] = soTienDaThu > 0
        ? getHinhThucByNguon_(maNguonTien)
        : '';
      row[headerIndex.NguonTienThu] = soTienDaThu > 0 ? maNguonTien : '';
      row[headerIndex.GhiChu] = ghiChu;
      row[headerIndex.TrangThaiThu] = getTrangThaiThuPhi_(hocPhi, soTienDaThu);
    }

    row[headerIndex.TrangThai] = 'ACTIVE';
    row[headerIndex.UpdatedAt] = new Date();

    if (chiCapNhatKhoanThu && !tamNghi) {
      savedSoPhieu = existingSoPhieu;
    } else {
      const ledgerResult = upsertThuChiHocPhiNoLock_({
        maKyHoc: session.maKyHoc,
        yearMonth: yearMonth,
        maHocSinh: maHocSinh,
        hoTen: headerIndex.HoTen !== undefined ? String(row[headerIndex.HoTen] || '').trim() : '',
        ngayThu: tamNghi ? '' : row[headerIndex.NgayDong],
        soTien: tamNghi ? 0 : number_(row[headerIndex.SoTienDaThu]),
        maNguonTien: tamNghi ? '' : String(row[headerIndex.NguonTienThu] || '').trim(),
        ghiChu: String(row[headerIndex.GhiChu] || '').trim(),
        tamNghi: tamNghi,
        soPhieu: existingSoPhieu
      });

      savedSoPhieu = ledgerResult && ledgerResult.soPhieu
        ? ledgerResult.soPhieu
        : existingSoPhieu;
    }
    row[headerIndex.SoPhieu] = savedSoPhieu;

    sheet.getRange(targetArrayIndex + 1, 1, 1, headers.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'UPSERT', 'THU_PHI', maHocSinh + '|' + yearMonth, null, {
    cheDoLuu: cheDoLuu, hocPhi: hocPhiInput, soTienDaThu: tamNghi ? 0 : soTienDaThuInput, soPhieu: savedSoPhieu
  });

  return jsonResponse_({
    success: true,
    soPhieu: savedSoPhieu,
    hocPhi: hocPhiInput,
    hocPhiCoBan: hocPhiCoBanInput,
    khoanThuThem: khoanThuThemInput,
    tongKhoanThuThem: tongKhoanThuThemInput,
    message: chiCapNhatKhoanThu && !tamNghi
      ? 'Đã cập nhật các khoản phải thu; chưa ghi nhận thu tiền.'
      : tamNghi
      ? 'Đã cập nhật trạng thái tạm nghỉ.'
      : (soTienDaThuInput > 0
        ? 'Đã cập nhật học phí và phiếu thu ' + (savedSoPhieu || '') + '.'
        : 'Đã cập nhật học phí; chưa phát sinh phiếu thu.')
  });
}

/**
 * Bảo đảm sheet tháng tồn tại và khởi tạo ảnh chụp đúng một lần.
 * Nếu sheet đã có dữ liệu của kỳ học thì chỉ đánh dấu đã khởi tạo,
 * không đồng bộ lại từ HocSinh.
 */
function ensureThuPhiMonthSnapshot_(maKyHoc, year, month, options) {
  const includeRows = !options || options.includeRows !== false;
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang khởi tạo dữ liệu tháng. Vui lòng thao tác lại.');
  }

  try {
    const sheet = ensureThuPhiMonthSheet_(maKyHoc, year, month);
    const sheetName = sheet.getName();
    const propertyKey = getThuPhiSnapshotPropertyKey_(maKyHoc, year, month);
    const props = PropertiesService.getScriptProperties();
    const marker = props.getProperty(propertyKey);
    const markerSheetId = marker ? String(marker).split('|')[0] : '';
    const markerMatchesSheet = markerSheetId === String(sheet.getSheetId());
    if (markerMatchesSheet) {
      return {
        sheet: sheet,
        created: false,
        rows: includeRows ? readObjectsNoCache_(sheetName) : []
      };
    }

    const rows = readObjectsNoCache_(sheetName);

    const hasExistingSnapshot = rows.some(row => {
      return String(row.MaKyHoc || '').trim() === maKyHoc &&
        String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
    });

    if (hasExistingSnapshot) {
      props.setProperty(
        propertyKey,
        String(sheet.getSheetId()) + '|' + String(Date.now())
      );

      return {
        sheet: sheet,
        created: false,
        rows: rows
      };
    }

    const now = new Date();
    const students = getHocSinhTheoKyHocForThuPhi_(maKyHoc);

    const snapshotRows = students.map(student => {
      const hocPhi = number_(student.hocPhi || defaultHocPhiByKhoi_(student.khoi));

      return {
        MaHocSinh: student.maHocSinh,
        MaKyHoc: maKyHoc,
        SapXep: student.sapXep,
        Khoi: student.khoi,
        TenKhoi: student.tenKhoi,
        Lop: student.lop,
        TenLop: student.tenLop,
        HoTen: student.hoTen,
        Truong: student.truong,
        GioiTinh: student.gioiTinh,
        SDTPhuHuynh: student.sdtPhuHuynh,
        NgayVao: toDateOnly_(student.ngayVaoRaw) || student.ngayVaoRaw || '',
        HocPhi: hocPhi,
        HocPhiGoc: hocPhi,
        HocPhiCoBan: hocPhi,
        KhoanThuThemJson: '[]',
        TongKhoanThuThem: 0,
        TamNghi: 'Không',
        SoTienDaThu: 0,
        DaDong: 'Không',
        ConLai: hocPhi,
        NgayDong: '',
        HinhThucThu: '',
        GhiChu: '',
        TrangThaiThu: 'Chưa thu',
        TrangThai: 'ACTIVE',
        CreatedAt: now,
        UpdatedAt: now
      };
    });

    appendObjectsToSheet_(sheet, snapshotRows, getThuPhiMonthHeaders_());

    // Marker vẫn được lưu ngay cả khi tháng được tạo lúc chưa có học sinh.
    // Nhờ đó học sinh thêm về sau cũng không tự chèn vào tháng đã chốt.
    props.setProperty(
      propertyKey,
      String(sheet.getSheetId()) + '|' + String(Date.now())
    );

    return {
      sheet: sheet,
      created: true,
      rows: rows.concat(snapshotRows)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Chủ động bổ sung một học sinh mới vào sheet thu phí tháng hiện tại.
 * Chỉ được gọi khi người dùng tick “Cập nhật thu phí” lúc thêm học sinh.
 */
function syncHocSinhToThuPhiMonth_(maKyHoc, yearMonth, student, hocPhiValue) {
  const ym = parseYearMonth_(yearMonth);
  const sheetName = getThuPhiMonthSheetName_(ym.year, ym.month);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) {
    return { updated: false, sheetName: sheetName };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang cập nhật sheet thu phí tháng. Vui lòng thao tác lại.');
  }

  try {
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    const requiredHeaders = [
      'MaHocSinh', 'MaKyHoc', 'SapXep', 'Khoi', 'TenKhoi', 'Lop', 'TenLop',
      'HoTen', 'Truong', 'GioiTinh', 'SDTPhuHuynh', 'NgayVao', 'HocPhi',
      'HocPhiGoc', 'TamNghi', 'SoTienDaThu', 'DaDong', 'ConLai',
      'TrangThaiThu', 'TrangThai', 'UpdatedAt'
    ];
    requiredHeaders.forEach(header => {
      if (index[header] === undefined) {
        throw new Error('Sheet ' + sheetName + ' thiếu cột: ' + header + '.');
      }
    });

    const studentId = String(student.MaHocSinh || '').trim();
    let targetIndex = -1;
    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      if (
        String(values[rowIndex][index.MaHocSinh] || '').trim() === studentId &&
        String(values[rowIndex][index.MaKyHoc] || '').trim() === maKyHoc
      ) {
        targetIndex = rowIndex;
        break;
      }
    }

    if (targetIndex === -1) {
      return { updated: false, sheetName: sheetName };
    }

    const lopInfo = getLopList_().find(item => item.maLop === String(student.Lop || '').trim());
    const khoiInfo = getKhoiList_().find(item => item.khoi === String(student.Khoi || '').trim());
    const row = values[targetIndex].slice();
    const khongThuPhi = toBoolean_(student.KhongThuPhi);

    if (khongThuPhi) {
      row[index.HocPhi] = 0;
      row[index.HocPhiGoc] = 0;
      row[index.ConLai] = 0;
      row[index.TrangThaiThu] = 'Không thu phí';
      row[index.TrangThai] = 'DELETED';
      row[index.UpdatedAt] = new Date();
      sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
      return { updated: true, removed: true, sheetName: sheetName };
    }

    const newHocPhi = number_(hocPhiValue);
    const soTienDaThu = number_(row[index.SoTienDaThu]);
    const tamNghi = toBoolean_(row[index.TamNghi]);
    const savedExtraItems = index.KhoanThuThemJson !== undefined
      ? parseKhoanThuThem_(row[index.KhoanThuThemJson])
      : [];
    const savedExtraTotal = savedExtraItems.reduce((sum, item) => sum + number_(item.soTien), 0);
    const newTotalFee = newHocPhi + savedExtraTotal;

    row[index.SapXep] = number_(student.SapXep);
    row[index.Khoi] = String(student.Khoi || '').trim();
    row[index.TenKhoi] = khoiInfo ? khoiInfo.tenKhoi : String(student.Khoi || '').trim();
    row[index.Lop] = String(student.Lop || '').trim();
    row[index.TenLop] = lopInfo ? lopInfo.tenLop : String(student.Lop || '').trim();
    row[index.HoTen] = String(student.HoTen || '').trim();
    row[index.Truong] = String(student.Truong || '').trim();
    row[index.GioiTinh] = String(student.GioiTinh || '').trim();
    row[index.SDTPhuHuynh] = String(student.SDTPhuHuynh || '').trim();
    row[index.NgayVao] = student.NgayVao || student.NgaySinh || '';
    row[index.HocPhiGoc] = newTotalFee;
    row[index.HocPhi] = tamNghi ? 0 : newTotalFee;
    if (index.HocPhiCoBan !== undefined) row[index.HocPhiCoBan] = newHocPhi;
    if (index.TongKhoanThuThem !== undefined) row[index.TongKhoanThuThem] = savedExtraTotal;
    row[index.ConLai] = tamNghi ? 0 : Math.max(newTotalFee - soTienDaThu, 0);
    row[index.DaDong] = !tamNghi && soTienDaThu > 0 ? 'Có' : 'Không';
    row[index.TrangThaiThu] = tamNghi
      ? 'Tạm nghỉ'
      : getTrangThaiThuPhi_(newTotalFee, soTienDaThu);
    row[index.TrangThai] = 'ACTIVE';
    row[index.UpdatedAt] = new Date();

    sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
    return { updated: true, sheetName: sheetName };
  } finally {
    lock.releaseLock();
  }
}

function addHocSinhToThuPhiMonth_(maKyHoc, yearMonth, maHocSinh) {
  const ym = parseYearMonth_(yearMonth);
  const snapshot = ensureThuPhiMonthSnapshot_(maKyHoc, ym.year, ym.month, { includeRows: false });
  const sheet = snapshot.sheet;
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang cập nhật danh sách thu phí. Vui lòng thao tác lại.');
  }

  try {
    const values = sheet.getDataRange().getValues();
    const headers = values && values.length
      ? values[0].map(header => String(header || '').trim())
      : getThuPhiMonthHeaders_();

    const headerIndex = headers.reduce((map, header, index) => {
      if (header) map[header] = index;
      return map;
    }, {});

    const alreadyExists = values.slice(1).some(row => {
      return String(row[headerIndex.MaHocSinh] || '').trim() === maHocSinh &&
        String(row[headerIndex.MaKyHoc] || '').trim() === maKyHoc &&
        String(row[headerIndex.TrangThai] || '').trim().toUpperCase() !== 'DELETED';
    });

    if (alreadyExists) {
      return {
        added: false,
        alreadyExists: true,
        sheetName: sheet.getName()
      };
    }

    const student = getHocSinhTheoKyHocForThuPhi_(maKyHoc)
      .find(item => item.maHocSinh === maHocSinh);

    if (!student) {
      throw new Error('Không tìm thấy học sinh vừa thêm trong kỳ học hiện tại.');
    }

    const hocPhi = number_(student.hocPhi || defaultHocPhiByKhoi_(student.khoi));
    const now = new Date();

    appendObjectsToSheet_(sheet, [{
      MaHocSinh: student.maHocSinh,
      MaKyHoc: maKyHoc,
      SapXep: student.sapXep,
      Khoi: student.khoi,
      TenKhoi: student.tenKhoi,
      Lop: student.lop,
      TenLop: student.tenLop,
      HoTen: student.hoTen,
      Truong: student.truong,
      GioiTinh: student.gioiTinh,
      SDTPhuHuynh: student.sdtPhuHuynh,
      NgayVao: toDateOnly_(student.ngayVaoRaw) || student.ngayVaoRaw || '',
      HocPhi: hocPhi,
      HocPhiGoc: hocPhi,
      HocPhiCoBan: hocPhi,
      KhoanThuThemJson: '[]',
      TongKhoanThuThem: 0,
      TamNghi: 'Không',
      SoTienDaThu: 0,
      DaDong: 'Không',
      ConLai: hocPhi,
      NgayDong: '',
      HinhThucThu: '',
      GhiChu: '',
      TrangThaiThu: 'Chưa thu',
      TrangThai: 'ACTIVE',
      CreatedAt: now,
      UpdatedAt: now
    }], getThuPhiMonthHeaders_());

    return {
      added: true,
      alreadyExists: false,
      sheetName: sheet.getName()
    };
  } finally {
    lock.releaseLock();
  }
}

function getThuPhiSnapshotPropertyKey_(maKyHoc, year, month) {
  return 'THUPHI_SNAPSHOT_' +
    String(year) + '_' +
    String(month).padStart(2, '0') + '_' +
    hashString_(maKyHoc);
}

/**
 * Tạo sheet tháng nếu chưa có và chỉ bổ sung header khi thật sự thiếu.
 */
function ensureThuPhiMonthSheet_(maKyHoc, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getThuPhiMonthSheetName_(year, month);
  const requiredHeaders = getThuPhiMonthHeaders_();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    formatHeader_(sheet);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn <= 0 || sheet.getLastRow() <= 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    formatHeader_(sheet);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const currentHeaders = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(header => String(header || '').trim())
    .filter(header => header);

  const finalHeaders = currentHeaders.slice();
  let changed = false;

  requiredHeaders.forEach(header => {
    if (finalHeaders.indexOf(header) === -1) {
      finalHeaders.push(header);
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
    formatHeader_(sheet);
  }

  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Ghi thêm nhiều dòng trong một lần setValues().
 */
function appendObjectsToSheet_(sheet, objects, requiredHeaders) {
  if (!objects || objects.length === 0) return;

  const lastColumn = sheet.getLastColumn();
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(header => String(header || '').trim());

  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      throw new Error('Sheet ' + sheet.getName() + ' thiếu cột: ' + header + '.');
    }
  });

  const values = objects.map(obj => {
    return headers.map(header => {
      return Object.prototype.hasOwnProperty.call(obj, header) ? obj[header] : '';
    });
  });

  sheet
    .getRange(sheet.getLastRow() + 1, 1, values.length, headers.length)
    .setValues(values);
}

function getHocSinhTheoKyHocForThuPhi_(maKyHoc) {
  const hocSinhRows = readObjects_(SHEET_HOCSINH);
  const relationRows = readObjects_(SHEET_HOCSINH_KYHOC);
  const lopRows = readObjects_(SHEET_LOP);
  const khoiRows = readObjects_(SHEET_KHOI);

  const lopNameMap = lopRows.reduce((map, row) => {
    map[String(row.MaLop || '').trim()] = String(row.TenLop || '').trim();
    return map;
  }, {});

  const khoiNameMap = khoiRows.reduce((map, row) => {
    map[String(row.Khoi || '').trim()] = String(row.TenKhoi || '').trim();
    return map;
  }, {});

  const kyHocMap = relationRows
    .filter(row => {
      return String(row.MaKyHoc || '').trim() === maKyHoc &&
        String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
    })
    .reduce((map, row) => {
      const maHocSinh = String(row.MaHocSinh || '').trim();

      if (maHocSinh) {
        map[maHocSinh] = {
          hocPhi: row.HocPhi
        };
      }

      return map;
    }, {});

  const hasKyHocMapping = Object.keys(kyHocMap).length > 0;

  return hocSinhRows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED')
    .filter(row => !toBoolean_(row.KhongThuPhi))
    .map(row => {
      const maHocSinh = String(row.MaHocSinh || '').trim();

      if (!maHocSinh) return null;

      const rowMaKyHoc = String(row.MaKyHoc || '').trim();

      if (hasKyHocMapping && !kyHocMap[maHocSinh]) return null;
      if (!hasKyHocMapping && rowMaKyHoc && rowMaKyHoc !== maKyHoc) return null;

      const khoi = String(row.Khoi || '').trim();
      const lop = String(row.Lop || '').trim();

      return {
        maHocSinh: maHocSinh,
        sapXep: number_(row.SapXep),
        hoTen: String(row.HoTen || '').trim(),
        khoi: khoi,
        tenKhoi: khoiNameMap[khoi] || khoi,
        lop: lop,
        tenLop: lopNameMap[lop] || lop,
        truong: String(row.Truong || '').trim() || 'THCS Long Phước',
        gioiTinh: String(row.GioiTinh || '').trim(),
        sdtPhuHuynh: String(row.SDTPhuHuynh || '').trim(),
        hocPhi: kyHocMap[maHocSinh] ? kyHocMap[maHocSinh].hocPhi : '',
        ngayVaoRaw: row.NgayVao || row.NgaySinh || row.CreatedAt || '',
        createdAt: row.CreatedAt || ''
      };
    })
    .filter(item => item !== null)
    .sort(compareStudentSort_);
}

function getThuPhiMonthSheetName_(year, month) {
  return 'Thang' + String(month).padStart(2, '0') + '.' + year;
}

function getThuPhiMonthHeaders_() {
  return [
    'MaHocSinh',
    'MaKyHoc',
    'SapXep',
    'Khoi',
    'TenKhoi',
    'Lop',
    'TenLop',
    'HoTen',
    'Truong',
    'GioiTinh',
    'SDTPhuHuynh',
    'NgayVao',
    'HocPhi',
    'HocPhiGoc',
    'HocPhiCoBan',
    'KhoanThuThemJson',
    'TongKhoanThuThem',
    'TamNghi',
    'SoTienDaThu',
    'DaDong',
    'ConLai',
    'SoPhieu',
    'NgayDong',
    'HinhThucThu',
    'NguonTienThu',
    'GhiChu',
    'TrangThaiThu',
    'TrangThai',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function getTrangThaiThuPhi_(hocPhi, soTienDaThu) {
  const hp = number_(hocPhi);
  const daThu = number_(soTienDaThu);

  if (daThu <= 0) return 'Chưa thu';
  if (hp > 0 && daThu < hp) return 'Thu một phần';
  if (hp > 0 && daThu === hp) return 'Đã thu đủ';
  if (hp > 0 && daThu > hp) return 'Thu dư';
  if (hp === 0 && daThu > 0) return 'Thu dư';

  return 'Chưa thu';
}

/* =========================================================
   QUẢN LÝ THU CHI
========================================================= */

function getDanhMucThuChiHeaders_() {
  return [
    'MaDanhMuc',
    'Loai',
    'TenDanhMuc',
    'MaHuMacDinh',
    'ThuTu',
    'TrangThai',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function getSoThuChiHeaders_() {
  return [
    'MaGiaoDich',
    'NgayGiaoDich',
    'MaKyHoc',
    'LoaiGiaoDich',
    'MaDanhMuc',
    'MaKeHoachChi',
    'MaNhanSu',
    'MaHuTaiChinh',
    'TenDanhMuc',
    'NoiDung',
    'SoTien',
    'HinhThuc',
    'MaNguonTien',
    'TenNguonTien',
    'MaNguonDoiUng',
    'MaNhomChuyen',
    'NguoiNopNhan',
    'SoPhieu',
    'SoChungTu',
    'ChungTuFileId',
    'ChungTuUrl',
    'GhiChu',
    'NguonDuLieu',
    'MaThamChieu',
    'PhamVi',
    'MaDanhMucGiaDinh',
    'LoaiGiaDinh',
    'MaGiaoDichLienKet',
    'TrangThai',
    'NguoiTao',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function getDoiTuongThuChiHeaders_() {
  return [
    'MaDoiTuong', 'Loai', 'TenDoiTuong', 'SoDienThoai', 'DiaChi',
    'GhiChu', 'TrangThai', 'CreatedAt', 'UpdatedAt'
  ];
}

function getKeHoachChiThangHeaders_() {
  return [
    'MaKeHoachChi', 'MaKyHoc', 'Thang', 'MaDanhMuc', 'TenKhoanChi',
    'NhomChi', 'MaNhanSu', 'NguoiNhan', 'SoTienPhaiChi', 'HanThanhToan', 'BatBuoc',
    'NguonKeHoach', 'MaThamChieu', 'GhiChu', 'TrangThai', 'CreatedAt', 'UpdatedAt'
  ];
}

function getCauHinhTaiChinhThangHeaders_() {
  return [
    'MaKyHoc', 'Thang', 'SoHocSinh', 'DoanhThuDuKien', 'TyLeThue',
    'NguongLuong', 'NguongBanTru', 'NguongTongChi', 'TyLeDuPhong', 'TyLeLoiNhuanMucTieu',
    'GhiChu', 'CreatedAt', 'UpdatedAt'
  ];
}

function getNhanSuTaiChinhHeaders_() {
  return [
    'MaNhanSu', 'MaKyHoc', 'HoTen', 'VaiTro', 'MaDanhMuc', 'NhomChi',
    'MucChiMacDinh', 'NgayThanhToan', 'LopPhuTrach', 'TuNgay', 'DenNgay',
    'TrangThai', 'GhiChu', 'CreatedAt', 'UpdatedAt'
  ];
}

function getKhoanChiDinhKyHeaders_() {
  return [
    'MaKhoanDinhKy', 'MaKyHoc', 'TenKhoanChi', 'MaDanhMuc', 'NhomChi',
    'MaNhanSu', 'LoaiKhoanNhanSu', 'PhuongPhapTinh', 'DinhMuc', 'NgayThanhToan', 'BatBuoc', 'TuThang',
    'DenThang', 'TrangThai', 'GhiChu', 'CreatedAt', 'UpdatedAt'
  ];
}

function getHuTaiChinhThangHeaders_() {
  return [
    'MaKyHoc', 'Thang', 'MaHu', 'TenHu', 'TyLePhanBo', 'ThuTu',
    'GhiChu', 'CreatedAt', 'UpdatedAt', 'VaiTroHeThong'
  ];
}

function getDanhMucHuTaiChinhHeaders_() {
  return [
    'MaHu', 'TenHu', 'TyLeMacDinh', 'ThuTu', 'TrangThai', 'GhiChu',
    'CreatedAt', 'UpdatedAt', 'VaiTroHeThong'
  ];
}

function getChotPhanBoHuHeaders_() {
  return [
    'MaChot', 'MaKyHoc', 'Thang', 'TrangThai', 'HocPhiCoSo', 'NgayChot', 'NguoiChot',
    'NgayMoKhoa', 'NguoiMoKhoa', 'LyDoMoKhoa', 'PhienBan', 'CreatedAt', 'UpdatedAt'
  ];
}

function getDanhMucGiaDinhHeaders_() {
  return [
    'MaDanhMucGiaDinh', 'TenDanhMuc', 'Loai', 'Nhom', 'TyLeMacDinh',
    'KieuGioiHan', 'ThuTu', 'TrangThai', 'GhiChu', 'CreatedAt', 'UpdatedAt'
  ];
}

function getGiaoDichGiaDinhHeaders_() {
  return [
    'MaGiaoDichGiaDinh', 'NgayGiaoDich', 'Thang', 'Loai', 'MaDanhMucGiaDinh',
    'TenDanhMuc', 'NoiDung', 'SoTien', 'NguonTien', 'GhiChu', 'TrangThai',
    'CreatedAt', 'UpdatedAt', 'MaKyHoc', 'MaGiaoDichSoThuChi'
  ];
}

function getCauHinhGiaDinhThangHeaders_() {
  return [
    'Thang', 'ThuNhapDuKien', 'QuyKhanCapMucTieu', 'GhiChu', 'CreatedAt', 'UpdatedAt',
    'ThuNhapKhacDuKien'
  ];
}

function getNguonTienHeaders_() {
  return [
    'MaKyHoc',
    'MaNguonTien',
    'TenNguonTien',
    'PhamVi',
    'ThuTu',
    'SoDuBanDau',
    'TrangThai',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function normalizeFinanceScope_(value, fallback) {
  const scope = String(value || fallback || 'TRUNG_TAM').trim().toUpperCase();
  return scope === 'GIA_DINH' ? 'GIA_DINH' : (scope === 'ALL' ? 'ALL' : 'TRUNG_TAM');
}

function getDefaultNguonTien_(scope) {
  const items = [
    ['VIETCOMBANK', 'Vietcombank', 1, 'TRUNG_TAM'],
    ['VIETINBANK', 'VietinBank', 2, 'TRUNG_TAM'],
    ['BIDV', 'BIDV', 3, 'TRUNG_TAM'],
    ['TIEN_MAT', 'Tiền mặt', 4, 'TRUNG_TAM'],
    ['KET', 'Két', 5, 'TRUNG_TAM'],
    ['GD_TIEN_MAT', 'Tiền mặt gia đình', 101, 'GIA_DINH'],
    ['GD_BIDV', 'BIDV cá nhân', 102, 'GIA_DINH'],
    ['GD_VIETCOMBANK', 'Vietcombank cá nhân', 103, 'GIA_DINH'],
    ['GD_VIETINBANK', 'VietinBank cá nhân', 104, 'GIA_DINH'],
    ['GD_TIET_KIEM', 'Tài khoản tiết kiệm gia đình', 105, 'GIA_DINH']
  ];
  const normalizedScope = normalizeFinanceScope_(scope || 'ALL', 'ALL');
  return normalizedScope === 'ALL' ? items : items.filter(item => item[3] === normalizedScope);
}

function getNguonTienDefinition_(maNguonTien) {
  const code = String(maNguonTien || '').trim().toUpperCase();
  const found = getDefaultNguonTien_().find(item => item[0] === code);

  return found
    ? { maNguonTien: found[0], tenNguonTien: found[1], thuTu: found[2], phamVi: found[3] }
    : null;
}

function getNguonTienName_(maNguonTien) {
  const found = getNguonTienDefinition_(maNguonTien);
  return found ? found.tenNguonTien : '';
}

function getHinhThucByNguon_(maNguonTien) {
  const code = String(maNguonTien || '').trim().toUpperCase();
  return code === 'VIETCOMBANK' || code === 'VIETINBANK' || code === 'BIDV' ||
    code === 'GD_VIETCOMBANK' || code === 'GD_VIETINBANK' || code === 'GD_BIDV' || code === 'GD_TIET_KIEM'
    ? 'Chuyển khoản'
    : (code === 'TIEN_MAT' || code === 'KET' || code === 'GD_TIEN_MAT' ? 'Tiền mặt' : '');
}


/**
 * Tạo tiền tố số phiếu theo loại nghiệp vụ và tháng chứng từ.
 * - Phiếu thu nhập tay: PT-YYYYMM-0001
 * - Phiếu chi nhập tay: PC-YYYYMM-0001
 * - Biên lai học phí:   PT-HP-YYYYMM-0001
 */
function buildSoPhieuPrefix_(loai, dateValue) {
  const type = String(loai || '').trim().toUpperCase();
  const date = toDateOnly_(dateValue) || new Date();
  const yearMonth = Utilities.formatDate(date, 'Asia/Ho_Chi_Minh', 'yyyyMM');

  if (type === 'HOC_PHI') return 'PT-HP-' + yearMonth + '-';
  if (type === 'THU_GIA_DINH') return 'PT-GD-' + yearMonth + '-';
  if (type === 'CHI_GIA_DINH') return 'PC-GD-' + yearMonth + '-';
  return (type === 'CHI' ? 'PC-' : 'PT-') + yearMonth + '-';
}

/**
 * Sinh số phiếu kế tiếp từ dữ liệu sổ thu chi đã đọc trên RAM.
 * Hàm này phải được gọi bên trong LockService khi ghi dữ liệu.
 */
function generateNextSoPhieuFromRows_(loai, dateValue, rows, index) {
  const prefix = buildSoPhieuPrefix_(loai, dateValue);
  const safePrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp('^' + safePrefix + '(\\d+)$', 'i');
  let maxNumber = 0;

  (rows || []).forEach(row => {
    const value = index && index.SoPhieu !== undefined
      ? String(row[index.SoPhieu] || '').trim()
      : '';
    const match = value.match(regex);
    if (!match) return;
    maxNumber = Math.max(maxNumber, Number(match[1]) || 0);
  });

  return prefix + String(maxNumber + 1).padStart(4, '0');
}

function inferNguonTienFromLegacy_(hinhThuc, tenNguonTien) {
  const text = normalizeText_(
    String(tenNguonTien || '') + ' ' + String(hinhThuc || '')
  );

  if (text.indexOf('vietcombank') !== -1 || text.indexOf('vcb') !== -1) return 'VIETCOMBANK';
  if (text.indexOf('vietinbank') !== -1 || text.indexOf('viettinbank') !== -1 || text.indexOf('ctg') !== -1) return 'VIETINBANK';
  if (text.indexOf('bidv') !== -1) return 'BIDV';
  if (text.indexOf('ket') !== -1) return 'KET';
  if (text.indexOf('tien mat') !== -1 || text === 'cash') return 'TIEN_MAT';

  return '';
}

function getNguonTienList_(maKyHoc, scope) {
  const rows = readObjectsNoCache_(SHEET_NGUONTIEN);
  const normalizedScope = normalizeFinanceScope_(scope || 'TRUNG_TAM');

  return rows
    .filter(row => {
      return String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
        String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED';
    })
    .map(row => ({
      maNguonTien: String(row.MaNguonTien || '').trim().toUpperCase(),
      tenNguonTien: String(row.TenNguonTien || '').trim(),
      phamVi: normalizeFinanceScope_(row.PhamVi || (String(row.MaNguonTien || '').trim().toUpperCase().indexOf('GD_') === 0 ? 'GIA_DINH' : 'TRUNG_TAM')),
      thuTu: number_(row.ThuTu) || 999,
      soDuBanDau: number_(row.SoDuBanDau),
      trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase()
    }))
    .filter(item => !!getNguonTienDefinition_(item.maNguonTien))
    .filter(item => normalizedScope === 'ALL' || item.phamVi === normalizedScope)
    .sort((a, b) => a.thuTu - b.thuTu);
}

function inferFamilySourceFromLegacy_(value) {
  const text = normalizeText_(value || '');
  if (text.indexOf('vietcombank') !== -1 || text.indexOf('vcb') !== -1) return 'GD_VIETCOMBANK';
  if (text.indexOf('vietinbank') !== -1 || text.indexOf('ctg') !== -1) return 'GD_VIETINBANK';
  if (text.indexOf('bidv') !== -1) return 'GD_BIDV';
  if (text.indexOf('tiet kiem') !== -1) return 'GD_TIET_KIEM';
  return 'GD_TIEN_MAT';
}

function isLegacyOwnerWithdrawalMatch_(legacy, center) {
  if (!legacy || !center) return false;
  const legacyContent = normalizeText_(legacy.noiDung || '');
  const looksLikeOwnerWithdrawal = legacyContent.indexOf('rut tien') !== -1 && (
    legacyContent.indexOf('luong') !== -1 ||
    legacyContent.indexOf('chu so huu') !== -1 ||
    legacyContent.indexOf('trung tam') !== -1
  );
  return String(legacy.loai || '').trim().toUpperCase() === 'CHI' &&
    looksLikeOwnerWithdrawal &&
    String(center.loai || '').trim().toUpperCase() === 'CHI' &&
    String(center.maDanhMuc || '').trim().toUpperCase() === 'CHI_GIA_DINH' &&
    normalizeFinanceScope_(center.phamVi) === 'TRUNG_TAM' &&
    String(center.trangThai || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG' &&
    String(legacy.ngayGiaoDich || '').trim() === String(center.ngayGiaoDich || '').trim() &&
    Math.abs(number_(legacy.soTien) - number_(center.soTien)) < 1;
}

function findMatchingCenterOwnerWithdrawal_(ledgerValues, ledgerIndex, maKyHoc, legacy, excludedRowIndex) {
  for (let i = 1; i < ledgerValues.length; i++) {
    if (i === excludedRowIndex) continue;
    const row = ledgerValues[i];
    if (String(row[ledgerIndex.MaKyHoc] || '').trim() !== String(maKyHoc || '').trim()) continue;
    const sourceType = String(row[ledgerIndex.NguonDuLieu] || '').trim().toUpperCase();
    const center = {
      maGiaoDich: String(row[ledgerIndex.MaGiaoDich] || '').trim(),
      ngayGiaoDich: formatDateForInput_(toDateOnly_(row[ledgerIndex.NgayGiaoDich])),
      loai: String(row[ledgerIndex.LoaiGiaoDich] || '').trim().toUpperCase(),
      maDanhMuc: String(row[ledgerIndex.MaDanhMuc] || '').trim().toUpperCase(),
      soTien: number_(row[ledgerIndex.SoTien]),
      phamVi: normalizeFinanceScope_(row[ledgerIndex.PhamVi] || (sourceType.indexOf('GIA_DINH') === 0 ? 'GIA_DINH' : 'TRUNG_TAM')),
      trangThai: String(row[ledgerIndex.TrangThai] || 'HOAT_DONG').trim().toUpperCase(),
      soPhieu: String(row[ledgerIndex.SoPhieu] || '').trim()
    };
    if (center.maGiaoDich && isLegacyOwnerWithdrawalMatch_(legacy, center)) return center;
  }
  return null;
}

function migrateLegacyFamilyTransactionsNoLock_(maKyHoc, ledgerSheet) {
  if (!maKyHoc) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const familySheet = ensureSheet_(ss, SHEET_GIAODICH_GIADINH, getGiaoDichGiaDinhHeaders_());
  const familyValues = familySheet.getDataRange().getValues();
  if (familyValues.length <= 1) return 0;
  const familyHeaders = familyValues[0].map(header => String(header || '').trim());
  const familyIndex = buildHeaderIndex_(familyHeaders);
  const ledgerValues = ledgerSheet.getDataRange().getValues();
  const ledgerHeaders = ledgerValues[0].map(header => String(header || '').trim());
  const ledgerIndex = buildHeaderIndex_(ledgerHeaders);
  const newRows = [];
  let familyChanged = false;
  const now = new Date();
  for (let i = 1; i < familyValues.length; i++) {
    const legacyId = String(familyValues[i][familyIndex.MaGiaoDichGiaDinh] || '').trim();
    const linkedId = String(familyValues[i][familyIndex.MaGiaoDichSoThuChi] || '').trim();
    const rowKyHoc = String(familyValues[i][familyIndex.MaKyHoc] || '').trim();
    if (!legacyId || linkedId || (rowKyHoc && rowKyHoc !== maKyHoc)) continue;
    const date = toDateOnly_(familyValues[i][familyIndex.NgayGiaoDich]);
    if (!date) continue;
    const originalFamilyType = String(familyValues[i][familyIndex.Loai] || 'CHI').trim().toUpperCase();
    const legacy = {
      ngayGiaoDich: formatDateForInput_(date),
      loai: originalFamilyType,
      noiDung: String(familyValues[i][familyIndex.NoiDung] || '').trim(),
      soTien: number_(familyValues[i][familyIndex.SoTien])
    };
    const centerWithdrawal = findMatchingCenterOwnerWithdrawal_(ledgerValues, ledgerIndex, maKyHoc, legacy, -1);
    const existingLinkedFamily = centerWithdrawal ? ledgerValues.slice(1).find(row =>
      String(row[ledgerIndex.MaKyHoc] || '').trim() === String(maKyHoc || '').trim() &&
      String(row[ledgerIndex.NguonDuLieu] || '').trim().toUpperCase() === 'GIA_DINH_RUT_CHU' &&
      String(row[ledgerIndex.MaGiaoDichLienKet] || '').trim() === centerWithdrawal.maGiaoDich
    ) : null;
    if (existingLinkedFamily) {
      familyValues[i][familyIndex.MaKyHoc] = maKyHoc;
      familyValues[i][familyIndex.MaGiaoDichSoThuChi] = String(existingLinkedFamily[ledgerIndex.MaGiaoDich] || '').trim();
      familyValues[i][familyIndex.UpdatedAt] = now;
      familyChanged = true;
      continue;
    }
    const familyType = centerWithdrawal ? 'THU' : originalFamilyType;
    const ledgerType = familyType === 'THU' ? 'THU' : 'CHI';
    const sourceCode = inferFamilySourceFromLegacy_(familyValues[i][familyIndex.NguonTien]);
    const source = getNguonTienDefinition_(sourceCode);
    const id = 'TC_GD_' + Utilities.getUuid().slice(0, 10).toUpperCase();
    const allReceiptRows = ledgerValues.slice(1).concat(newRows);
    const receipt = generateNextSoPhieuFromRows_(ledgerType === 'THU' ? 'THU_GIA_DINH' : 'CHI_GIA_DINH', date, allReceiptRows, ledgerIndex);
    const row = new Array(ledgerHeaders.length).fill('');
    row[ledgerIndex.MaGiaoDich] = id;
    row[ledgerIndex.NgayGiaoDich] = date;
    row[ledgerIndex.MaKyHoc] = maKyHoc;
    row[ledgerIndex.LoaiGiaoDich] = ledgerType;
    row[ledgerIndex.MaDanhMuc] = centerWithdrawal ? 'GD_THU_CHU_SO_HUU' : String(familyValues[i][familyIndex.MaDanhMucGiaDinh] || '').trim();
    row[ledgerIndex.TenDanhMuc] = centerWithdrawal ? 'Nhận tiền từ trung tâm' : String(familyValues[i][familyIndex.TenDanhMuc] || '').trim();
    row[ledgerIndex.NoiDung] = centerWithdrawal ? ('Nhận tiền từ trung tâm - ' + legacy.noiDung) : legacy.noiDung;
    row[ledgerIndex.SoTien] = number_(familyValues[i][familyIndex.SoTien]);
    row[ledgerIndex.HinhThuc] = getHinhThucByNguon_(sourceCode);
    row[ledgerIndex.MaNguonTien] = sourceCode;
    row[ledgerIndex.TenNguonTien] = source ? source.tenNguonTien : '';
    row[ledgerIndex.SoPhieu] = receipt;
    row[ledgerIndex.SoChungTu] = centerWithdrawal ? centerWithdrawal.soPhieu : '';
    row[ledgerIndex.GhiChu] = String(familyValues[i][familyIndex.GhiChu] || '').trim();
    row[ledgerIndex.NguonDuLieu] = centerWithdrawal ? 'GIA_DINH_RUT_CHU' : 'GIA_DINH_CU';
    row[ledgerIndex.MaThamChieu] = centerWithdrawal ? ('OWNER_DRAW|' + centerWithdrawal.maGiaoDich) : ('LEGACY_GD|' + legacyId);
    row[ledgerIndex.PhamVi] = 'GIA_DINH';
    row[ledgerIndex.MaDanhMucGiaDinh] = centerWithdrawal ? 'GD_THU_CHU_SO_HUU' : String(familyValues[i][familyIndex.MaDanhMucGiaDinh] || '').trim();
    row[ledgerIndex.LoaiGiaDinh] = familyType;
    row[ledgerIndex.MaGiaoDichLienKet] = centerWithdrawal ? centerWithdrawal.maGiaoDich : legacyId;
    row[ledgerIndex.TrangThai] = String(familyValues[i][familyIndex.TrangThai] || 'ACTIVE').trim().toUpperCase() === 'ACTIVE' ? 'HOAT_DONG' : 'DA_HUY';
    row[ledgerIndex.NguoiTao] = 'Migration';
    row[ledgerIndex.CreatedAt] = familyValues[i][familyIndex.CreatedAt] || now;
    row[ledgerIndex.UpdatedAt] = now;
    newRows.push(row);
    familyValues[i][familyIndex.MaKyHoc] = maKyHoc;
    familyValues[i][familyIndex.MaGiaoDichSoThuChi] = id;
    familyValues[i][familyIndex.UpdatedAt] = now;
    familyChanged = true;
  }
  if (newRows.length) {
    ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, newRows.length, ledgerHeaders.length).setValues(newRows);
  }
  if (familyChanged) {
    familySheet.getRange(2, 1, familyValues.length - 1, familyHeaders.length).setValues(familyValues.slice(1));
  }
  return newRows.length;
}

function reconcileMigratedOwnerWithdrawalsNoLock_(maKyHoc, ledgerSheet) {
  if (!maKyHoc) return 0;
  const values = ledgerSheet.getDataRange().getValues();
  if (values.length <= 1) return 0;
  const headers = values[0].map(header => String(header || '').trim());
  const index = buildHeaderIndex_(headers);
  let changed = 0;
  const now = new Date();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[index.MaKyHoc] || '').trim() !== String(maKyHoc || '').trim()) continue;
    if (String(row[index.NguonDuLieu] || '').trim().toUpperCase() !== 'GIA_DINH_CU') continue;
    if (normalizeFinanceScope_(row[index.PhamVi]) !== 'GIA_DINH') continue;
    const legacy = {
      ngayGiaoDich: formatDateForInput_(toDateOnly_(row[index.NgayGiaoDich])),
      loai: String(row[index.LoaiGiaoDich] || '').trim().toUpperCase(),
      noiDung: String(row[index.NoiDung] || '').trim(),
      soTien: number_(row[index.SoTien])
    };
    const centerWithdrawal = findMatchingCenterOwnerWithdrawal_(values, index, maKyHoc, legacy, i);
    if (!centerWithdrawal) continue;
    const alreadyLinked = values.some((candidate, candidateIndex) => candidateIndex !== i &&
      String(candidate[index.MaKyHoc] || '').trim() === String(maKyHoc || '').trim() &&
      String(candidate[index.NguonDuLieu] || '').trim().toUpperCase() === 'GIA_DINH_RUT_CHU' &&
      String(candidate[index.MaGiaoDichLienKet] || '').trim() === centerWithdrawal.maGiaoDich &&
      String(candidate[index.TrangThai] || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG'
    );
    if (alreadyLinked) {
      row[index.NguonDuLieu] = 'GIA_DINH_CU_DOI_SOAT';
      row[index.TrangThai] = 'DA_HUY';
      row[index.GhiChu] = [String(row[index.GhiChu] || '').trim(), 'Đã loại khỏi giao dịch gia đình do trùng đợt rút tiền trung tâm.'].filter(Boolean).join(' | ');
    } else {
      row[index.LoaiGiaoDich] = 'THU';
      row[index.MaDanhMuc] = 'GD_THU_CHU_SO_HUU';
      row[index.TenDanhMuc] = 'Nhận tiền từ trung tâm';
      row[index.NoiDung] = 'Nhận tiền từ trung tâm - ' + legacy.noiDung;
      row[index.SoPhieu] = generateNextSoPhieuFromRows_('THU_GIA_DINH', legacy.ngayGiaoDich, values.slice(1), index);
      row[index.SoChungTu] = centerWithdrawal.soPhieu;
      row[index.NguoiNopNhan] = 'Trung tâm';
      row[index.NguonDuLieu] = 'GIA_DINH_RUT_CHU';
      row[index.MaThamChieu] = 'OWNER_DRAW|' + centerWithdrawal.maGiaoDich;
      row[index.MaDanhMucGiaDinh] = 'GD_THU_CHU_SO_HUU';
      row[index.LoaiGiaDinh] = 'THU';
      row[index.MaGiaoDichLienKet] = centerWithdrawal.maGiaoDich;
      row[index.GhiChu] = [String(row[index.GhiChu] || '').trim(), 'Đã đối soát từ phiếu rút tiền trung tâm cũ.'].filter(Boolean).join(' | ');
    }
    row[index.UpdatedAt] = now;
    changed++;
  }
  if (changed) ledgerSheet.getRange(2, 1, values.length - 1, headers.length).setValues(values.slice(1));
  return changed;
}


function getDefaultDanhMucThuChi_() {
  return [
    ['THU_HOC_PHI', 'THU', 'Thu học phí', 1, ''],
    ['THU_GHI_DANH', 'THU', 'Phí ghi danh', 2, ''],
    ['THU_BAN_HANG', 'THU', 'Bán hàng, giáo trình', 3, ''],
    ['THU_TAI_TRO', 'THU', 'Tài trợ', 4, ''],
    ['THU_HOAN_PHI', 'THU', 'Hoàn phí', 5, ''],
    ['THU_KHAC', 'THU', 'Thu khác', 99, ''],
    ['CHI_LUONG', 'CHI', 'Lương giáo viên, nhân viên', 1, 'LUONG'],
    ['CHI_BAN_TRU', 'CHI', 'Cơm bán trú, người nấu ăn', 2, 'VAN_HANH'],
    ['CHI_THUE', 'CHI', 'Thuê mặt bằng', 3, 'VAN_HANH'],
    ['CHI_DIEN_NUOC', 'CHI', 'Điện, nước, Internet', 4, 'VAN_HANH'],
    ['CHI_HOC_CU', 'CHI', 'Học cụ, tài liệu', 5, 'VAN_HANH'],
    ['CHI_THUE_PHI', 'CHI', 'Thuế, phí và nghĩa vụ tài chính', 6, 'VAN_HANH'],
    ['CHI_GIA_DINH', 'CHI', 'Chi gia đình / chủ sở hữu', 7, 'LOI_NHUAN'],
    ['CHI_MARKETING', 'CHI', 'Marketing, quảng cáo', 8, 'DAU_TU'],
    ['CHI_SUA_CHUA', 'CHI', 'Sửa chữa, bảo trì', 9, 'VAN_HANH'],
    ['CHI_KHAC', 'CHI', 'Chi khác', 99, 'VAN_HANH']
  ];
}

function normalizeMaHuTaiChinh_(value) {
  const code = String(value || '').trim().toUpperCase();
  // Các mã đã tạo ở phiên bản trước có thể chứa dấu gạch ngang từ UUID.
  // Giữ tương thích để các hũ đã lưu không bị lọc khỏi giao diện.
  return /^[A-Z0-9_-]{2,40}$/.test(code) ? code : '';
}

function inferHuTaiChinhForCategory_(maDanhMuc, tenDanhMuc) {
  const code = String(maDanhMuc || '').trim().toUpperCase();
  const seeded = getDefaultDanhMucThuChi_().find(item => item[0] === code);
  if (seeded && seeded[4]) return seeded[4];
  const text = normalizeText_(String(maDanhMuc || '') + ' ' + String(tenDanhMuc || ''));
  if (text.indexOf('luong') !== -1) return 'LUONG';
  if (text.indexOf('gia dinh') !== -1 || text.indexOf('chu so huu') !== -1 || text.indexOf('loi nhuan') !== -1) return 'LOI_NHUAN';
  if (text.indexOf('dau tu') !== -1 || text.indexOf('phat trien') !== -1 || text.indexOf('marketing') !== -1) return 'DAU_TU';
  if (text.indexOf('du phong') !== -1) return 'DU_PHONG';
  if (text.indexOf('phuc loi') !== -1 || text.indexOf('van hoa') !== -1) return 'PHUC_LOI';
  return 'VAN_HANH';
}

function getDefaultDanhMucGiaDinh_() {
  return [
    ['GD_THU_CHU_SO_HUU', 'Nhận tiền từ trung tâm', 'THU', 'THU_NHAP', 0, 'NONE', 0],
    ['GD_AN_UONG', 'Ăn uống và sinh hoạt', 'CHI', 'THIET_YEU', 40, 'MAX', 1],
    ['GD_HOC_TAP', 'Học tập của các con', 'CHI', 'GIAO_DUC', 15, 'MAX', 2],
    ['GD_Y_TE', 'Y tế và bảo hiểm', 'CHI', 'Y_TE', 5, 'MAX', 3],
    ['GD_GIAI_TRI', 'Ăn ngoài, giải trí, du lịch', 'CHI', 'GIAI_TRI', 10, 'MAX', 4],
    ['GD_QUY_KHAN_CAP', 'Quỹ khẩn cấp', 'TIET_KIEM', 'TIET_KIEM', 15, 'MIN', 5],
    ['GD_TIET_KIEM', 'Tiết kiệm và đầu tư dài hạn', 'TIET_KIEM', 'TIET_KIEM', 10, 'MIN', 6],
    ['GD_KHAC', 'Hiếu hỉ và khoản khác', 'CHI', 'KHAC', 5, 'MAX', 7],
    ['GD_THU_LUONG_NGOAI', 'Lương và thu nhập cá nhân bên ngoài', 'THU', 'THU_NHAP', 0, 'NONE', 90],
    ['GD_THU_LAI_TIET_KIEM', 'Tiền lãi tiết kiệm và đầu tư', 'THU', 'THU_NHAP', 0, 'NONE', 91],
    ['GD_THU_KHAC', 'Thu nhập khác', 'THU', 'THU_NHAP', 0, 'NONE', 99]
  ];
}

function ensureQuanLyTaiChinhSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const schemaKey = 'QUANLY_TAICHINH_SCHEMA_VERSION';
  const schemaVersion = '6';
  const requiredSheets = [
    SHEET_NHANSU_TAICHINH, SHEET_KHOANCHI_DINHKY, SHEET_DANHMUC_HU_TAICHINH,
    SHEET_HU_TAICHINH_THANG, SHEET_CHOT_PHANBO_HU, SHEET_DANHMUC_GIADINH,
    SHEET_GIAODICH_GIADINH, SHEET_CAUHINH_GIADINH_THANG
  ];
  if (props.getProperty(schemaKey) === schemaVersion && requiredSheets.every(name => !!ss.getSheetByName(name))) return;
  ensureSheet_(ss, SHEET_NHANSU_TAICHINH, getNhanSuTaiChinhHeaders_());
  ensureSheet_(ss, SHEET_KHOANCHI_DINHKY, getKhoanChiDinhKyHeaders_());
  const jarCategorySheet = ensureSheet_(ss, SHEET_DANHMUC_HU_TAICHINH, getDanhMucHuTaiChinhHeaders_());
  ensureSheet_(ss, SHEET_HU_TAICHINH_THANG, getHuTaiChinhThangHeaders_());
  const allocationLockSheet = ensureSheet_(ss, SHEET_CHOT_PHANBO_HU, getChotPhanBoHuHeaders_());
  const familyCategorySheet = ensureSheet_(ss, SHEET_DANHMUC_GIADINH, getDanhMucGiaDinhHeaders_());
  ensureSheet_(ss, SHEET_GIAODICH_GIADINH, getGiaoDichGiaDinhHeaders_());
  ensureSheet_(ss, SHEET_CAUHINH_GIADINH_THANG, getCauHinhGiaDinhThangHeaders_());
  const rows = readObjectsNoCache_(SHEET_DANHMUC_GIADINH);
  const existing = rows.reduce((map, row) => {
    const id = String(row.MaDanhMucGiaDinh || '').trim();
    if (id) map[id] = true;
    return map;
  }, {});
  const now = new Date();
  const jarCategoryRows = readObjectsNoCache_(SHEET_DANHMUC_HU_TAICHINH);
  const existingJarCodes = jarCategoryRows.reduce((map, row) => {
    const code = String(row.MaHu || '').trim().toUpperCase();
    if (code) map[code] = true;
    return map;
  }, {});
  const missingJars = getDefaultHuTaiChinh_().filter(item => !existingJarCodes[item.code]).map(item => ({
    MaHu: item.code, TenHu: item.name, TyLeMacDinh: item.ratio, ThuTu: item.order,
    TrangThai: 'ACTIVE', GhiChu: item.note, CreatedAt: now, UpdatedAt: now
  }));
  if (missingJars.length) appendObjectsToSheet_(jarCategorySheet, missingJars, getDanhMucHuTaiChinhHeaders_());
  const ownerJarMigrationKey = 'OWNER_JAR_ROLE_MIGRATION_V1';
  if (props.getProperty(ownerJarMigrationKey) !== '1') {
    const refreshedJarRows = readObjectsNoCache_(SHEET_DANHMUC_HU_TAICHINH);
    const hasOwnerCompensationJar = refreshedJarRows.some(row =>
      String(row.TrangThai || 'ACTIVE').trim().toUpperCase() === 'ACTIVE' &&
      String(row.VaiTroHeThong || '').trim().toUpperCase() === 'OWNER_COMPENSATION'
    );
    if (!hasOwnerCompensationJar) {
      const inferredOwnerJar = refreshedJarRows.find(row => {
        if (String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'ACTIVE') return false;
        const name = normalizeText_(row.TenHu);
        return name.indexOf('luong chu trung tam') !== -1 || name.indexOf('thu lao chu') !== -1;
      });
      if (inferredOwnerJar) {
        upsertFinanceObject_(SHEET_DANHMUC_HU_TAICHINH, getDanhMucHuTaiChinhHeaders_(), 'MaHu', String(inferredOwnerJar.MaHu || '').trim(), {
          VaiTroHeThong: 'OWNER_COMPENSATION'
        });
      }
    }
    props.setProperty(ownerJarMigrationKey, '1');
  }
  const missing = getDefaultDanhMucGiaDinh_().filter(item => !existing[item[0]]).map(item => ({
    MaDanhMucGiaDinh: item[0], TenDanhMuc: item[1], Loai: item[2], Nhom: item[3],
    TyLeMacDinh: item[4], KieuGioiHan: item[5], ThuTu: item[6], TrangThai: 'ACTIVE',
    GhiChu: 'Danh mục gợi ý, có thể chỉnh sửa', CreatedAt: now, UpdatedAt: now
  }));
  if (missing.length) appendObjectsToSheet_(familyCategorySheet, missing, getDanhMucGiaDinhHeaders_());
  const allocationRows = allocationLockSheet.getDataRange().getValues();
  if (allocationRows.length > 1) {
    const allocationIndex = buildHeaderIndex_(allocationRows[0].map(item => String(item || '').trim()));
    let repaired = false;
    for (let i = 1; i < allocationRows.length; i++) {
      const currentMonth = String(allocationRows[i][allocationIndex.Thang] || '').trim();
      const lockId = String(allocationRows[i][allocationIndex.MaChot] || '').trim();
      if (currentMonth || !lockId) continue;
      const separator = lockId.lastIndexOf('|');
      const inferredMonth = separator >= 0 ? lockId.slice(separator + 1) : '';
      if (/^\d{4}-(0[1-9]|1[0-2])$/.test(inferredMonth)) {
        allocationRows[i][allocationIndex.Thang] = inferredMonth;
        repaired = true;
      }
    }
    if (repaired) allocationLockSheet.getRange(1, 1, allocationRows.length, allocationRows[0].length).setValues(allocationRows);
  }
  props.setProperty(schemaKey, schemaVersion);
}

function ensureThuChiSheets_(maKyHoc) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const schemaVersion = '14';
  const schemaKey = 'THUCHI_SCHEMA_VERSION';
  const sourceKey = maKyHoc
    ? 'THUCHI_NGUON_INIT_' + String(maKyHoc || '').trim()
    : '';

  const schemaReady =
    props.getProperty(schemaKey) === schemaVersion &&
    ss.getSheetByName(SHEET_DANHMUC_THUCHI) &&
    ss.getSheetByName(SHEET_SOTHUCHI) &&
    ss.getSheetByName(SHEET_NGUONTIEN) &&
    ss.getSheetByName(SHEET_DOITUONG_THUCHI) &&
    ss.getSheetByName(SHEET_KEHOACH_CHI_THANG) &&
    ss.getSheetByName(SHEET_CAUHINH_TAICHINH_THANG) &&
    ss.getSheetByName(SHEET_DANHMUC_HU_TAICHINH) &&
    ss.getSheetByName(SHEET_CHOT_PHANBO_HU) &&
    props.getProperty('THUCHI_JAR_MIGRATION_V1') === '1';

  const sourceReady = !maKyHoc || props.getProperty(sourceKey) === '1';

  if (schemaReady && sourceReady) return;

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang khởi tạo module thu chi. Vui lòng thao tác lại.');
  }

  try {
    const categorySheet = ensureSheet_(
      ss,
      SHEET_DANHMUC_THUCHI,
      getDanhMucThuChiHeaders_()
    );

    const ledgerSheet = ensureSheet_(
      ss,
      SHEET_SOTHUCHI,
      getSoThuChiHeaders_()
    );

    const sourceSheet = ensureSheet_(
      ss,
      SHEET_NGUONTIEN,
      getNguonTienHeaders_()
    );

    ensureSheet_(ss, SHEET_DOITUONG_THUCHI, getDoiTuongThuChiHeaders_());
    ensureSheet_(ss, SHEET_KEHOACH_CHI_THANG, getKeHoachChiThangHeaders_());
    ensureSheet_(ss, SHEET_CAUHINH_TAICHINH_THANG, getCauHinhTaiChinhThangHeaders_());
    ensureQuanLyTaiChinhSheets_();

    const categoryRows = readObjectsNoCache_(SHEET_DANHMUC_THUCHI);
    const existingCategoryCodes = categoryRows.reduce((map, row) => {
      const code = String(row.MaDanhMuc || '').trim();
      if (code) map[code] = true;
      return map;
    }, {});

    const now = new Date();
    const missingCategories = getDefaultDanhMucThuChi_()
      .filter(item => !existingCategoryCodes[item[0]])
      .map(item => ({
        MaDanhMuc: item[0],
        Loai: item[1],
        TenDanhMuc: item[2],
        MaHuMacDinh: item[4],
        ThuTu: item[3],
        TrangThai: 'ACTIVE',
        CreatedAt: now,
        UpdatedAt: now
      }));

    if (missingCategories.length) {
      appendObjectsToSheet_(
        categorySheet,
        missingCategories,
        getDanhMucThuChiHeaders_()
      );
    }

    if (maKyHoc) {
      const sourceRows = readObjectsNoCache_(SHEET_NGUONTIEN);
      const currentKyHoc = String(maKyHoc || '').trim();
      const existingSourceCodes = sourceRows.reduce((map, row) => {
        if (String(row.MaKyHoc || '').trim() !== currentKyHoc) return map;
        const code = String(row.MaNguonTien || '').trim().toUpperCase();
        if (code) map[code] = true;
        return map;
      }, {});

      const missingSources = getDefaultNguonTien_()
        .filter(item => !existingSourceCodes[item[0]])
        .map(item => ({
          MaKyHoc: currentKyHoc,
          MaNguonTien: item[0],
          TenNguonTien: item[1],
          PhamVi: item[3],
          ThuTu: item[2],
          SoDuBanDau: 0,
          TrangThai: 'ACTIVE',
          CreatedAt: now,
          UpdatedAt: now
        }));

      if (missingSources.length) {
        appendObjectsToSheet_(
          sourceSheet,
          missingSources,
          getNguonTienHeaders_()
        );
      }

      const sourceValues = sourceSheet.getDataRange().getValues();
      if (sourceValues.length > 1) {
        const sourceHeaders = sourceValues[0].map(header => String(header || '').trim());
        const sourceIndex = buildHeaderIndex_(sourceHeaders);
        let sourceScopeChanged = false;
        for (let i = 1; i < sourceValues.length; i++) {
          if (String(sourceValues[i][sourceIndex.MaKyHoc] || '').trim() !== currentKyHoc) continue;
          if (String(sourceValues[i][sourceIndex.PhamVi] || '').trim()) continue;
          const definition = getNguonTienDefinition_(sourceValues[i][sourceIndex.MaNguonTien]);
          sourceValues[i][sourceIndex.PhamVi] = definition ? definition.phamVi : 'TRUNG_TAM';
          sourceValues[i][sourceIndex.UpdatedAt] = now;
          sourceScopeChanged = true;
        }
        if (sourceScopeChanged) sourceSheet.getRange(2, 1, sourceValues.length - 1, sourceHeaders.length).setValues(sourceValues.slice(1));
      }

      props.setProperty(sourceKey, '1');
    }

    const migrationKey = 'THUCHI_SOURCE_MIGRATION_V1';

    if (props.getProperty(migrationKey) !== '1') {
      const values = ledgerSheet.getDataRange().getValues();

      if (values && values.length > 1) {
        const headers = values[0].map(header => String(header || '').trim());
        const index = buildHeaderIndex_(headers);
        let changed = false;

        for (let i = 1; i < values.length; i++) {
          const currentCode = String(
            values[i][index.MaNguonTien] || ''
          ).trim().toUpperCase();

          if (currentCode) continue;

          const inferred = inferNguonTienFromLegacy_(
            values[i][index.HinhThuc],
            values[i][index.TenNguonTien]
          );

          if (!inferred) continue;

          values[i][index.MaNguonTien] = inferred;
          values[i][index.TenNguonTien] = getNguonTienName_(inferred);
          changed = true;
        }

        if (changed) {
          ledgerSheet
            .getRange(2, 1, values.length - 1, headers.length)
            .setValues(values.slice(1));
        }
      }

      props.setProperty(migrationKey, '1');
    }

    // Bổ sung số phiếu cho dữ liệu cũ. Chỉ thực hiện một lần và ghi theo lô.
    const receiptMigrationKey = 'THUCHI_RECEIPT_MIGRATION_V1';

    if (props.getProperty(receiptMigrationKey) !== '1') {
      const values = ledgerSheet.getDataRange().getValues();

      if (values && values.length > 1) {
        const headers = values[0].map(header => String(header || '').trim());
        const index = buildHeaderIndex_(headers);
        const maxByPrefix = {};
        let changed = false;

        // Ghi nhận số thứ tự lớn nhất đang có của từng nhóm phiếu.
        for (let i = 1; i < values.length; i++) {
          const soPhieu = String(values[i][index.SoPhieu] || '').trim();
          const match = soPhieu.match(/^(PT-HP-|PT-|PC-)(\d{6})-(\d+)$/i);
          if (!match) continue;

          const prefix = match[1].toUpperCase() + match[2] + '-';
          maxByPrefix[prefix] = Math.max(
            Number(maxByPrefix[prefix] || 0),
            Number(match[3]) || 0
          );
        }

        for (let i = 1; i < values.length; i++) {
          const row = values[i];
          const maGiaoDich = String(row[index.MaGiaoDich] || '').trim();
          const currentReceipt = String(row[index.SoPhieu] || '').trim();
          const source = String(row[index.NguonDuLieu] || '').trim().toUpperCase();

          if (!maGiaoDich || currentReceipt || source === 'CHUYEN_NOI_BO') continue;

          const receiptType = source === 'HOC_PHI'
            ? 'HOC_PHI'
            : String(row[index.LoaiGiaoDich] || '').trim().toUpperCase();

          if (receiptType !== 'HOC_PHI' && receiptType !== 'THU' && receiptType !== 'CHI') {
            continue;
          }

          const receiptDate = row[index.NgayGiaoDich] || row[index.CreatedAt] || new Date();
          const prefix = buildSoPhieuPrefix_(receiptType, receiptDate);
          const nextNumber = Number(maxByPrefix[prefix] || 0) + 1;
          maxByPrefix[prefix] = nextNumber;
          row[index.SoPhieu] = prefix + String(nextNumber).padStart(4, '0');
          changed = true;
        }

        if (changed) {
          ledgerSheet
            .getRange(2, 1, values.length - 1, headers.length)
            .setValues(values.slice(1));
        }
      }

      props.setProperty(receiptMigrationKey, '1');
    }

    const jarMigrationKey = 'THUCHI_JAR_MIGRATION_V1';
    if (props.getProperty(jarMigrationKey) !== '1') {
      const categoryValues = categorySheet.getDataRange().getValues();
      const categoryHeaders = categoryValues[0].map(header => String(header || '').trim());
      const categoryIndex = buildHeaderIndex_(categoryHeaders);
      const categoryJarMap = {};
      let categoryChanged = false;

      for (let i = 1; i < categoryValues.length; i++) {
        const type = String(categoryValues[i][categoryIndex.Loai] || '').trim().toUpperCase();
        const categoryCode = String(categoryValues[i][categoryIndex.MaDanhMuc] || '').trim();
        let jarCode = normalizeMaHuTaiChinh_(categoryValues[i][categoryIndex.MaHuMacDinh]);
        if (type === 'CHI' && !jarCode) {
          jarCode = inferHuTaiChinhForCategory_(categoryCode, categoryValues[i][categoryIndex.TenDanhMuc]);
          categoryValues[i][categoryIndex.MaHuMacDinh] = jarCode;
          categoryValues[i][categoryIndex.UpdatedAt] = now;
          categoryChanged = true;
        }
        if (categoryCode) categoryJarMap[categoryCode] = jarCode;
      }

      if (categoryChanged && categoryValues.length > 1) {
        categorySheet
          .getRange(2, 1, categoryValues.length - 1, categoryHeaders.length)
          .setValues(categoryValues.slice(1));
      }

      const ledgerValues = ledgerSheet.getDataRange().getValues();
      if (ledgerValues.length > 1) {
        const ledgerHeaders = ledgerValues[0].map(header => String(header || '').trim());
        const ledgerIndex = buildHeaderIndex_(ledgerHeaders);
        let ledgerChanged = false;

        for (let i = 1; i < ledgerValues.length; i++) {
          if (String(ledgerValues[i][ledgerIndex.LoaiGiaoDich] || '').trim().toUpperCase() !== 'CHI') continue;
          if (normalizeMaHuTaiChinh_(ledgerValues[i][ledgerIndex.MaHuTaiChinh])) continue;
          const categoryCode = String(ledgerValues[i][ledgerIndex.MaDanhMuc] || '').trim();
          ledgerValues[i][ledgerIndex.MaHuTaiChinh] = categoryJarMap[categoryCode] || inferHuTaiChinhForCategory_(categoryCode, ledgerValues[i][ledgerIndex.TenDanhMuc]);
          ledgerValues[i][ledgerIndex.UpdatedAt] = now;
          ledgerChanged = true;
        }

        if (ledgerChanged) {
          ledgerSheet
            .getRange(2, 1, ledgerValues.length - 1, ledgerHeaders.length)
            .setValues(ledgerValues.slice(1));
        }
      }

      props.setProperty(jarMigrationKey, '1');
    }

    const scopeValues = ledgerSheet.getDataRange().getValues();
    if (scopeValues.length > 1) {
      const scopeHeaders = scopeValues[0].map(header => String(header || '').trim());
      const scopeIndex = buildHeaderIndex_(scopeHeaders);
      let ledgerScopeChanged = false;
      for (let i = 1; i < scopeValues.length; i++) {
        if (String(scopeValues[i][scopeIndex.PhamVi] || '').trim()) continue;
        const sourceType = String(scopeValues[i][scopeIndex.NguonDuLieu] || '').trim().toUpperCase();
        scopeValues[i][scopeIndex.PhamVi] = sourceType.indexOf('GIA_DINH') === 0 ? 'GIA_DINH' : 'TRUNG_TAM';
        ledgerScopeChanged = true;
      }
      if (ledgerScopeChanged) ledgerSheet.getRange(2, 1, scopeValues.length - 1, scopeHeaders.length).setValues(scopeValues.slice(1));
    }
    if (maKyHoc) {
      migrateLegacyFamilyTransactionsNoLock_(String(maKyHoc || '').trim(), ledgerSheet);
      reconcileMigratedOwnerWithdrawalsNoLock_(String(maKyHoc || '').trim(), ledgerSheet);
    }
    props.setProperty(schemaKey, schemaVersion);
  } finally {
    lock.releaseLock();
  }
}

function normalizeNhomChiTaiChinh_(value) {
  const group = String(value || '').trim().toUpperCase();
  return ['LUONG', 'BAN_TRU', 'VAN_HANH', 'HOC_CU', 'THUE', 'GIA_DINH', 'KHAC'].indexOf(group) !== -1
    ? group
    : 'KHAC';
}

function getCauHinhTaiChinhThang_(maKyHoc, yearMonth) {
  const rows = readObjectsNoCache_(SHEET_CAUHINH_TAICHINH_THANG);
  const row = rows.find(item => {
    return String(item.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
      String(item.Thang || '').trim() === String(yearMonth || '').trim();
  }) || {};

  return {
    thang: yearMonth,
    soHocSinh: number_(row.SoHocSinh),
    doanhThuDuKien: number_(row.DoanhThuDuKien),
    tyLeThue: number_(row.TyLeThue) || 2,
    nguongLuong: number_(row.NguongLuong) || 35,
    nguongBanTru: number_(row.NguongBanTru) || 25,
    nguongTongChi: number_(row.NguongTongChi) || 65,
    tyLeDuPhong: number_(row.TyLeDuPhong) || 10,
    tyLeLoiNhuanMucTieu: number_(row.TyLeLoiNhuanMucTieu) || 20,
    ghiChu: String(row.GhiChu || '').trim()
  };
}

function buildKeHoachChiThangData_(maKyHoc, yearMonth, activeTransactions, closingBalance, actualRevenue) {
  const config = getCauHinhTaiChinhThang_(maKyHoc, yearMonth);
  const rows = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG);
  const staffMap = getNhanSuTaiChinhList_(maKyHoc).reduce((map, item) => { map[item.maNhanSu] = item; return map; }, {});
  const recurringStaffMap = readObjectsNoCache_(SHEET_KHOANCHI_DINHKY).filter(row => String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim()).reduce((map, row) => {
    map[String(row.MaKhoanDinhKy || '').trim()] = String(row.MaNhanSu || '').trim();
    return map;
  }, {});
  const paidMap = (activeTransactions || []).reduce((map, item) => {
    if (item.loai !== 'CHI' || !item.maKeHoachChi) return map;
    map[item.maKeHoachChi] = (map[item.maKeHoachChi] || 0) + number_(item.soTien);
    return map;
  }, {});
  const plans = rows
    .filter(row => {
      return String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
        String(row.Thang || '').trim() === yearMonth &&
        String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED';
    })
    .map(row => {
      const id = String(row.MaKeHoachChi || '').trim();
      const planned = number_(row.SoTienPhaiChi);
      const paid = number_(paidMap[id]);
      const reference = String(row.MaThamChieu || '').trim();
      const staffId = String(row.MaNhanSu || '').trim() || (reference.indexOf('NHANSU|') === 0 ? reference.slice(8) : (reference.indexOf('DINHKY|') === 0 ? String(recurringStaffMap[reference.slice(7)] || '') : ''));
      return {
        maKeHoachChi: id,
        thang: yearMonth,
        maDanhMuc: String(row.MaDanhMuc || '').trim(),
        tenKhoanChi: String(row.TenKhoanChi || '').trim(),
        nhomChi: normalizeNhomChiTaiChinh_(row.NhomChi),
        maNhanSu: staffId,
        nguoiNhan: staffId && staffMap[staffId] ? staffMap[staffId].hoTen : String(row.NguoiNhan || '').trim(),
        soTienPhaiChi: planned,
        soTienDaChi: paid,
        conPhaiChi: Math.max(planned - paid, 0),
        vuotKeHoach: Math.max(paid - planned, 0),
        tienDo: planned > 0 ? paid * 100 / planned : (paid > 0 ? 100 : 0),
        trangThaiThanhToan: paid > planned && planned >= 0 ? 'VUOT' : (planned > 0 && paid >= planned ? 'DA_CHI' : (paid > 0 ? 'MOT_PHAN' : 'CHUA_CHI')),
        hanThanhToan: formatDateForInput_(toDateOnly_(row.HanThanhToan)),
        batBuoc: toBoolean_(row.BatBuoc),
        nguonKeHoach: String(row.NguonKeHoach || '').trim().toUpperCase(),
        maThamChieu: reference,
        ghiChu: String(row.GhiChu || '').trim()
      };
    })
    .sort((a, b) => {
      const groupOrder = { LUONG: 1, BAN_TRU: 2, VAN_HANH: 3, HOC_CU: 4, THUE: 5, GIA_DINH: 6, KHAC: 7 };
      const groupCompare = number_(groupOrder[a.nhomChi] || 99) - number_(groupOrder[b.nhomChi] || 99);
      if (groupCompare !== 0) return groupCompare;
      const dueCompare = String(a.hanThanhToan || '').localeCompare(String(b.hanThanhToan || ''));
      if (dueCompare !== 0) return dueCompare;
      return a.tenKhoanChi.localeCompare(b.tenKhoanChi, 'vi');
    });

  const totalPlan = plans.reduce((sum, item) => sum + item.soTienPhaiChi, 0);
  const totalPaid = plans.reduce((sum, item) => sum + item.soTienDaChi, 0);
  const totalRemaining = plans.reduce((sum, item) => sum + item.conPhaiChi, 0);
  const revenueBasis = config.doanhThuDuKien > 0 ? config.doanhThuDuKien : number_(actualRevenue);
  const reserveTarget = revenueBasis * Math.max(0, number_(config.tyLeDuPhong)) / 100;
  const availableCash = number_(closingBalance) - totalRemaining;
  const safeToSpendCash = availableCash - reserveTarget;
  const groupTotals = plans.reduce((map, item) => {
    map[item.nhomChi] = (map[item.nhomChi] || 0) + item.soTienPhaiChi;
    return map;
  }, {});
  const percent = amount => revenueBasis > 0 ? amount * 100 / revenueBasis : 0;
  const warnings = [];

  plans.forEach(item => {
    if (item.batBuoc && item.soTienPhaiChi <= 0) {
      warnings.push({ level: 'warning', message: item.tenKhoanChi + ' chưa được xác định số tiền phải chi.' });
    }

    if (item.vuotKeHoach > 0) {
      warnings.push({ level: 'danger', message: item.tenKhoanChi + ' đã vượt kế hoạch ' + formatMoneyText_(item.vuotKeHoach) + '.' });
    }
  });

  if (revenueBasis > 0 && percent(groupTotals.LUONG || 0) > config.nguongLuong) {
    warnings.push({ level: 'danger', message: 'Quỹ lương chiếm ' + percent(groupTotals.LUONG || 0).toFixed(1) + '% doanh thu, vượt ngưỡng ' + config.nguongLuong + '%.' });
  }
  if (revenueBasis > 0 && percent(groupTotals.BAN_TRU || 0) > config.nguongBanTru) {
    warnings.push({ level: 'warning', message: 'Chi phí bán trú chiếm ' + percent(groupTotals.BAN_TRU || 0).toFixed(1) + '% doanh thu, vượt ngưỡng ' + config.nguongBanTru + '%.' });
  }
  if (revenueBasis > 0 && percent(totalPlan) > config.nguongTongChi) {
    warnings.push({ level: 'danger', message: 'Tổng nghĩa vụ chi chiếm ' + percent(totalPlan).toFixed(1) + '% doanh thu, vượt ngưỡng ' + config.nguongTongChi + '%.' });
  }
  const expectedTax = revenueBasis * Math.max(0, number_(config.tyLeThue)) / 100;
  if (expectedTax > 0 && number_(groupTotals.THUE) < expectedTax) {
    warnings.push({ level: 'warning', message: 'Khoản dự phòng thuế đang thấp hơn mức cấu hình ' + config.tyLeThue + '%; còn thiếu ' + formatMoneyText_(expectedTax - number_(groupTotals.THUE)) + '.' });
  }
  if (number_(closingBalance) < totalRemaining) {
    warnings.push({ level: 'danger', message: 'Số dư hiện tại không đủ thanh toán các khoản còn phải chi trong tháng.' });
  } else if (availableCash < reserveTarget) {
    warnings.push({ level: 'warning', message: 'Sau khi giữ tiền trả các khoản còn lại, số dư chưa đạt quỹ dự phòng ' + config.tyLeDuPhong + '% (' + formatMoneyText_(reserveTarget) + ').' });
  }
  if (!plans.length) {
    warnings.push({ level: 'info', message: 'Tháng này chưa có kế hoạch chi. Hãy tạo kế hoạch để xác định tiền có thể sử dụng.' });
  }

  return {
    thang: yearMonth,
    config: config,
    items: plans,
    summary: {
      actualRevenue: number_(actualRevenue),
      revenueBasis: revenueBasis,
      currentCash: number_(closingBalance),
      totalPlan: totalPlan,
      totalPaid: totalPaid,
      totalRemaining: totalRemaining,
      availableCash: availableCash,
      reserveTarget: reserveTarget,
      safeToSpendCash: safeToSpendCash,
      expectedSurplus: revenueBasis - totalPlan,
      costPerStudent: config.soHocSinh > 0 ? totalPlan / config.soHocSinh : 0,
      boardingCostPerStudent: config.soHocSinh > 0 ? number_(groupTotals.BAN_TRU) / config.soHocSinh : 0,
      groupTotals: groupTotals,
      groupRates: {
        salary: percent(groupTotals.LUONG || 0),
        boarding: percent(groupTotals.BAN_TRU || 0),
        totalExpense: percent(totalPlan)
      }
    },
    warnings: warnings
  };
}

function formatMoneyText_(value) {
  return Math.round(number_(value)).toLocaleString('vi-VN') + 'đ';
}

function saveCauHinhTaiChinhThang(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);
  data = data || {};
  const yearMonth = String(data.thang || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(yearMonth)) throw new Error('Tháng tài chính không hợp lệ.');
  assertFinancePeriodOpen_(session, yearMonth);
  const values = {
    SoHocSinh: Math.max(0, number_(data.soHocSinh)),
    DoanhThuDuKien: Math.max(0, moneyNumber_(data.doanhThuDuKien)),
    TyLeThue: Math.max(0, number_(data.tyLeThue) || 2),
    NguongLuong: Math.max(0, number_(data.nguongLuong) || 35),
    NguongBanTru: Math.max(0, number_(data.nguongBanTru) || 25),
    NguongTongChi: Math.max(0, number_(data.nguongTongChi) || 65),
    TyLeDuPhong: Math.max(0, number_(data.tyLeDuPhong) || 10),
    TyLeLoiNhuanMucTieu: Math.max(0, Math.min(80, number_(data.tyLeLoiNhuanMucTieu) || 20)),
    GhiChu: String(data.ghiChu || '').trim()
  };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật cấu hình tài chính.');
  try {
    const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_CAUHINH_TAICHINH_THANG, getCauHinhTaiChinhThangHeaders_());
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(item => String(item || '').trim());
    const index = buildHeaderIndex_(headers);
    let target = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][index.MaKyHoc] || '').trim() === session.maKyHoc && String(rows[i][index.Thang] || '').trim() === yearMonth) { target = i; break; }
    }
    const now = new Date();
    const row = target >= 0 ? rows[target].slice() : new Array(headers.length).fill('');
    row[index.MaKyHoc] = session.maKyHoc;
    row[index.Thang] = yearMonth;
    Object.keys(values).forEach(key => { row[index[key]] = values[key]; });
    row[index.CreatedAt] = target >= 0 ? (row[index.CreatedAt] || now) : now;
    row[index.UpdatedAt] = now;
    if (target >= 0) sheet.getRange(target + 1, 1, 1, headers.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  safeWriteAuditLog_(session, 'UPSERT', 'CAU_HINH_TAI_CHINH', yearMonth, null, values);
  return jsonResponse_({ success: true, message: 'Đã cập nhật cấu hình tài chính tháng.' });
}

function saveKeHoachChiThang(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);
  data = data || {};
  const idInput = String(data.maKeHoachChi || '').trim();
  const yearMonth = String(data.thang || '').trim();
  const name = String(data.tenKhoanChi || '').trim();
  const categoryId = String(data.maDanhMuc || '').trim();
  const amount = Math.max(0, moneyNumber_(data.soTienPhaiChi));
  const dueText = String(data.hanThanhToan || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(yearMonth)) throw new Error('Tháng kế hoạch không hợp lệ.');
  assertFinancePeriodOpen_(session, yearMonth);
  if (!name || !categoryId) throw new Error('Vui lòng nhập tên khoản chi và danh mục.');
  if (dueText && !toDateOnly_(dueText)) throw new Error('Hạn thanh toán không hợp lệ.');
  const category = readObjectsNoCache_(SHEET_DANHMUC_THUCHI).find(row => String(row.MaDanhMuc || '').trim() === categoryId && String(row.Loai || '').trim().toUpperCase() === 'CHI');
  if (!category) throw new Error('Danh mục chi không hợp lệ.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật kế hoạch chi.');
  let savedId = idInput;
  try {
    const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_KEHOACH_CHI_THANG, getKeHoachChiThangHeaders_());
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(item => String(item || '').trim());
    const index = buildHeaderIndex_(headers);
    let target = -1;
    if (savedId) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][index.MaKeHoachChi] || '').trim() === savedId && String(rows[i][index.MaKyHoc] || '').trim() === session.maKyHoc) { target = i; break; }
      }
      if (target < 0) throw new Error('Không tìm thấy khoản kế hoạch cần sửa.');
    } else savedId = 'KHCHI_' + Utilities.getUuid().slice(0, 10).toUpperCase();
    const now = new Date();
    const row = target >= 0 ? rows[target].slice() : new Array(headers.length).fill('');
    row[index.MaKeHoachChi] = savedId;
    row[index.MaKyHoc] = session.maKyHoc;
    row[index.Thang] = yearMonth;
    row[index.MaDanhMuc] = categoryId;
    row[index.TenKhoanChi] = name;
    row[index.NhomChi] = normalizeNhomChiTaiChinh_(data.nhomChi);
    row[index.NguoiNhan] = String(data.nguoiNhan || '').trim();
    row[index.SoTienPhaiChi] = amount;
    row[index.HanThanhToan] = dueText ? parseInputDate_(dueText) : '';
    row[index.BatBuoc] = toBoolean_(data.batBuoc) ? 'Có' : 'Không';
    if (index.NguonKeHoach !== undefined) row[index.NguonKeHoach] = String(data.nguonKeHoach || row[index.NguonKeHoach] || 'NHAP_TAY').trim().toUpperCase();
    if (index.MaThamChieu !== undefined) row[index.MaThamChieu] = String(data.maThamChieu || row[index.MaThamChieu] || '').trim();
    row[index.GhiChu] = String(data.ghiChu || '').trim();
    row[index.TrangThai] = 'ACTIVE';
    row[index.CreatedAt] = target >= 0 ? (row[index.CreatedAt] || now) : now;
    row[index.UpdatedAt] = now;
    if (target >= 0) sheet.getRange(target + 1, 1, 1, headers.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  safeWriteAuditLog_(session, idInput ? 'UPDATE' : 'CREATE', 'KE_HOACH_CHI', savedId, null, { thang: yearMonth, tenKhoanChi: name, soTienPhaiChi: amount });
  return jsonResponse_({ success: true, maKeHoachChi: savedId, message: 'Đã lưu khoản phải chi.' });
}

function deleteKeHoachChiThang(token, id) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);
  const planId = String(id || '').trim();
  const currentPlan = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG).find(row =>
    String(row.MaKeHoachChi || '').trim() === planId && String(row.MaKyHoc || '').trim() === session.maKyHoc
  );
  if (!currentPlan) throw new Error('Không tìm thấy khoản kế hoạch cần xoá.');
  assertFinancePeriodOpen_(session, String(currentPlan.Thang || '').trim());
  const hasLinkedPayment = readObjectsNoCache_(SHEET_SOTHUCHI).some(row => {
    return String(row.MaKyHoc || '').trim() === session.maKyHoc &&
      String(row.MaKeHoachChi || '').trim() === planId &&
      String(row.TrangThai || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG';
  });
  if (hasLinkedPayment) throw new Error('Khoản này đã có phiếu chi. Hãy huỷ phiếu chi liên quan trước khi xoá kế hoạch.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật kế hoạch chi.');
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KEHOACH_CHI_THANG);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(item => String(item || '').trim());
    const index = buildHeaderIndex_(headers);
    let target = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][index.MaKeHoachChi] || '').trim() === planId && String(rows[i][index.MaKyHoc] || '').trim() === session.maKyHoc) { target = i; break; }
    }
    if (target < 0) throw new Error('Không tìm thấy khoản kế hoạch cần xoá.');
    rows[target][index.TrangThai] = 'DELETED';
    rows[target][index.UpdatedAt] = new Date();
    sheet.getRange(target + 1, 1, 1, headers.length).setValues([rows[target]]);
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  safeWriteAuditLog_(session, 'DELETE', 'KE_HOACH_CHI', planId, currentPlan, null);
  return jsonResponse_({ success: true, message: 'Đã xoá khoản phải chi khỏi kế hoạch.' });
}

function khoiTaoKeHoachChiThangMauLegacy_(token, yearMonth) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);
  yearMonth = String(yearMonth || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(yearMonth)) throw new Error('Tháng kế hoạch không hợp lệ.');
  const existing = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG).some(row => String(row.MaKyHoc || '').trim() === session.maKyHoc && String(row.Thang || '').trim() === yearMonth && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED');
  if (existing) throw new Error('Tháng này đã có kế hoạch chi. Vui lòng chỉnh sửa trực tiếp để tránh trùng dữ liệu.');
  const due = yearMonth + '-28';
  const templates = [
    ['CHI_LUONG', 'Lương Thầy Quý', 'LUONG', 'Thầy Quý', 11000000, true, 'Lương giáo viên'],
    ['CHI_LUONG', 'Lương Cô Xuân', 'LUONG', 'Cô Xuân', 8000000, true, 'Lương giáo viên'],
    ['CHI_LUONG', 'Lương Cô Nhi', 'LUONG', 'Cô Nhi', 6000000, true, 'Lương giáo viên'],
    ['CHI_LUONG', 'Lương Cô Trâm', 'LUONG', 'Cô Trâm', 8000000, true, 'Lương giáo viên'],
    ['CHI_LUONG', 'Lương Cô Hiền', 'LUONG', 'Cô Hiền', 5000000, true, 'Lương giáo viên'],
    ['CHI_LUONG', 'Lương Cô Nhiên', 'LUONG', 'Cô Nhiên', 3000000, true, 'Lương giáo viên'],
    ['CHI_LUONG', 'Khoản quản lý Thầy Lãm', 'LUONG', 'Thầy Lãm', 0, true, 'Chủ cơ sở - chưa xác định mức chi'],
    ['CHI_BAN_TRU', 'Thực phẩm cơm bán trú', 'BAN_TRU', '', 36000000, true, '45 triệu bán trú trừ 9 triệu lương người nấu'],
    ['CHI_BAN_TRU', 'Lương người nấu ăn', 'BAN_TRU', 'Người nấu ăn', 9000000, true, 'Đã nằm trong tổng 45 triệu bán trú'],
    ['CHI_DIEN_NUOC', 'Tiền điện', 'VAN_HANH', '', 3000000, true, ''],
    ['CHI_DIEN_NUOC', 'Tiền Internet', 'VAN_HANH', '', 300000, true, ''],
    ['CHI_THUE', 'Thuê mặt bằng', 'VAN_HANH', '', 0, false, 'Cơ sở tại nhà, không phát sinh'],
    ['CHI_HOC_CU', 'Dụng cụ học tập', 'HOC_CU', '', 2000000, true, ''],
    ['CHI_KHAC', 'Chi phí khác', 'KHAC', '', 5000000, true, ''],
    ['CHI_THUE_PHI', 'Dự phòng thuế 2%', 'THUE', '', 3360000, true, 'Tạm tính 2% trên doanh thu dự kiến 168 triệu']
  ];
  const categoryRows = readObjectsNoCache_(SHEET_DANHMUC_THUCHI);
  const categoryMap = categoryRows.reduce((map, row) => { map[String(row.MaDanhMuc || '').trim()] = row; return map; }, {});
  const now = new Date();
  const objects = templates.map(item => ({
    MaKeHoachChi: 'KHCHI_' + Utilities.getUuid().slice(0, 10).toUpperCase(),
    MaKyHoc: session.maKyHoc, Thang: yearMonth, MaDanhMuc: item[0],
    TenKhoanChi: item[1], NhomChi: item[2], NguoiNhan: item[3],
    SoTienPhaiChi: item[4], HanThanhToan: parseInputDate_(due), BatBuoc: item[5] ? 'Có' : 'Không',
    GhiChu: item[6], TrangThai: 'ACTIVE', CreatedAt: now, UpdatedAt: now
  })).filter(item => !!categoryMap[item.MaDanhMuc]);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang khởi tạo kế hoạch chi.');
  try {
    const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_KEHOACH_CHI_THANG, getKeHoachChiThangHeaders_());
    appendObjectsToSheet_(sheet, objects, getKeHoachChiThangHeaders_());
  } finally { lock.releaseLock(); }
  saveCauHinhTaiChinhThang(token, {
    thang: yearMonth, soHocSinh: 86, doanhThuDuKien: 168000000, tyLeThue: 2,
    nguongLuong: 35, nguongBanTru: 25, nguongTongChi: 65, tyLeDuPhong: 10,
    ghiChu: 'Kế hoạch tháng hè; không phát sinh xe đưa rước.'
  });
  bumpDataVersion_();
  return jsonResponse_({ success: true, count: objects.length, message: 'Đã tạo kế hoạch chi mẫu tháng ' + yearMonth.slice(5, 7) + '/' + yearMonth.slice(0, 4) + '.' });
}

function upsertFinanceObject_(sheetName, headers, idHeader, idValue, values) {
  const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), sheetName, headers);
  const rows = sheet.getDataRange().getValues();
  const currentHeaders = rows[0].map(item => String(item || '').trim());
  const index = buildHeaderIndex_(currentHeaders);
  let target = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][index[idHeader]] || '').trim() === idValue) { target = i; break; }
  }
  const now = new Date();
  const row = target >= 0 ? rows[target].slice() : new Array(currentHeaders.length).fill('');
  Object.keys(values).forEach(key => {
    if (index[key] !== undefined) row[index[key]] = values[key];
  });
  row[index[idHeader]] = idValue;
  if (index.CreatedAt !== undefined) row[index.CreatedAt] = target >= 0 ? (row[index.CreatedAt] || now) : now;
  if (index.UpdatedAt !== undefined) row[index.UpdatedAt] = now;
  if (target >= 0) sheet.getRange(target + 1, 1, 1, currentHeaders.length).setValues([row]);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, currentHeaders.length).setValues([row]);
}

function getNhanSuTaiChinhList_(maKyHoc) {
  return readObjects_(SHEET_NHANSU_TAICHINH).filter(row => String(row.MaKyHoc || '').trim() === maKyHoc).map(row => ({
    maNhanSu: String(row.MaNhanSu || '').trim(), hoTen: String(row.HoTen || '').trim(),
    vaiTro: String(row.VaiTro || '').trim(), maDanhMuc: String(row.MaDanhMuc || 'CHI_LUONG').trim(),
    nhomChi: normalizeNhomChiTaiChinh_(row.NhomChi || 'LUONG'), mucChiMacDinh: number_(row.MucChiMacDinh),
    ngayThanhToan: Math.max(1, Math.min(31, number_(row.NgayThanhToan) || 28)),
    lopPhuTrach: String(row.LopPhuTrach || '').trim(), tuNgay: formatDateForInput_(toDateOnly_(row.TuNgay)),
    denNgay: formatDateForInput_(toDateOnly_(row.DenNgay)), trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase(),
    ghiChu: String(row.GhiChu || '').trim()
  })).filter(item => item.maNhanSu && item.hoTen).sort((a, b) => a.hoTen.localeCompare(b.hoTen, 'vi'));
}

function saveNhanSuTaiChinh(token, data) {
  const session = requireSession_(token, 'finance.write'); ensureThuChiSheets_(session.maKyHoc); data = data || {};
  const id = String(data.maNhanSu || '').trim() || ('NS_' + Utilities.getUuid().slice(0, 10).toUpperCase());
  const name = String(data.hoTen || '').trim();
  if (!name) throw new Error('Vui lòng nhập họ tên nhân sự.');
  const categoryId = String(data.maDanhMuc || 'CHI_LUONG').trim();
  const category = readObjectsNoCache_(SHEET_DANHMUC_THUCHI).find(row => String(row.MaDanhMuc || '').trim() === categoryId && String(row.Loai || '').trim().toUpperCase() === 'CHI');
  if (!category) throw new Error('Danh mục chi của nhân sự không hợp lệ.');
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật nhân sự.');
  try {
    upsertFinanceObject_(SHEET_NHANSU_TAICHINH, getNhanSuTaiChinhHeaders_(), 'MaNhanSu', id, {
      MaKyHoc: session.maKyHoc, HoTen: name, VaiTro: String(data.vaiTro || '').trim(), MaDanhMuc: categoryId,
      NhomChi: normalizeNhomChiTaiChinh_(data.nhomChi || 'LUONG'), MucChiMacDinh: Math.max(0, moneyNumber_(data.mucChiMacDinh)),
      NgayThanhToan: Math.max(1, Math.min(31, number_(data.ngayThanhToan) || 28)), LopPhuTrach: String(data.lopPhuTrach || '').trim(),
      TuNgay: data.tuNgay ? parseInputDate_(data.tuNgay) : '', DenNgay: data.denNgay ? parseInputDate_(data.denNgay) : '',
      TrangThai: String(data.trangThai || 'ACTIVE').trim().toUpperCase() === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE', GhiChu: String(data.ghiChu || '').trim()
    });
  } finally { lock.releaseLock(); }
  bumpDataVersion_(); return jsonResponse_({ success: true, maNhanSu: id, message: 'Đã lưu thông tin nhân sự.' });
}

function setTrangThaiNhanSuTaiChinh(token, id, enabled) {
  const session = requireSession_(token, 'finance.write'); ensureThuChiSheets_(session.maKyHoc);
  const rows = readObjectsNoCache_(SHEET_NHANSU_TAICHINH); const current = rows.find(row => String(row.MaNhanSu || '').trim() === String(id || '').trim() && String(row.MaKyHoc || '').trim() === session.maKyHoc);
  if (!current) throw new Error('Không tìm thấy nhân sự.');
  current.TrangThai = toBoolean_(enabled) ? 'ACTIVE' : 'INACTIVE';
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật nhân sự.');
  try { upsertFinanceObject_(SHEET_NHANSU_TAICHINH, getNhanSuTaiChinhHeaders_(), 'MaNhanSu', String(id), current); } finally { lock.releaseLock(); }
  bumpDataVersion_(); return jsonResponse_({ success: true, message: toBoolean_(enabled) ? 'Đã kích hoạt nhân sự.' : 'Đã ngừng nhân sự.' });
}

function financeYearMonthValue_(value) {
  if (!value) return '';
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(0[1-9]|1[0-2])/);
  if (match) return match[1] + '-' + match[2];
  const date = toDateOnly_(value);
  return date ? Utilities.formatDate(date, 'Asia/Ho_Chi_Minh', 'yyyy-MM') : '';
}

function getKhoanChiDinhKyList_(maKyHoc, yearMonth) {
  const staffMap = getNhanSuTaiChinhList_(maKyHoc).reduce((map, item) => { map[item.maNhanSu] = item; return map; }, {});
  const month = String(yearMonth || '').trim();
  let planMap = {};
  let paidMap = {};
  if (/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
    planMap = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG).filter(row =>
      String(row.MaKyHoc || '').trim() === maKyHoc && String(row.Thang || '').trim() === month &&
      String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED' &&
      String(row.MaThamChieu || '').trim().indexOf('DINHKY|') === 0
    ).reduce((map, row) => { map[String(row.MaThamChieu || '').trim().slice(7)] = row; return map; }, {});
    paidMap = readObjectsNoCache_(SHEET_SOTHUCHI).filter(row =>
      String(row.MaKyHoc || '').trim() === maKyHoc && String(row.LoaiGiaoDich || '').trim().toUpperCase() === 'CHI' &&
      String(row.TrangThai || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG'
    ).reduce((map, row) => {
      const id = String(row.MaKeHoachChi || '').trim();
      if (id) map[id] = (map[id] || 0) + number_(row.SoTien);
      return map;
    }, {});
  }
  return readObjects_(SHEET_KHOANCHI_DINHKY).filter(row => String(row.MaKyHoc || '').trim() === maKyHoc).map(row => {
    const id = String(row.MaKhoanDinhKy || '').trim();
    const maNhanSu = String(row.MaNhanSu || '').trim();
    const plan = planMap[id] || null;
    const planned = plan ? number_(plan.SoTienPhaiChi) : 0;
    const paid = plan ? number_(paidMap[String(plan.MaKeHoachChi || '').trim()]) : 0;
    return {
      maKhoanDinhKy: id, tenKhoanChi: String(row.TenKhoanChi || '').trim(),
      maDanhMuc: String(row.MaDanhMuc || '').trim(), nhomChi: normalizeNhomChiTaiChinh_(row.NhomChi),
      maNhanSu: maNhanSu, nguoiNhan: maNhanSu && staffMap[maNhanSu] ? staffMap[maNhanSu].hoTen : '',
      loaiKhoanNhanSu: String(row.LoaiKhoanNhanSu || '').trim().toUpperCase(),
      phuongPhapTinh: String(row.PhuongPhapTinh || 'FIXED').trim().toUpperCase(), dinhMuc: number_(row.DinhMuc),
      ngayThanhToan: Math.max(1, Math.min(31, number_(row.NgayThanhToan) || 28)), batBuoc: toBoolean_(row.BatBuoc),
      tuThang: financeYearMonthValue_(row.TuThang), denThang: financeYearMonthValue_(row.DenThang),
      trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase(), ghiChu: String(row.GhiChu || '').trim(),
      maKeHoachChi: plan ? String(plan.MaKeHoachChi || '').trim() : '', soTienPhaiChi: planned, soTienDaChi: paid,
      conPhaiChi: Math.max(planned - paid, 0), vuotKeHoach: Math.max(paid - planned, 0),
      tienDo: planned > 0 ? paid * 100 / planned : (paid > 0 ? 100 : 0),
      trangThaiThanhToan: !plan ? 'CHUA_TAO' : (paid > planned ? 'VUOT' : (planned > 0 && paid >= planned ? 'DA_CHI' : (paid > 0 ? 'MOT_PHAN' : 'CHUA_CHI')))
    };
  }).filter(item => item.maKhoanDinhKy && item.tenKhoanChi).sort((a, b) => a.tenKhoanChi.localeCompare(b.tenKhoanChi, 'vi'));
}

function saveKhoanChiDinhKy(token, data) {
  const session = requireSession_(token, 'finance.write'); ensureThuChiSheets_(session.maKyHoc); data = data || {};
  const id = String(data.maKhoanDinhKy || '').trim() || ('KCDK_' + Utilities.getUuid().slice(0, 10).toUpperCase());
  const name = String(data.tenKhoanChi || '').trim(); const categoryId = String(data.maDanhMuc || '').trim();
  if (!name || !categoryId) throw new Error('Vui lòng nhập tên khoản chi và danh mục.');
  const methods = ['FIXED', 'PER_STUDENT', 'PERCENT_REVENUE', 'MANUAL'];
  const method = String(data.phuongPhapTinh || 'FIXED').trim().toUpperCase();
  if (methods.indexOf(method) === -1) throw new Error('Phương pháp tính khoản chi không hợp lệ.');
  const category = readObjectsNoCache_(SHEET_DANHMUC_THUCHI).find(row => String(row.MaDanhMuc || '').trim() === categoryId && String(row.Loai || '').trim().toUpperCase() === 'CHI');
  if (!category) throw new Error('Danh mục chi không hợp lệ.');
  let maNhanSu = '';
  let loaiKhoanNhanSu = '';
  if (categoryId === 'CHI_LUONG') {
    maNhanSu = String(data.maNhanSu || '').trim();
    const staff = getNhanSuTaiChinhList_(session.maKyHoc).find(item => item.maNhanSu === maNhanSu && item.trangThai === 'ACTIVE');
    if (!staff) throw new Error('Vui lòng chọn người nhận từ danh sách nhân sự đang làm việc.');
    loaiKhoanNhanSu = String(data.loaiKhoanNhanSu || '').trim().toUpperCase();
    if (['PHU_CAP', 'THUONG', 'KHAC'].indexOf(loaiKhoanNhanSu) === -1) {
      throw new Error('Khoản lương định kỳ chỉ dùng cho phụ cấp, thưởng hoặc khoản khác. Lương chính quản lý tại tab Nhân sự.');
    }
  }
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật khoản chi định kỳ.');
  try {
    upsertFinanceObject_(SHEET_KHOANCHI_DINHKY, getKhoanChiDinhKyHeaders_(), 'MaKhoanDinhKy', id, {
      MaKyHoc: session.maKyHoc, TenKhoanChi: name, MaDanhMuc: categoryId, NhomChi: normalizeNhomChiTaiChinh_(data.nhomChi),
      MaNhanSu: maNhanSu, LoaiKhoanNhanSu: loaiKhoanNhanSu,
      PhuongPhapTinh: method, DinhMuc: Math.max(0, number_(data.dinhMuc)), NgayThanhToan: Math.max(1, Math.min(31, number_(data.ngayThanhToan) || 28)),
      BatBuoc: toBoolean_(data.batBuoc) ? 'Có' : 'Không', TuThang: String(data.tuThang || '').trim(), DenThang: String(data.denThang || '').trim(),
      TrangThai: String(data.trangThai || 'ACTIVE').trim().toUpperCase() === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE', GhiChu: String(data.ghiChu || '').trim()
    });
  } finally { lock.releaseLock(); }
  let syncMessage = '';
  const syncMonth = String(data.thangKeHoach || '').trim();
  if (/^\d{4}-(0[1-9]|1[0-2])$/.test(syncMonth)) {
    try {
      const synced = syncKhoanChiDinhKyPlan_(session, id, syncMonth);
      syncMessage = synced.created ? ' Đã tạo kế hoạch chi tháng tương ứng.' : ' Đã cập nhật kế hoạch chi tháng tương ứng.';
    } catch (error) {
      syncMessage = ' Khoản định kỳ đã lưu nhưng chưa đồng bộ được kế hoạch tháng: ' + String(error && error.message || error);
    }
  }
  bumpDataVersion_(); return jsonResponse_({ success: true, maKhoanDinhKy: id, message: 'Đã lưu khoản chi định kỳ.' + syncMessage });
}

function syncKhoanChiDinhKyPlan_(session, id, yearMonth) {
  yearMonth = String(yearMonth || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(yearMonth)) throw new Error('Tháng kế hoạch không hợp lệ.');
  assertFinancePeriodOpen_(session, yearMonth);
  const row = readObjectsNoCache_(SHEET_KHOANCHI_DINHKY).find(item =>
    String(item.MaKyHoc || '').trim() === session.maKyHoc && String(item.MaKhoanDinhKy || '').trim() === String(id || '').trim()
  );
  if (!row || String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'ACTIVE') throw new Error('Khoản chi định kỳ không còn hoạt động.');
  const fromMonth = financeYearMonthValue_(row.TuThang);
  const toMonth = financeYearMonthValue_(row.DenThang);
  if ((fromMonth && fromMonth > yearMonth) || (toMonth && toMonth < yearMonth)) throw new Error('Khoản chi chưa có hiệu lực trong tháng đã chọn.');
  const config = getCauHinhTaiChinhThang_(session.maKyHoc, yearMonth);
  const method = String(row.PhuongPhapTinh || 'FIXED').trim().toUpperCase();
  let amount = number_(row.DinhMuc);
  if (method === 'PER_STUDENT') amount = number_(row.DinhMuc) * number_(config.soHocSinh);
  if (method === 'PERCENT_REVENUE') amount = number_(row.DinhMuc) * number_(config.doanhThuDuKien) / 100;
  const reference = 'DINHKY|' + String(id || '').trim();
  const existing = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG).find(item =>
    String(item.MaKyHoc || '').trim() === session.maKyHoc && String(item.Thang || '').trim() === yearMonth &&
    String(item.MaThamChieu || '').trim() === reference && String(item.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED'
  );
  if (method === 'MANUAL' && existing) amount = number_(existing.SoTienPhaiChi);
  if (method === 'MANUAL' && !existing) amount = 0;
  const maNhanSu = String(row.MaNhanSu || '').trim();
  const staff = maNhanSu ? getNhanSuTaiChinhList_(session.maKyHoc).find(item => item.maNhanSu === maNhanSu) : null;
  const planId = existing ? String(existing.MaKeHoachChi || '').trim() : ('KHCHI_' + Utilities.getUuid().slice(0, 10).toUpperCase());
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang đồng bộ kế hoạch chi tháng.');
  try {
    upsertFinanceObject_(SHEET_KEHOACH_CHI_THANG, getKeHoachChiThangHeaders_(), 'MaKeHoachChi', planId, {
      MaKyHoc: session.maKyHoc, Thang: yearMonth, MaDanhMuc: String(row.MaDanhMuc || '').trim(),
      TenKhoanChi: String(row.TenKhoanChi || '').trim(), NhomChi: normalizeNhomChiTaiChinh_(row.NhomChi),
      MaNhanSu: maNhanSu, NguoiNhan: staff ? staff.hoTen : '', SoTienPhaiChi: Math.max(0, amount),
      HanThanhToan: parseInputDate_(financeDueDate_(yearMonth, row.NgayThanhToan)), BatBuoc: toBoolean_(row.BatBuoc) ? 'Có' : 'Không',
      NguonKeHoach: 'DINHKY', MaThamChieu: reference, GhiChu: String(row.GhiChu || '').trim(), TrangThai: 'ACTIVE'
    });
  } finally { lock.releaseLock(); }
  return { created: !existing, maKeHoachChi: planId };
}

function syncKhoanChiDinhKyVaoKeHoach(token, id, yearMonth) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);
  const result = syncKhoanChiDinhKyPlan_(session, String(id || '').trim(), String(yearMonth || '').trim());
  bumpDataVersion_();
  safeWriteAuditLog_(session, result.created ? 'CREATE' : 'UPDATE', 'KE_HOACH_CHI_DINH_KY', result.maKeHoachChi, null, { maKhoanDinhKy: id, thang: yearMonth });
  return jsonResponse_({ success: true, maKeHoachChi: result.maKeHoachChi, message: result.created ? 'Đã tạo kế hoạch tháng từ khoản chi định kỳ.' : 'Đã cập nhật kế hoạch tháng từ khoản chi định kỳ.' });
}

function setTrangThaiKhoanChiDinhKy(token, id, enabled) {
  const session = requireSession_(token, 'finance.write'); ensureThuChiSheets_(session.maKyHoc);
  const rows = readObjectsNoCache_(SHEET_KHOANCHI_DINHKY); const current = rows.find(row => String(row.MaKhoanDinhKy || '').trim() === String(id || '').trim() && String(row.MaKyHoc || '').trim() === session.maKyHoc);
  if (!current) throw new Error('Không tìm thấy khoản chi định kỳ.');
  current.TrangThai = toBoolean_(enabled) ? 'ACTIVE' : 'INACTIVE';
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật khoản chi.');
  try { upsertFinanceObject_(SHEET_KHOANCHI_DINHKY, getKhoanChiDinhKyHeaders_(), 'MaKhoanDinhKy', String(id), current); } finally { lock.releaseLock(); }
  bumpDataVersion_(); return jsonResponse_({ success: true, message: toBoolean_(enabled) ? 'Đã kích hoạt khoản chi.' : 'Đã ngừng khoản chi.' });
}

function financeDueDate_(yearMonth, day) {
  const parts = yearMonth.split('-'); const year = Number(parts[0]); const month = Number(parts[1]);
  const lastDay = new Date(year, month, 0).getDate();
  return yearMonth + '-' + String(Math.min(lastDay, Math.max(1, number_(day) || 28))).padStart(2, '0');
}

function taoKeHoachTaiChinhTuDanhMuc(token, yearMonth) {
  const session = requireSession_(token, 'finance.write'); ensureThuChiSheets_(session.maKyHoc); yearMonth = String(yearMonth || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(yearMonth)) throw new Error('Tháng kế hoạch không hợp lệ.');
  assertFinancePeriodOpen_(session, yearMonth);
  const config = getCauHinhTaiChinhThang_(session.maKyHoc, yearMonth);
  const staff = getNhanSuTaiChinhList_(session.maKyHoc).filter(item => item.trangThai === 'ACTIVE' && (!item.tuNgay || item.tuNgay.slice(0, 7) <= yearMonth) && (!item.denNgay || item.denNgay.slice(0, 7) >= yearMonth));
  const recurring = getKhoanChiDinhKyList_(session.maKyHoc, yearMonth).filter(item => item.trangThai === 'ACTIVE' && (!item.tuThang || item.tuThang <= yearMonth) && (!item.denThang || item.denThang >= yearMonth));
  if (!staff.length && !recurring.length) throw new Error('Chưa có nhân sự hoặc khoản chi định kỳ đang hoạt động để tạo kế hoạch.');
  const existingRefs = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG).filter(row => String(row.MaKyHoc || '').trim() === session.maKyHoc && String(row.Thang || '').trim() === yearMonth && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED').reduce((map, row) => { const ref = String(row.MaThamChieu || '').trim(); if (ref) map[ref] = true; return map; }, {});
  const now = new Date(); const rows = [];
  staff.forEach(item => {
    const ref = 'NHANSU|' + item.maNhanSu; if (existingRefs[ref]) return;
    rows.push({ MaKeHoachChi: 'KHCHI_' + Utilities.getUuid().slice(0, 10).toUpperCase(), MaKyHoc: session.maKyHoc, Thang: yearMonth,
      MaDanhMuc: item.maDanhMuc || 'CHI_LUONG', TenKhoanChi: 'Lương ' + item.hoTen, NhomChi: item.nhomChi || 'LUONG', MaNhanSu: item.maNhanSu, NguoiNhan: item.hoTen,
      SoTienPhaiChi: item.mucChiMacDinh, HanThanhToan: parseInputDate_(financeDueDate_(yearMonth, item.ngayThanhToan)), BatBuoc: 'Có',
      NguonKeHoach: 'NHANSU', MaThamChieu: ref, GhiChu: [item.vaiTro, item.lopPhuTrach ? 'Phụ trách ' + item.lopPhuTrach : '', item.ghiChu].filter(Boolean).join(' · '), TrangThai: 'ACTIVE', CreatedAt: now, UpdatedAt: now });
  });
  recurring.forEach(item => {
    const ref = 'DINHKY|' + item.maKhoanDinhKy; if (existingRefs[ref]) return;
    let amount = item.dinhMuc;
    if (item.phuongPhapTinh === 'PER_STUDENT') amount = item.dinhMuc * number_(config.soHocSinh);
    if (item.phuongPhapTinh === 'PERCENT_REVENUE') amount = item.dinhMuc * number_(config.doanhThuDuKien) / 100;
    if (item.phuongPhapTinh === 'MANUAL') amount = 0;
    rows.push({ MaKeHoachChi: 'KHCHI_' + Utilities.getUuid().slice(0, 10).toUpperCase(), MaKyHoc: session.maKyHoc, Thang: yearMonth,
      MaDanhMuc: item.maDanhMuc, TenKhoanChi: item.tenKhoanChi, NhomChi: item.nhomChi, MaNhanSu: item.maNhanSu, NguoiNhan: item.nguoiNhan, SoTienPhaiChi: amount,
      HanThanhToan: parseInputDate_(financeDueDate_(yearMonth, item.ngayThanhToan)), BatBuoc: item.batBuoc ? 'Có' : 'Không',
      NguonKeHoach: 'DINHKY', MaThamChieu: ref, GhiChu: item.ghiChu, TrangThai: 'ACTIVE', CreatedAt: now, UpdatedAt: now });
  });
  if (rows.length) {
    const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang tạo kế hoạch tài chính.');
    try { appendObjectsToSheet_(ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_KEHOACH_CHI_THANG, getKeHoachChiThangHeaders_()), rows, getKeHoachChiThangHeaders_()); } finally { lock.releaseLock(); }
  }
  bumpDataVersion_(); safeWriteAuditLog_(session, 'GENERATE', 'KE_HOACH_CHI', yearMonth, null, { soKhoan: rows.length }); return jsonResponse_({ success: true, count: rows.length, message: rows.length ? 'Đã tạo ' + rows.length + ' khoản phải chi từ danh mục.' : 'Kế hoạch đã có đủ các khoản từ danh mục, không tạo trùng.' });
}

function khoiTaoKeHoachChiThangMau(token, yearMonth) {
  return taoKeHoachTaiChinhTuDanhMuc(token, yearMonth);
}

function getDanhMucGiaDinhList_() {
  return readObjects_(SHEET_DANHMUC_GIADINH).map(row => ({
    maDanhMucGiaDinh: String(row.MaDanhMucGiaDinh || '').trim(), tenDanhMuc: String(row.TenDanhMuc || '').trim(),
    loai: String(row.Loai || 'CHI').trim().toUpperCase(), nhom: String(row.Nhom || 'KHAC').trim().toUpperCase(),
    tyLeMacDinh: Math.max(0, number_(row.TyLeMacDinh)), kieuGioiHan: String(row.KieuGioiHan || 'MAX').trim().toUpperCase(),
    thuTu: number_(row.ThuTu) || 999, trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase(), ghiChu: String(row.GhiChu || '').trim()
  })).filter(item => item.maDanhMucGiaDinh && item.tenDanhMuc).sort((a, b) => a.thuTu - b.thuTu || a.tenDanhMuc.localeCompare(b.tenDanhMuc, 'vi'));
}

function saveDanhMucGiaDinh(token, data) {
  requireSession_(token, 'finance.write'); ensureQuanLyTaiChinhSheets_(); data = data || {};
  const id = String(data.maDanhMucGiaDinh || '').trim() || ('GD_' + Utilities.getUuid().slice(0, 10).toUpperCase());
  const name = String(data.tenDanhMuc || '').trim(); const type = String(data.loai || 'CHI').trim().toUpperCase();
  if (!name) throw new Error('Vui lòng nhập tên danh mục gia đình.');
  if (['THU', 'CHI', 'TIET_KIEM'].indexOf(type) === -1) throw new Error('Loại danh mục gia đình không hợp lệ.');
  const rate = Math.max(0, Math.min(100, number_(data.tyLeMacDinh)));
  const limit = type === 'TIET_KIEM' ? 'MIN' : (type === 'THU' ? 'NONE' : 'MAX');
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật danh mục gia đình.');
  try { upsertFinanceObject_(SHEET_DANHMUC_GIADINH, getDanhMucGiaDinhHeaders_(), 'MaDanhMucGiaDinh', id, {
    TenDanhMuc: name, Loai: type, Nhom: String(data.nhom || 'KHAC').trim().toUpperCase(), TyLeMacDinh: rate,
    KieuGioiHan: limit, ThuTu: Math.max(1, number_(data.thuTu) || 99), TrangThai: String(data.trangThai || 'ACTIVE').trim().toUpperCase() === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE',
    GhiChu: String(data.ghiChu || '').trim()
  }); } finally { lock.releaseLock(); }
  bumpDataVersion_(); return jsonResponse_({ success: true, maDanhMucGiaDinh: id, message: 'Đã lưu danh mục tài chính gia đình.' });
}

function saveGiaoDichGiaDinh(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);
  ensureQuanLyTaiChinhSheets_();
  data = data || {};
  const requestedId = String(data.maGiaoDichGiaDinh || '').trim();
  const dateText = String(data.ngayGiaoDich || '').trim();
  const categoryId = String(data.maDanhMucGiaDinh || '').trim();
  const sourceCode = String(data.maNguonTien || data.nguonTien || '').trim().toUpperCase();
  const amount = moneyNumber_(data.soTien);
  const content = String(data.noiDung || '').trim();
  if (!toDateOnly_(dateText)) throw new Error('Ngày giao dịch gia đình không hợp lệ.');
  assertFinancePeriodOpen_(session, dateText.slice(0, 7));
  if (amount <= 0) throw new Error('Số tiền phải lớn hơn 0.');
  if (!content) throw new Error('Vui lòng nhập nội dung giao dịch gia đình.');
  const category = getDanhMucGiaDinhList_().find(item => item.maDanhMucGiaDinh === categoryId && item.trangThai === 'ACTIVE');
  if (!category) throw new Error('Danh mục gia đình không tồn tại hoặc đã ngừng sử dụng.');
  const source = getNguonTienList_(session.maKyHoc, 'GIA_DINH').find(item => item.maNguonTien === sourceCode && item.trangThai === 'ACTIVE');
  if (!source) throw new Error('Vui lòng chọn đúng nguồn tiền gia đình.');
  const ledgerType = category.loai === 'THU' ? 'THU' : 'CHI';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật giao dịch gia đình.');
  let savedId = '';
  let savedReceipt = '';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureSheet_(ss, SHEET_SOTHUCHI, getSoThuChiHeaders_());
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    let targetIndex = -1;
    if (requestedId) {
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][index.MaGiaoDich] || '').trim() === requestedId && String(values[i][index.MaKyHoc] || '').trim() === session.maKyHoc) {
          targetIndex = i;
          break;
        }
      }
      if (targetIndex === -1) throw new Error('Không tìm thấy giao dịch gia đình cần sửa.');
      if (normalizeFinanceScope_(values[targetIndex][index.PhamVi]) !== 'GIA_DINH') throw new Error('Giao dịch không thuộc phạm vi gia đình.');
      const existingSourceType = String(values[targetIndex][index.NguonDuLieu] || '').trim().toUpperCase();
      if (existingSourceType === 'GIA_DINH_RUT_CHU' || existingSourceType === 'GIA_DINH_CHUYEN') throw new Error('Giao dịch liên kết hoặc chuyển nguồn phải điều chỉnh từ nghiệp vụ gốc.');
      if (String(values[targetIndex][index.TrangThai] || 'HOAT_DONG').trim().toUpperCase() !== 'HOAT_DONG') throw new Error('Phiếu đã hủy không thể sửa.');
    }
    const now = new Date();
    const row = targetIndex >= 0 ? values[targetIndex].slice() : new Array(headers.length).fill('');
    const id = requestedId || ('TC_GD_' + Utilities.getUuid().slice(0, 10).toUpperCase());
    const existingReceipt = targetIndex >= 0 ? String(row[index.SoPhieu] || '').trim() : '';
    const receipt = existingReceipt || generateNextSoPhieuFromRows_(ledgerType === 'THU' ? 'THU_GIA_DINH' : 'CHI_GIA_DINH', dateText, values.slice(1), index);
    row[index.MaGiaoDich] = id;
    row[index.NgayGiaoDich] = parseInputDate_(dateText);
    row[index.MaKyHoc] = session.maKyHoc;
    row[index.LoaiGiaoDich] = ledgerType;
    row[index.MaDanhMuc] = categoryId;
    row[index.MaKeHoachChi] = '';
    row[index.MaHuTaiChinh] = '';
    row[index.TenDanhMuc] = category.tenDanhMuc;
    row[index.NoiDung] = content;
    row[index.SoTien] = amount;
    row[index.HinhThuc] = getHinhThucByNguon_(sourceCode);
    row[index.MaNguonTien] = sourceCode;
    row[index.TenNguonTien] = source.tenNguonTien;
    row[index.MaNguonDoiUng] = '';
    row[index.MaNhomChuyen] = '';
    row[index.NguoiNopNhan] = String(data.nguoiNopNhan || '').trim();
    row[index.SoPhieu] = receipt;
    row[index.SoChungTu] = String(data.soChungTu || '').trim();
    row[index.GhiChu] = String(data.ghiChu || '').trim();
    row[index.NguonDuLieu] = String(data.nguonDuLieu || 'GIA_DINH').trim().toUpperCase();
    row[index.MaThamChieu] = 'GD|' + id;
    row[index.PhamVi] = 'GIA_DINH';
    row[index.MaDanhMucGiaDinh] = categoryId;
    row[index.LoaiGiaDinh] = category.loai;
    row[index.MaGiaoDichLienKet] = String(data.maGiaoDichLienKet || '').trim();
    row[index.TrangThai] = 'HOAT_DONG';
    row[index.NguoiTao] = session.tenDangNhap || session.hoTen || session.maNguoiDung;
    row[index.CreatedAt] = targetIndex >= 0 ? (row[index.CreatedAt] || now) : now;
    row[index.UpdatedAt] = now;
    if (targetIndex >= 0) sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
    savedId = id;
    savedReceipt = receipt;
  } finally {
    lock.releaseLock();
  }
  bumpDataVersion_();
  safeWriteAuditLog_(session, requestedId ? 'UPDATE' : 'CREATE', 'GIAO_DICH_GIA_DINH', savedId, null, { thang: dateText.slice(0, 7), soTien: amount, danhMuc: categoryId, maNguonTien: sourceCode });
  return jsonResponse_({ success: true, maGiaoDichGiaDinh: savedId, maGiaoDich: savedId, soPhieu: savedReceipt, message: 'Đã lưu ' + (ledgerType === 'THU' ? 'phiếu thu ' : 'phiếu chi ') + savedReceipt + ' của gia đình.' });
}

function deleteGiaoDichGiaDinh(token, id) {
  const session = requireSession_(token, 'finance.write'); ensureThuChiSheets_(session.maKyHoc); ensureQuanLyTaiChinhSheets_();
  const ledgerRow = readObjectsNoCache_(SHEET_SOTHUCHI).find(row =>
    String(row.MaGiaoDich || '').trim() === String(id || '').trim() &&
    String(row.MaKyHoc || '').trim() === session.maKyHoc &&
    normalizeFinanceScope_(row.PhamVi || (String(row.NguonDuLieu || '').trim().toUpperCase().indexOf('GIA_DINH') === 0 ? 'GIA_DINH' : 'TRUNG_TAM')) === 'GIA_DINH'
  );
  if (ledgerRow) {
    const sourceType = String(ledgerRow.NguonDuLieu || '').trim().toUpperCase();
    if (sourceType === 'GIA_DINH_RUT_CHU') throw new Error('Khoản nhận từ trung tâm phải hủy tại danh sách rút tiền chủ sở hữu.');
    if (sourceType === 'GIA_DINH_CHUYEN') throw new Error('Khoản chuyển nguồn phải hủy theo cả nhóm chuyển tiền.');
    return cancelThuChiGiaoDich(token, id);
  }
  const rows = readObjectsNoCache_(SHEET_GIAODICH_GIADINH); const current = rows.find(row => String(row.MaGiaoDichGiaDinh || '').trim() === String(id || '').trim());
  if (!current) throw new Error('Không tìm thấy giao dịch gia đình.'); current.TrangThai = 'DELETED';
  assertFinancePeriodOpen_(session, String(current.Thang || formatDateForInput_(current.NgayGiaoDich).slice(0, 7)).trim());
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật giao dịch gia đình.');
  try { upsertFinanceObject_(SHEET_GIAODICH_GIADINH, getGiaoDichGiaDinhHeaders_(), 'MaGiaoDichGiaDinh', String(id), current); } finally { lock.releaseLock(); }
  bumpDataVersion_(); safeWriteAuditLog_(session, 'DELETE', 'GIAO_DICH_GIA_DINH', String(id), current, null); return jsonResponse_({ success: true, message: 'Đã xoá giao dịch gia đình.' });
}

function saveCauHinhGiaDinhThang(token, data) {
  const session = requireSession_(token, 'finance.write'); ensureQuanLyTaiChinhSheets_(); data = data || {}; const month = String(data.thang || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Tháng ngân sách gia đình không hợp lệ.');
  assertFinancePeriodOpen_(session, month);
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật ngân sách gia đình.');
  try { upsertFinanceObject_(SHEET_CAUHINH_GIADINH_THANG, getCauHinhGiaDinhThangHeaders_(), 'Thang', month, {
    ThuNhapKhacDuKien: Math.max(0, moneyNumber_(data.thuNhapDuKien)), QuyKhanCapMucTieu: Math.max(0, moneyNumber_(data.quyKhanCapMucTieu)), GhiChu: String(data.ghiChu || '').trim()
  }); } finally { lock.releaseLock(); }
  bumpDataVersion_(); safeWriteAuditLog_(session, 'UPSERT', 'CAU_HINH_GIA_DINH', month, null, { thuNhapDuKien: Math.max(0, moneyNumber_(data.thuNhapDuKien)), quyKhanCapMucTieu: Math.max(0, moneyNumber_(data.quyKhanCapMucTieu)) }); return jsonResponse_({ success: true, message: 'Đã lưu cấu hình ngân sách gia đình.' });
}

function getDefaultHuTaiChinh_() {
  return [
    { code: 'VAN_HANH', name: 'Hũ vận hành', ratio: 40, order: 1, note: 'Chi phí vận hành, bán trú, học cụ, thuế và chi phí khác.' },
    { code: 'LUONG', name: 'Hũ lương trung tâm', ratio: 30, order: 2, note: 'Lương giáo viên và nhân viên.' },
    { code: 'LOI_NHUAN', name: 'Hũ lợi nhuận', ratio: 10, order: 3, note: 'Phần lợi nhuận giữ riêng sau hoạt động.' },
    { code: 'DAU_TU', name: 'Hũ đầu tư phát triển', ratio: 8, order: 4, note: 'Nâng cấp cơ sở, thiết bị và phát triển dịch vụ.' },
    { code: 'DU_PHONG', name: 'Hũ quỹ dự phòng', ratio: 7, order: 5, note: 'Dự phòng biến động doanh thu và chi phí bất thường.' },
    { code: 'PHUC_LOI', name: 'Hũ phúc lợi - văn hoá', ratio: 5, order: 6, note: 'Phúc lợi, gắn kết đội ngũ và hoạt động văn hoá.' }
  ];
}

function getDanhMucHuTaiChinhList_(includeDeleted) {
  return readObjects_(SHEET_DANHMUC_HU_TAICHINH).map(row => ({
    code: normalizeMaHuTaiChinh_(row.MaHu),
    name: String(row.TenHu || '').trim(),
    ratio: Math.max(0, Math.min(100, number_(row.TyLeMacDinh))),
    order: Math.max(1, number_(row.ThuTu) || 999),
    status: String(row.TrangThai || 'ACTIVE').trim().toUpperCase(),
    note: String(row.GhiChu || '').trim(),
    systemRole: String(row.VaiTroHeThong || '').trim().toUpperCase()
  })).filter(item => item.code && item.name && (includeDeleted || item.status === 'ACTIVE'))
    .sort((a, b) => a.order - b.order || a.name.localeCompare(b.name, 'vi'));
}

function saveDanhMucHuTaiChinh(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureQuanLyTaiChinhSheets_();
  data = data || {};
  const inputCode = normalizeMaHuTaiChinh_(data.maHu || data.code);
  const name = String(data.tenHu || data.name || '').trim();
  const ratio = number_(data.tyLeMacDinh === undefined ? data.ratio : data.tyLeMacDinh);
  const order = number_(data.thuTu === undefined ? data.order : data.thuTu);
  const systemRole = String(data.vaiTroHeThong || data.systemRole || '').trim().toUpperCase() === 'OWNER_COMPENSATION' ? 'OWNER_COMPENSATION' : '';
  if (!name) throw new Error('Vui lòng nhập tên hũ tài chính.');
  if (ratio < 0 || ratio > 100) throw new Error('Tỷ lệ mặc định phải từ 0% đến 100%.');
  if (order < 1 || Math.floor(order) !== order) throw new Error('Thứ tự hũ phải là số nguyên lớn hơn hoặc bằng 1.');
  const code = inputCode || ('HU_' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase());
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật danh mục hũ tài chính.');
  try {
    const currentRows = readObjectsNoCache_(SHEET_DANHMUC_HU_TAICHINH);
    const current = inputCode ? currentRows.find(row => normalizeMaHuTaiChinh_(row.MaHu) === inputCode && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED') : null;
    if (inputCode && !current) throw new Error('Không tìm thấy hũ tài chính cần sửa.');
    const duplicated = currentRows.some(row => normalizeMaHuTaiChinh_(row.MaHu) !== inputCode && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED' && normalizeText_(row.TenHu) === normalizeText_(name));
    if (duplicated) throw new Error('Tên hũ tài chính đã tồn tại.');
    if (systemRole && currentRows.some(row => normalizeMaHuTaiChinh_(row.MaHu) !== inputCode && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() === 'ACTIVE' && String(row.VaiTroHeThong || '').trim().toUpperCase() === systemRole)) {
      throw new Error('Chỉ được có một Hũ lương chủ trung tâm đang hoạt động.');
    }
    upsertFinanceObject_(SHEET_DANHMUC_HU_TAICHINH, getDanhMucHuTaiChinhHeaders_(), 'MaHu', code, {
      TenHu: name, TyLeMacDinh: ratio, ThuTu: order, TrangThai: 'ACTIVE',
      GhiChu: String(data.ghiChu || data.note || '').trim(), VaiTroHeThong: systemRole
    });
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  safeWriteAuditLog_(session, inputCode ? 'UPDATE' : 'CREATE', 'DANH_MUC_HU_TAI_CHINH', code, null, { tenHu: name, tyLeMacDinh: ratio, thuTu: order });
  return jsonResponse_({ success: true, maHu: code, message: inputCode ? 'Đã cập nhật hũ tài chính.' : 'Đã thêm hũ tài chính.' });
}

function deleteDanhMucHuTaiChinh(token, codeInput) {
  const session = requireSession_(token, 'finance.write');
  ensureQuanLyTaiChinhSheets_();
  const code = normalizeMaHuTaiChinh_(codeInput);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật danh mục hũ tài chính.');
  let current;
  try {
    const rows = readObjectsNoCache_(SHEET_DANHMUC_HU_TAICHINH);
    current = rows.find(row => normalizeMaHuTaiChinh_(row.MaHu) === code && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED');
    if (!current) throw new Error('Không tìm thấy hũ tài chính cần xoá.');
    const activeCount = rows.filter(row => normalizeMaHuTaiChinh_(row.MaHu) && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() === 'ACTIVE').length;
    if (activeCount <= 1) throw new Error('Phải giữ lại ít nhất một hũ tài chính đang hoạt động.');
    const usedByCategory = readObjectsNoCache_(SHEET_DANHMUC_THUCHI).some(row => normalizeMaHuTaiChinh_(row.MaHuMacDinh) === code && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED');
    const usedByTransaction = readObjectsNoCache_(SHEET_SOTHUCHI).some(row => normalizeMaHuTaiChinh_(row.MaHuTaiChinh) === code);
    const usedByConfig = readObjectsNoCache_(SHEET_HU_TAICHINH_THANG).some(row => normalizeMaHuTaiChinh_(row.MaHu) === code);
    if (usedByCategory || usedByTransaction || usedByConfig) throw new Error('Hũ đã được danh mục, phiếu chi hoặc cấu hình tháng sử dụng nên không thể xoá. Hãy chuyển các liên kết sang hũ khác trước để bảo toàn lịch sử tài chính.');
    current.TrangThai = 'DELETED';
    upsertFinanceObject_(SHEET_DANHMUC_HU_TAICHINH, getDanhMucHuTaiChinhHeaders_(), 'MaHu', code, current);
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  safeWriteAuditLog_(session, 'DELETE', 'DANH_MUC_HU_TAI_CHINH', code, current, null);
  return jsonResponse_({ success: true, message: 'Đã xoá hũ tài chính chưa phát sinh dữ liệu.' });
}

function getChotPhanBoHu_(maKyHoc, month) {
  const rows = readObjectsNoCache_(SHEET_CHOT_PHANBO_HU);
  const lockId = String(maKyHoc || '').trim() + '|' + String(month || '').trim();
  const row = rows.find(item =>
    String(item.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
    String(item.Thang || '').trim() === String(month || '').trim()
  ) || rows.find(item => String(item.MaChot || '').trim() === lockId) || {};
  const locked = String(row.TrangThai || 'DRAFT').trim().toUpperCase() === 'LOCKED';
  const today = new Date();
  const isCurrentMonth = formatDateForInput_(today).slice(0, 7) === month;
  return {
    month: month,
    locked: locked,
    status: locked ? 'LOCKED' : 'DRAFT',
    baseTuition: number_(row.HocPhiCoSo),
    lockedAt: formatDateDisplay_(row.NgayChot),
    lockedBy: String(row.NguoiChot || '').trim(),
    unlockedAt: formatDateDisplay_(row.NgayMoKhoa),
    unlockedBy: String(row.NguoiMoKhoa || '').trim(),
    unlockReason: String(row.LyDoMoKhoa || '').trim(),
    version: Math.max(0, number_(row.PhienBan)),
    overdue: !locked && isCurrentMonth && today.getDate() > 5
  };
}

function assertPhanBoHuChuaChot_(session, month) {
  if (getChotPhanBoHu_(session.maKyHoc, month).locked) {
    throw new Error('Phương án phân bổ hũ tháng này đã chốt. Chỉ Chủ sở hữu mới có thể mở chốt và điều chỉnh lại.');
  }
}

function getHocPhiPhaiThuThang_(maKyHoc, month) {
  const parts = String(month || '').split('-');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getThuPhiMonthSheetName_(Number(parts[0]), Number(parts[1])));
  if (!sheet) return 0;
  return readObjectsNoCache_(sheet.getName()).filter(row =>
    String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
    String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED' &&
    !toBoolean_(row.TamNghi)
  ).reduce((sum, row) => sum + Math.max(0, number_(row.HocPhi)), 0);
}

function lockPhanBoHuTaiChinh(token, month) {
  const session = requireSession_(token, 'finance.close');
  ensureThuChiSheets_(session.maKyHoc);
  month = GovernanceService.validateMonth(month);
  assertFinancePeriodOpen_(session, month);
  if (session.vaiTro !== 'OWNER') throw new Error('Chỉ Chủ sở hữu được chốt phương án phân bổ hũ tài chính.');
  const current = getChotPhanBoHu_(session.maKyHoc, month);
  if (current.locked) return jsonResponse_({ success: true, status: current, message: 'Phương án phân bổ đã được chốt trước đó.' });
  const jars = getHuTaiChinhThang_(session.maKyHoc, month);
  const ratioTotal = jars.reduce((sum, item) => sum + number_(item.ratio), 0);
  if (Math.abs(ratioTotal - 100) > 0.01) throw new Error('Tổng tỷ lệ các hũ phải bằng đúng 100% trước khi chốt.');
  const ownerJars = getDanhMucHuTaiChinhList_(false).filter(item => item.systemRole === 'OWNER_COMPENSATION');
  if (ownerJars.length !== 1) throw new Error('Cần đánh dấu đúng một hũ là “Hũ lương chủ trung tâm” trước khi chốt.');
  const baseTuition = getHocPhiPhaiThuThang_(session.maKyHoc, month);
  if (baseTuition <= 0) throw new Error('Chưa có học phí phải thu của tháng nên chưa thể chốt phương án phân bổ.');
  saveHuTaiChinhThang(token, { thang: month, jars: jars.map(item => ({ code: item.code, ratio: item.ratio, note: item.note })) });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang chốt phương án phân bổ hũ.');
  try {
    const lockId = session.maKyHoc + '|' + month;
    upsertFinanceObject_(SHEET_CHOT_PHANBO_HU, getChotPhanBoHuHeaders_(), 'MaChot', lockId, {
      MaKyHoc: session.maKyHoc, Thang: month, TrangThai: 'LOCKED', HocPhiCoSo: baseTuition,
      NgayChot: new Date(), NguoiChot: session.tenDangNhap || session.hoTen || session.maNguoiDung,
      NgayMoKhoa: '', NguoiMoKhoa: '', LyDoMoKhoa: '', PhienBan: current.version + 1
    });
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  const status = getChotPhanBoHu_(session.maKyHoc, month);
  safeWriteAuditLog_(session, 'LOCK', 'PHAN_BO_HU', month, current, status, { hocPhiCoSo: baseTuition, tongTyLe: ratioTotal });
  return jsonResponse_({ success: true, status: status, message: 'Đã chốt phương án phân bổ hũ theo học phí phải thu tại thời điểm chốt.' });
}

function unlockPhanBoHuTaiChinh(token, month, reason) {
  const session = requireSession_(token, 'finance.close');
  ensureThuChiSheets_(session.maKyHoc);
  month = GovernanceService.validateMonth(month);
  assertFinancePeriodOpen_(session, month);
  if (session.vaiTro !== 'OWNER') throw new Error('Chỉ Chủ sở hữu được mở chốt phương án phân bổ hũ tài chính.');
  reason = String(reason || '').trim();
  if (!reason) throw new Error('Vui lòng nhập lý do mở chốt để lưu nhật ký kiểm soát.');
  const current = getChotPhanBoHu_(session.maKyHoc, month);
  if (!current.locked) return jsonResponse_({ success: true, status: current, message: 'Phương án phân bổ đang ở trạng thái bản nháp.' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang mở chốt phương án phân bổ hũ.');
  try {
    const lockId = session.maKyHoc + '|' + month;
    upsertFinanceObject_(SHEET_CHOT_PHANBO_HU, getChotPhanBoHuHeaders_(), 'MaChot', lockId, {
      MaKyHoc: session.maKyHoc, Thang: month, TrangThai: 'DRAFT', HocPhiCoSo: current.baseTuition,
      NgayMoKhoa: new Date(), NguoiMoKhoa: session.tenDangNhap || session.hoTen || session.maNguoiDung,
      LyDoMoKhoa: reason, PhienBan: current.version
    });
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  const status = getChotPhanBoHu_(session.maKyHoc, month);
  safeWriteAuditLog_(session, 'UNLOCK', 'PHAN_BO_HU', month, current, status, { lyDo: reason });
  return jsonResponse_({ success: true, status: status, message: 'Đã mở chốt. Mọi thay đổi tiếp theo sẽ được ghi nhận ở phiên bản phân bổ mới.' });
}

function getHuTaiChinhThang_(maKyHoc, month) {
  const saved = readObjectsNoCache_(SHEET_HU_TAICHINH_THANG).filter(row =>
    String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
    String(row.Thang || '').trim() === month
  );
  const savedMap = saved.reduce((map, row) => {
    map[String(row.MaHu || '').trim().toUpperCase()] = row;
    return map;
  }, {});
  const catalog = getDanhMucHuTaiChinhList_(false);
  const catalogMap = catalog.reduce((map, item) => { map[item.code] = item; return map; }, {});
  if (getChotPhanBoHu_(maKyHoc, month).locked && saved.length) {
    return saved.map(row => {
      const code = normalizeMaHuTaiChinh_(row.MaHu);
      const definition = catalogMap[code] || {};
      return {
        code: code,
        name: String(row.TenHu || definition.name || code).trim(),
        ratio: Math.max(0, Math.min(100, number_(row.TyLePhanBo))),
        order: number_(row.ThuTu) || number_(definition.order) || 999,
        note: String(row.GhiChu || definition.note || '').trim(),
        systemRole: String(row.VaiTroHeThong || definition.systemRole || '').trim().toUpperCase()
      };
    }).filter(item => item.code).sort((a, b) => a.order - b.order);
  }
  return catalog.map(item => {
    const row = savedMap[item.code] || {};
    return {
      code: item.code,
      name: item.name,
      ratio: row.TyLePhanBo === '' || row.TyLePhanBo === undefined ? item.ratio : Math.max(0, Math.min(100, number_(row.TyLePhanBo))),
      order: number_(row.ThuTu) || item.order,
      note: item.note,
      systemRole: item.systemRole
    };
  }).sort((a, b) => a.order - b.order);
}

function saveHuTaiChinhThang(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureQuanLyTaiChinhSheets_();
  data = data || {};
  const month = String(data.thang || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Tháng cấu hình hũ tài chính không hợp lệ.');
  assertFinancePeriodOpen_(session, month);
  assertPhanBoHuChuaChot_(session, month);
  const definitions = getDanhMucHuTaiChinhList_(false);
  const allowed = definitions.reduce((map, item) => { map[item.code] = item; return map; }, {});
  const incoming = Array.isArray(data.jars) ? data.jars : [];
  if (!definitions.length || incoming.length !== definitions.length) throw new Error('Cần cấu hình đầy đủ tất cả hũ tài chính đang hoạt động.');
  const inputMap = {};
  incoming.forEach(item => {
    const code = String(item && (item.code || item.maHu) || '').trim().toUpperCase();
    if (!allowed[code] || inputMap[code]) throw new Error('Danh sách hũ tài chính không hợp lệ hoặc bị trùng.');
    const ratio = number_(item.ratio === undefined ? item.tyLePhanBo : item.ratio);
    if (ratio < 0 || ratio > 100) throw new Error('Tỷ lệ mỗi hũ phải từ 0% đến 100%.');
    inputMap[code] = { ratio: ratio, note: String(item.note || item.ghiChu || '').trim() };
  });
  const ratioTotal = definitions.reduce((sum, item) => sum + inputMap[item.code].ratio, 0);
  if (Math.abs(ratioTotal - 100) > 0.01) throw new Error('Tổng tỷ lệ các hũ phải bằng 100%. Hiện tại là ' + ratioTotal.toFixed(1) + '%.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật hũ tài chính.');
  try {
    const sheet = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_HU_TAICHINH_THANG, getHuTaiChinhThangHeaders_());
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(item => String(item || '').trim());
    const index = buildHeaderIndex_(headers);
    const now = new Date();
    definitions.forEach(definition => {
      let target = -1;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][index.MaKyHoc] || '').trim() === session.maKyHoc &&
            String(rows[i][index.Thang] || '').trim() === month &&
            String(rows[i][index.MaHu] || '').trim().toUpperCase() === definition.code) {
          target = i;
          break;
        }
      }
      const row = target >= 0 ? rows[target].slice() : new Array(headers.length).fill('');
      row[index.MaKyHoc] = session.maKyHoc;
      row[index.Thang] = month;
      row[index.MaHu] = definition.code;
      row[index.TenHu] = definition.name;
      row[index.TyLePhanBo] = inputMap[definition.code].ratio;
      row[index.ThuTu] = definition.order;
      row[index.GhiChu] = inputMap[definition.code].note || definition.note;
      row[index.VaiTroHeThong] = definition.systemRole || '';
      row[index.CreatedAt] = target >= 0 ? (row[index.CreatedAt] || now) : now;
      row[index.UpdatedAt] = now;
      if (target >= 0) {
        rows[target] = row;
        sheet.getRange(target + 1, 1, 1, headers.length).setValues([row]);
      } else {
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
        rows.push(row);
      }
    });
  } finally { lock.releaseLock(); }
  bumpDataVersion_();
  safeWriteAuditLog_(session, 'UPSERT', 'HU_TAI_CHINH', month, null, { tongTyLe: ratioTotal, soHu: definitions.length });
  return jsonResponse_({ success: true, message: 'Đã lưu tỷ lệ phân bổ các hũ tài chính trong tháng.' });
}

function getHuTaiChinhCodeForTransaction_(transaction, planGroupMap, categoryJarMap) {
  const storedJar = normalizeMaHuTaiChinh_(transaction.maHuTaiChinh);
  if (storedJar) return storedJar;
  const categoryJar = normalizeMaHuTaiChinh_(categoryJarMap[transaction.maDanhMuc]);
  if (categoryJar) return categoryJar;
  const plannedGroup = String(planGroupMap[transaction.maKeHoachChi] || '').trim().toUpperCase();
  if (plannedGroup === 'LUONG') return 'LUONG';
  const text = normalizeText_([transaction.maDanhMuc, transaction.tenDanhMuc, transaction.noiDung, plannedGroup].join(' '));
  if (text.indexOf('loi nhuan') !== -1) return 'LOI_NHUAN';
  if (text.indexOf('dau tu') !== -1 || text.indexOf('phat trien') !== -1) return 'DAU_TU';
  if (text.indexOf('du phong') !== -1) return 'DU_PHONG';
  if (text.indexOf('phuc loi') !== -1 || text.indexOf('van hoa') !== -1) return 'PHUC_LOI';
  if (text.indexOf('luong') !== -1 || String(transaction.maDanhMuc || '').toUpperCase() === 'CHI_LUONG') return 'LUONG';
  return 'VAN_HANH';
}

function buildHuTaiChinhData_(maKyHoc, month, feeSummary, allTransactions, planData) {
  const jars = getHuTaiChinhThang_(maKyHoc, month);
  const allocationStatus = getChotPhanBoHu_(maKyHoc, month);
  const categoryJarMap = readObjectsNoCache_(SHEET_DANHMUC_THUCHI).reduce((map, row) => {
    map[String(row.MaDanhMuc || '').trim()] = normalizeMaHuTaiChinh_(row.MaHuMacDinh);
    return map;
  }, {});
  const planGroupMap = (planData.items || []).reduce((map, item) => {
    map[item.maKeHoachChi] = item.nhomChi;
    return map;
  }, {});
  const actualByJar = {};
  let currentTuitionRevenue = 0;
  (allTransactions || []).forEach(transaction => {
    if (transaction.nguonDuLieu === 'CHUYEN_NOI_BO' || !transaction.ngayGiaoDich) return;
    const transactionMonth = transaction.ngayGiaoDich.slice(0, 7);
    if (transactionMonth !== month) return;
    const isTuition = transaction.loai === 'THU' && transaction.maDanhMuc === 'THU_HOC_PHI';
    if (isTuition) {
      currentTuitionRevenue += number_(transaction.soTien);
      return;
    }
    if (transaction.loai !== 'CHI') return;
    const code = getHuTaiChinhCodeForTransaction_(transaction, planGroupMap, categoryJarMap);
    actualByJar[code] = (actualByJar[code] || 0) + number_(transaction.soTien);
  });
  const currentExpectedRevenue = number_(feeSummary.expected);
  const plannedTuitionRevenue = allocationStatus.locked ? number_(allocationStatus.baseTuition) : currentExpectedRevenue;
  const result = FinanceDomain.calculateJars({ revenue: plannedTuitionRevenue, jars: jars, actualByJar: actualByJar });
  const warnings = [];
  if (allocationStatus.overdue) warnings.push({ level: 'warning', message: 'Đã qua ngày 05 nhưng phương án phân bổ hũ tháng này chưa được chốt.' });
  if (allocationStatus.locked && currentExpectedRevenue > plannedTuitionRevenue) warnings.push({ level: 'info', message: 'Có ' + formatMoneyText_(currentExpectedRevenue - plannedTuitionRevenue) + ' học phí phát sinh sau chốt. Khoản này đang để ở trạng thái chưa phân bổ, không tự động chuyển sang tài chính gia đình.' });
  if (allocationStatus.locked && currentExpectedRevenue < plannedTuitionRevenue) warnings.push({ level: 'warning', message: 'Học phí phải thu hiện thấp hơn số đã chốt ' + formatMoneyText_(plannedTuitionRevenue - currentExpectedRevenue) + '. Hãy kiểm tra học sinh nghỉ hoặc điều chỉnh học phí.' });
  if (Math.abs(result.summary.ratioTotal - 100) > 0.01) warnings.push({ level: 'danger', message: 'Tổng tỷ lệ các hũ đang là ' + result.summary.ratioTotal.toFixed(1) + '%, cần điều chỉnh về 100%.' });
  if (!plannedTuitionRevenue) warnings.push({ level: 'warning', message: 'Tháng này chưa có dữ liệu học phí phải thu nên chưa thể lập ngân sách cho các hũ.' });
  if (Math.abs(number_(feeSummary.collected) - currentTuitionRevenue) > 1) warnings.push({ level: 'warning', message: 'Học phí đã thu trên danh sách (' + formatMoneyText_(feeSummary.collected) + ') chưa khớp Sổ thu chi (' + formatMoneyText_(currentTuitionRevenue) + '). Hãy đồng bộ học phí trước khi sử dụng báo cáo hũ.' });
  if (result.summary.actualTotal > currentTuitionRevenue && result.summary.actualTotal > 0) warnings.push({ level: 'warning', message: 'Tổng chi theo hũ đang lớn hơn học phí đã ghi sổ trong tháng. Đây là cảnh báo dòng tiền, không làm thay đổi ngân sách kế hoạch.' });
  result.items.forEach(item => {
    const definition = jars.find(jar => jar.code === item.code) || {};
    item.systemRole = definition.systemRole || '';
    if (item.remaining < 0) warnings.push({ level: 'danger', message: item.name + ' đã vượt ngân sách kế hoạch ' + formatMoneyText_(Math.abs(item.remaining)) + '.' });
    else if (item.usedPercent >= 90) warnings.push({ level: 'warning', message: item.name + ' đã sử dụng ' + item.usedPercent.toFixed(1) + '% ngân sách tháng.' });
  });
  result.expectedRevenue = plannedTuitionRevenue;
  result.currentExpectedRevenue = currentExpectedRevenue;
  result.postLockRevenue = allocationStatus.locked ? Math.max(currentExpectedRevenue - plannedTuitionRevenue, 0) : 0;
  result.revenueReductionAfterLock = allocationStatus.locked ? Math.max(plannedTuitionRevenue - currentExpectedRevenue, 0) : 0;
  result.actualCollected = currentTuitionRevenue;
  result.collectionRate = plannedTuitionRevenue > 0 ? currentTuitionRevenue * 100 / plannedTuitionRevenue : 0;
  result.catalog = getDanhMucHuTaiChinhList_(false);
  result.allocationStatus = allocationStatus;
  result.warnings = warnings;
  return result;
}

function saveRutTienChuSoHuu(token, data) {
  const session = requireSession_(token, 'finance.write');
  if (session.vaiTro !== 'OWNER') throw new Error('Chỉ Chủ sở hữu được xác nhận rút tiền từ trung tâm về gia đình.');
  ensureThuChiSheets_(session.maKyHoc);
  data = data || {};
  const dateText = String(data.ngayGiaoDich || '').trim();
  if (!toDateOnly_(dateText)) throw new Error('Ngày rút tiền không hợp lệ.');
  const month = dateText.slice(0, 7);
  assertFinancePeriodOpen_(session, month);
  const allocationStatus = getChotPhanBoHu_(session.maKyHoc, month);
  if (!allocationStatus.locked) throw new Error('Cần chốt phương án phân bổ hũ trước khi rút tiền cho gia đình.');
  const ownerJar = getHuTaiChinhThang_(session.maKyHoc, month).find(item => item.systemRole === 'OWNER_COMPENSATION');
  if (!ownerJar) throw new Error('Chưa cấu hình Hũ lương chủ trung tâm.');
  const amount = moneyNumber_(data.soTien);
  if (amount <= 0) throw new Error('Số tiền rút phải lớn hơn 0.');
  const allocated = number_(allocationStatus.baseTuition) * number_(ownerJar.ratio) / 100;
  const activeTransactions = readObjectsNoCache_(SHEET_SOTHUCHI)
    .filter(row => String(row.MaKyHoc || '').trim() === session.maKyHoc)
    .map(mapThuChiTransaction_)
    .filter(item => item.trangThai === 'HOAT_DONG' && item.phamVi === 'TRUNG_TAM');
  const withdrawn = activeTransactions.filter(item =>
    item.trangThai === 'HOAT_DONG' && item.loai === 'CHI' && item.ngayGiaoDich.slice(0, 7) === month &&
    item.maDanhMuc === 'CHI_GIA_DINH' && item.maHuTaiChinh === ownerJar.code
  ).reduce((sum, item) => sum + number_(item.soTien), 0);
  if (withdrawn + amount > allocated + 1) {
    throw new Error('Số tiền rút vượt hạn mức Hũ lương chủ trung tâm còn lại ' + formatMoneyText_(Math.max(allocated - withdrawn, 0)) + '.');
  }
  const sourceCode = String(data.maNguonTien || '').trim().toUpperCase();
  const source = getNguonTienList_(session.maKyHoc).find(item => item.maNguonTien === sourceCode && item.trangThai === 'ACTIVE');
  if (!source) throw new Error('Nguồn tiền trung tâm không hợp lệ hoặc đã ngừng sử dụng.');
  const familySourceCode = String(data.maNguonTienGiaDinh || '').trim().toUpperCase();
  const familySource = getNguonTienList_(session.maKyHoc, 'GIA_DINH').find(item => item.maNguonTien === familySourceCode && item.trangThai === 'ACTIVE');
  if (!familySource) throw new Error('Vui lòng chọn nguồn tiền gia đình sẽ nhận khoản rút này.');
  const sourceBalance = number_(source.soDuBanDau) + activeTransactions.filter(item =>
    item.maNguonTien === sourceCode && item.ngayGiaoDich && item.ngayGiaoDich <= dateText
  ).reduce((sum, item) => sum + (item.loai === 'THU' ? number_(item.soTien) : -number_(item.soTien)), 0);
  if (amount > sourceBalance + 1) throw new Error('Nguồn ' + source.tenNguonTien + ' chỉ còn ' + formatMoneyText_(Math.max(sourceBalance, 0)) + ', không đủ để thực hiện đợt rút này.');
  const centerResult = JSON.parse(saveThuChiGiaoDich(token, {
    loai: 'CHI', ngayGiaoDich: dateText, maNguonTien: sourceCode,
    maDanhMuc: 'CHI_GIA_DINH', maHuTaiChinh: ownerJar.code, soTien: amount,
    noiDung: String(data.noiDung || 'Rút tiền chủ sở hữu về tài chính gia đình').trim(),
    nguoiNopNhan: String(data.nguoiNhan || session.hoTen || 'Chủ cơ sở').trim(),
    soChungTu: String(data.soChungTu || '').trim(),
    ghiChu: String(data.ghiChu || '').trim()
  }));
  try {
    const familyResult = JSON.parse(saveGiaoDichGiaDinh(token, {
      ngayGiaoDich: dateText,
      maDanhMucGiaDinh: 'GD_THU_CHU_SO_HUU',
      maNguonTien: familySourceCode,
      soTien: amount,
      noiDung: String(data.noiDung || 'Nhận tiền từ trung tâm').trim(),
      nguoiNopNhan: 'Trung tâm',
      soChungTu: centerResult.soPhieu || '',
      ghiChu: String(data.ghiChu || '').trim(),
      nguonDuLieu: 'GIA_DINH_RUT_CHU',
      maGiaoDichLienKet: centerResult.maGiaoDich
    }));
    return jsonResponse_({
      success: true,
      maGiaoDich: centerResult.maGiaoDich,
      soPhieu: centerResult.soPhieu,
      maGiaoDichGiaDinh: familyResult.maGiaoDichGiaDinh,
      message: 'Đã ghi nhận rút tiền từ trung tâm và thu tiền vào nguồn ' + familySource.tenNguonTien + '.'
    });
  } catch (error) {
    try { cancelThuChiGiaoDich(token, centerResult.maGiaoDich); } catch (rollbackError) { /* Giữ audit để quản trị viên đối chiếu. */ }
    throw error;
  }
}

function cancelRutTienChuSoHuu(token, transactionId) {
  const session = requireSession_(token, 'finance.write');
  if (session.vaiTro !== 'OWNER') throw new Error('Chỉ Chủ sở hữu được hủy đợt rút tiền về gia đình.');
  const row = readObjectsNoCache_(SHEET_SOTHUCHI).find(item =>
    String(item.MaGiaoDich || '').trim() === String(transactionId || '').trim() &&
    String(item.MaKyHoc || '').trim() === session.maKyHoc
  );
  if (!row || String(row.MaDanhMuc || '').trim() !== 'CHI_GIA_DINH') throw new Error('Không tìm thấy đợt rút tiền cần hủy.');
  const transactionMonth = formatDateForInput_(toDateOnly_(row.NgayGiaoDich)).slice(0, 7);
  const ownerJar = getHuTaiChinhThang_(session.maKyHoc, transactionMonth).find(item => item.systemRole === 'OWNER_COMPENSATION');
  if (!ownerJar || normalizeMaHuTaiChinh_(row.MaHuTaiChinh) !== ownerJar.code) throw new Error('Giao dịch không thuộc Hũ lương chủ trung tâm.');
  const linkedFamily = readObjectsNoCache_(SHEET_SOTHUCHI).find(item =>
    String(item.MaKyHoc || '').trim() === session.maKyHoc &&
    String(item.MaGiaoDichLienKet || '').trim() === String(transactionId || '').trim() &&
    normalizeFinanceScope_(item.PhamVi || (String(item.NguonDuLieu || '').trim().toUpperCase().indexOf('GIA_DINH') === 0 ? 'GIA_DINH' : 'TRUNG_TAM')) === 'GIA_DINH' &&
    String(item.TrangThai || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG'
  );
  if (linkedFamily) cancelThuChiGiaoDich(token, String(linkedFamily.MaGiaoDich || '').trim());
  return cancelThuChiGiaoDich(token, transactionId);
}

function buildFinanceFeeSummary_(maKyHoc, ym) {
  const feeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getThuPhiMonthSheetName_(ym.year, ym.month));
  const rows = feeSheet ? readObjectsNoCache_(feeSheet.getName()).filter(row =>
    String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
    String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED'
  ) : [];
  const summary = { totalStudents: 0, payingStudents: 0, expected: 0, collected: 0, remaining: 0, classes: [] };
  const classMap = {};
  rows.forEach(row => {
    const fee = toBoolean_(row.TamNghi) ? 0 : number_(row.HocPhi);
    const collected = toBoolean_(row.TamNghi) ? 0 : number_(row.SoTienDaThu);
    if (fee <= 0) return;
    summary.totalStudents++;
    summary.payingStudents++;
    summary.expected += fee;
    summary.collected += collected;
    summary.remaining += Math.max(fee - collected, 0);
    const key = String(row.Lop || row.TenLop || 'Chưa xếp lớp').trim();
    if (!classMap[key]) classMap[key] = { lop: key, tenLop: String(row.TenLop || key).trim(), students: 0, expected: 0, collected: 0 };
    classMap[key].students++;
    classMap[key].expected += fee;
    classMap[key].collected += collected;
  });
  summary.classes = Object.keys(classMap).map(key => classMap[key]).sort((a, b) => String(a.lop).localeCompare(String(b.lop), 'vi', { numeric: true }));
  return summary;
}

function getActiveFinanceTransactions_(maKyHoc) {
  return readObjectsNoCache_(SHEET_SOTHUCHI)
    .filter(row => String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim())
    .map(mapThuChiTransaction_)
    .filter(item => item.maGiaoDich && item.ngayDate && item.trangThai === 'HOAT_DONG' && item.phamVi === 'TRUNG_TAM');
}

function getFinanceMonthContext_(session, yearMonth, options) {
  options = options || {};
  const ym = parseYearMonth_(yearMonth);
  const month = ym.year + '-' + String(ym.month).padStart(2, '0');
  const monthStart = parseInputDate_(month + '-01');
  const monthEnd = new Date(ym.year, ym.month, 0, 23, 59, 59, 999);
  const fee = options.fee === false ? null : buildFinanceFeeSummary_(session.maKyHoc, ym);
  const transactions = options.transactions === false ? [] : getActiveFinanceTransactions_(session.maKyHoc);
  const monthTransactions = transactions.filter(item =>
    item.ngayDate.getTime() >= monthStart.getTime() && item.ngayDate.getTime() <= monthEnd.getTime() &&
    item.nguonDuLieu !== 'CHUYEN_NOI_BO'
  );
  return { ym: ym, month: month, monthStart: monthStart, monthEnd: monthEnd, fee: fee, transactions: transactions, monthTransactions: monthTransactions };
}

function getFinanceCurrentCash_(maKyHoc, transactions, monthEnd) {
  return getNguonTienList_(maKyHoc).reduce((sum, source) => sum + number_(source.soDuBanDau), 0) +
    (transactions || []).filter(item => item.ngayDate.getTime() <= monthEnd.getTime())
      .reduce((sum, item) => sum + (item.loai === 'THU' ? item.soTien : -item.soTien), 0);
}

function getFinanceExpenseCategories_() {
  return readObjects_(SHEET_DANHMUC_THUCHI).filter(row =>
    String(row.Loai || '').trim().toUpperCase() === 'CHI' &&
    String(row.TrangThai || 'ACTIVE').trim().toUpperCase() === 'ACTIVE'
  ).map(row => ({
    maDanhMuc: String(row.MaDanhMuc || '').trim(),
    tenDanhMuc: String(row.TenDanhMuc || '').trim()
  }));
}

function getFinanceReceivers_() {
  return readObjects_(SHEET_DOITUONG_THUCHI).filter(row => {
    const type = String(row.Loai || 'BOTH').trim().toUpperCase();
    const status = String(row.TrangThai || 'ACTIVE').trim().toUpperCase();
    return (type === 'CHI' || type === 'BOTH') && status === 'ACTIVE';
  }).map(row => ({
    maDoiTuong: String(row.MaDoiTuong || '').trim(),
    loai: String(row.Loai || 'BOTH').trim().toUpperCase(),
    tenDoiTuong: String(row.TenDoiTuong || '').trim(),
    soDienThoai: String(row.SoDienThoai || '').trim(),
    diaChi: String(row.DiaChi || '').trim(),
    ghiChu: String(row.GhiChu || '').trim()
  })).filter(item => item.maDoiTuong && item.tenDoiTuong)
    .sort((a, b) => a.tenDoiTuong.localeCompare(b.tenDoiTuong, 'vi'));
}

function getTongQuanTaiChinhData(token, yearMonth) {
  const session = requireSession_(token, 'finance.read');
  ensureThuChiSheets_(session.maKyHoc);
  ensureQuanLyTaiChinhSheets_();
  const context = getFinanceMonthContext_(session, yearMonth);
  const monthTransactions = context.monthTransactions;
  const cashIncome = monthTransactions.reduce((sum, item) => sum + (item.loai === 'THU' ? item.soTien : 0), 0);
  const businessCashExpense = monthTransactions.reduce((sum, item) => sum + (item.loai === 'CHI' && item.maDanhMuc !== 'CHI_GIA_DINH' ? item.soTien : 0), 0);
  const ownerDraw = monthTransactions.reduce((sum, item) => sum + (item.loai === 'CHI' && item.maDanhMuc === 'CHI_GIA_DINH' ? item.soTien : 0), 0);
  const otherIncome = monthTransactions.reduce((sum, item) => sum + (item.loai === 'THU' && item.nguonDuLieu !== 'HOC_PHI' ? item.soTien : 0), 0);
  const currentCash = getFinanceCurrentCash_(session.maKyHoc, context.transactions, context.monthEnd);
  const plan = buildKeHoachChiThangData_(session.maKyHoc, context.month, context.transactions, currentCash, context.fee.expected || cashIncome);
  const businessPlans = plan.items.filter(item => item.nhomChi !== 'GIA_DINH');
  const plannedExpense = businessPlans.reduce((sum, item) => sum + item.soTienPhaiChi, 0);
  const remainingObligations = businessPlans.reduce((sum, item) => sum + item.conPhaiChi, 0);
  const config = plan.config || getCauHinhTaiChinhThang_(session.maKyHoc, context.month);
  const revenueForecast = number_(config.doanhThuDuKien) || context.fee.expected || cashIncome;
  const accruedRevenue = context.fee.expected + otherIncome;
  const studentCount = number_(config.soHocSinh) || context.fee.payingStudents;
  const currentAverageFee = context.fee.payingStudents > 0 ? context.fee.expected / context.fee.payingStudents : (studentCount > 0 ? revenueForecast / studentCount : 0);
  const variableTotal = businessPlans.filter(item => item.nhomChi === 'BAN_TRU' || item.nhomChi === 'HOC_CU').reduce((sum, item) => sum + item.soTienPhaiChi, 0);
  const pricing = FinanceDomain.calculatePricing({
    studentCount: studentCount, currentAverageFee: currentAverageFee, plannedExpense: plannedExpense,
    variableExpense: variableTotal, targetMargin: Math.max(0, Math.min(80, number_(config.tyLeLoiNhuanMucTieu) || 20)), collectionRate: 100
  });
  const performance = FinanceDomain.calculatePerformance({
    cashIncome: cashIncome, cashExpense: businessCashExpense, currentCash: currentCash,
    revenueForecast: revenueForecast, accruedRevenue: accruedRevenue, plannedExpense: plannedExpense,
    remainingObligations: remainingObligations, reserveTarget: number_(plan.summary.reserveTarget), ownerDraw: ownerDraw
  });
  return jsonResponse_({
    month: context.month, fee: context.fee, plan: { summary: plan.summary, warnings: plan.warnings },
    performance: performance, pricing: pricing, periodLock: GovernanceService.periodStatus(session, context.month)
  });
}

function getKeHoachTaiChinhData(token, yearMonth, section) {
  const session = requireSession_(token, 'finance.read');
  ensureThuChiSheets_(session.maKyHoc);
  ensureQuanLyTaiChinhSheets_();
  const requested = String(section || 'jars').trim().toLowerCase();
  const allowed = ['jars', 'staff', 'recurring', 'plan'];
  const activeSection = allowed.indexOf(requested) === -1 ? 'jars' : requested;
  const ym = parseYearMonth_(yearMonth);
  const month = ym.year + '-' + String(ym.month).padStart(2, '0');
  const base = { month: month, section: activeSection, periodLock: GovernanceService.periodStatus(session, month) };
  if (activeSection === 'staff') {
    base.staff = getNhanSuTaiChinhList_(session.maKyHoc);
    base.categories = getFinanceExpenseCategories_();
    return jsonResponse_(base);
  }
  if (activeSection === 'recurring') {
    base.recurringExpenses = getKhoanChiDinhKyList_(session.maKyHoc, month);
    base.categories = getFinanceExpenseCategories_();
    base.staff = getNhanSuTaiChinhList_(session.maKyHoc);
    return jsonResponse_(base);
  }
  const context = getFinanceMonthContext_(session, month);
  const currentCash = getFinanceCurrentCash_(session.maKyHoc, context.transactions, context.monthEnd);
  const plan = buildKeHoachChiThangData_(session.maKyHoc, month, context.transactions, currentCash, context.fee.expected);
  if (activeSection === 'jars') {
    base.jars = buildHuTaiChinhData_(session.maKyHoc, month, context.fee, context.transactions, plan);
    base.jars.canManageAllocation = session.vaiTro === 'OWNER';
    return jsonResponse_(base);
  }
  base.config = plan.config;
  base.plan = plan;
  base.categories = getFinanceExpenseCategories_();
  base.people = getFinanceReceivers_();
  base.staff = getNhanSuTaiChinhList_(session.maKyHoc);
  return jsonResponse_(base);
}

function getFamilyTransactions_(maKyHoc, month) {
  const ledgerItems = readObjectsNoCache_(SHEET_SOTHUCHI)
    .filter(row => String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim())
    .map(mapThuChiTransaction_)
    .filter(item => item.maGiaoDich && item.ngayDate && item.phamVi === 'GIA_DINH' && String(item.ngayGiaoDich || '').slice(0, 7) === month)
    .map(item => ({
      maGiaoDichGiaDinh: item.maGiaoDich,
      maGiaoDich: item.maGiaoDich,
      ledgerType: item.loai,
      maNhomChuyen: item.maNhomChuyen,
      ngayGiaoDich: item.ngayGiaoDich,
      loai: item.loaiGiaDinh || item.loai,
      maDanhMucGiaDinh: item.maDanhMucGiaDinh || item.maDanhMuc,
      tenDanhMuc: item.tenDanhMuc,
      noiDung: item.noiDung,
      soTien: item.soTien,
      maNguonTien: item.maNguonTien,
      tenNguonTien: item.tenNguonTien,
      nguoiNopNhan: item.nguoiNopNhan,
      soPhieu: item.soPhieu,
      soChungTu: item.soChungTu,
      ghiChu: item.ghiChu,
      trangThai: item.trangThai,
      nguonDuLieu: item.nguonDuLieu,
      maGiaoDichLienKet: item.maGiaoDichLienKet
    }));
  const ledgerReferences = ledgerItems.reduce((map, item) => { map[item.maGiaoDichGiaDinh] = true; return map; }, {});
  const legacyItems = readObjectsNoCache_(SHEET_GIAODICH_GIADINH).filter(row =>
    String(row.Thang || '').trim() === month &&
    String(row.TrangThai || 'ACTIVE').trim().toUpperCase() === 'ACTIVE' &&
    !String(row.MaGiaoDichSoThuChi || '').trim() &&
    !ledgerReferences[String(row.MaGiaoDichGiaDinh || '').trim()]
  ).map(row => ({
    maGiaoDichGiaDinh: String(row.MaGiaoDichGiaDinh || '').trim(),
    maGiaoDich: '',
    ledgerType: String(row.Loai || '').trim().toUpperCase() === 'THU' ? 'THU' : 'CHI',
    maNhomChuyen: '',
    ngayGiaoDich: formatDateForInput_(toDateOnly_(row.NgayGiaoDich)),
    loai: String(row.Loai || '').trim().toUpperCase(),
    maDanhMucGiaDinh: String(row.MaDanhMucGiaDinh || '').trim(),
    tenDanhMuc: String(row.TenDanhMuc || '').trim(),
    noiDung: String(row.NoiDung || '').trim(),
    soTien: number_(row.SoTien),
    maNguonTien: '',
    tenNguonTien: String(row.NguonTien || '').trim() || 'Chưa phân loại (dữ liệu cũ)',
    nguoiNopNhan: '',
    soPhieu: '',
    soChungTu: '',
    ghiChu: String(row.GhiChu || '').trim(),
    trangThai: 'HOAT_DONG',
    nguonDuLieu: 'GIA_DINH_CU',
    maGiaoDichLienKet: ''
  }));
  return ledgerItems.concat(legacyItems).sort((a, b) => String(b.ngayGiaoDich).localeCompare(String(a.ngayGiaoDich)) || String(b.soPhieu || '').localeCompare(String(a.soPhieu || '')));
}

function buildFamilySourceSummaries_(maKyHoc, month) {
  const ym = parseYearMonth_(month);
  const monthStart = parseInputDate_(month + '-01');
  const monthEnd = new Date(ym.year, ym.month, 0, 23, 59, 59, 999);
  const transactions = readObjectsNoCache_(SHEET_SOTHUCHI)
    .filter(row => String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim())
    .map(mapThuChiTransaction_)
    .filter(item => item.maGiaoDich && item.ngayDate && item.trangThai === 'HOAT_DONG' && item.phamVi === 'GIA_DINH' && item.ngayDate.getTime() <= monthEnd.getTime());
  return getNguonTienList_(maKyHoc, 'GIA_DINH').map(source => {
    const sourceTransactions = transactions.filter(item => item.maNguonTien === source.maNguonTien);
    const openingBalance = number_(source.soDuBanDau) + sourceTransactions.filter(item => item.ngayDate.getTime() < monthStart.getTime()).reduce((sum, item) => sum + (item.loai === 'THU' ? item.soTien : -item.soTien), 0);
    const inMonth = sourceTransactions.filter(item => item.ngayDate.getTime() >= monthStart.getTime());
    const transferIn = inMonth.reduce((sum, item) => sum + (item.nguonDuLieu === 'GIA_DINH_CHUYEN' && item.loai === 'THU' ? item.soTien : 0), 0);
    const transferOut = inMonth.reduce((sum, item) => sum + (item.nguonDuLieu === 'GIA_DINH_CHUYEN' && item.loai === 'CHI' ? item.soTien : 0), 0);
    const income = inMonth.reduce((sum, item) => sum + (item.nguonDuLieu !== 'GIA_DINH_CHUYEN' && item.loai === 'THU' ? item.soTien : 0), 0);
    const expense = inMonth.reduce((sum, item) => sum + (item.nguonDuLieu !== 'GIA_DINH_CHUYEN' && item.loai === 'CHI' ? item.soTien : 0), 0);
    return Object.assign({}, source, {
      openingBalance: openingBalance,
      totalThu: income,
      totalChi: expense,
      chuyenVao: transferIn,
      chuyenRa: transferOut,
      closingBalance: openingBalance + income - expense + transferIn - transferOut
    });
  });
}

function getTaiChinhGiaDinhData(token, yearMonth) {
  const session = requireSession_(token, 'finance.read');
  ensureThuChiSheets_(session.maKyHoc);
  ensureQuanLyTaiChinhSheets_();
  const context = getFinanceMonthContext_(session, yearMonth);
  const jars = buildHuTaiChinhData_(session.maKyHoc, context.month, context.fee, context.transactions, { items: [] });
  const ownerJar = (jars.items || []).find(item => item.systemRole === 'OWNER_COMPENSATION') || null;
  const ownerWithdrawals = ownerJar ? context.monthTransactions.filter(item =>
    item.loai === 'CHI' && item.maDanhMuc === 'CHI_GIA_DINH' && item.maHuTaiChinh === ownerJar.code
  ).map(item => ({
    maGiaoDich: item.maGiaoDich, ngayGiaoDich: item.ngayGiaoDich, ngayDisplay: item.ngayDisplay,
    soTien: item.soTien, maNguonTien: item.maNguonTien, tenNguonTien: item.tenNguonTien,
    nguoiNhan: item.nguoiNopNhan, soPhieu: item.soPhieu, noiDung: item.noiDung, ghiChu: item.ghiChu
  })).sort((a, b) => String(b.ngayGiaoDich).localeCompare(String(a.ngayGiaoDich))) : [];
  const linkedOwnerDraw = ownerWithdrawals.reduce((sum, item) => sum + number_(item.soTien), 0);
  const totalOwnerDraw = context.monthTransactions.reduce((sum, item) => sum + (item.loai === 'CHI' && item.maDanhMuc === 'CHI_GIA_DINH' ? item.soTien : 0), 0);
  const unlinkedOwnerDraw = Math.max(totalOwnerDraw - linkedOwnerDraw, 0);
  const familyCategories = getDanhMucGiaDinhList_();
  const familyRows = getFamilyTransactions_(session.maKyHoc, context.month);
  const activeFamilyRows = familyRows.filter(item => item.trangThai === 'HOAT_DONG');
  const familySourceSummaries = buildFamilySourceSummaries_(session.maKyHoc, context.month);
  const configRow = readObjectsNoCache_(SHEET_CAUHINH_GIADINH_THANG).find(row => String(row.Thang || '').trim() === context.month) || {};
  const manualIncome = activeFamilyRows.reduce((sum, item) => sum + (item.loai === 'THU' && item.nguonDuLieu !== 'GIA_DINH_RUT_CHU' ? item.soTien : 0), 0);
  const familyExpense = activeFamilyRows.reduce((sum, item) => sum + (item.loai === 'CHI' ? item.soTien : 0), 0);
  const familySaving = Math.max(0, activeFamilyRows.reduce((sum, item) => {
    if (item.loai === 'TIET_KIEM') return sum + item.soTien;
    if (item.nguonDuLieu === 'GIA_DINH_CHUYEN' && item.maNguonTien === 'GD_TIET_KIEM') return sum + (item.ledgerType === 'THU' ? item.soTien : -item.soTien);
    return sum;
  }, 0));
  const familySource = FinanceDomain.calculateFamilySource({
    plannedOwnerIncome: ownerJar ? number_(ownerJar.allocated) : 0, ownerReceived: linkedOwnerDraw,
    otherIncome: manualIncome, plannedOtherIncome: Math.max(0, number_(configRow.ThuNhapKhacDuKien)),
    totalExpense: familyExpense, totalSaving: familySaving
  });
  const actualMap = activeFamilyRows.reduce((map, item) => { map[item.maDanhMucGiaDinh] = (map[item.maDanhMucGiaDinh] || 0) + item.soTien; return map; }, {});
  const warnings = [];
  const activeBudgetCategories = familyCategories.filter(item => item.trangThai === 'ACTIVE' && item.loai !== 'THU');
  if (!ownerJar) warnings.push({ level: 'danger', message: 'Chưa xác định Hũ lương chủ trung tâm nên chưa thể lập nguồn thu gia đình tự động.' });
  if (jars.allocationStatus && !jars.allocationStatus.locked) warnings.push({ level: 'warning', message: 'Phương án phân bổ hũ chưa chốt; nguồn thu gia đình theo kế hoạch vẫn có thể thay đổi.' });
  if (unlinkedOwnerDraw > 0) warnings.push({ level: 'warning', message: 'Có ' + formatMoneyText_(unlinkedOwnerDraw) + ' khoản chi gia đình chưa liên kết đúng Hũ lương chủ trung tâm.' });
  const ratioTotal = activeBudgetCategories.reduce((sum, item) => sum + item.tyLeMacDinh, 0);
  if (Math.abs(ratioTotal - 100) > 0.01) warnings.push({ level: 'warning', message: 'Tổng tỷ lệ phân bổ gia đình hiện là ' + ratioTotal.toFixed(1) + '%, cần điều chỉnh về 100%.' });
  const budget = activeBudgetCategories.map(item => {
    const amount = familySource.budgetIncome * item.tyLeMacDinh / 100;
    const actual = number_(actualMap[item.maDanhMucGiaDinh]);
    if (item.kieuGioiHan === 'MAX' && actual > amount && amount > 0) warnings.push({ level: 'danger', message: item.tenDanhMuc + ' đã vượt ngân sách ' + formatMoneyText_(actual - amount) + '.' });
    if (item.kieuGioiHan === 'MIN' && actual < amount) warnings.push({ level: 'warning', message: item.tenDanhMuc + ' còn thiếu ' + formatMoneyText_(amount - actual) + ' so với mục tiêu.' });
    return Object.assign({}, item, { nganSach: amount, thucTe: actual, conLai: item.kieuGioiHan === 'MIN' ? Math.max(amount - actual, 0) : amount - actual });
  });
  if (familyExpense + familySaving > familySource.totalIncome && familySource.totalIncome > 0) warnings.push({ level: 'danger', message: 'Tổng chi và tiết kiệm gia đình đang lớn hơn thu nhập thực tế trong tháng.' });
  return jsonResponse_({
    month: context.month,
    periodLock: GovernanceService.periodStatus(session, context.month),
    family: {
      categories: familyCategories, transactions: familyRows, withdrawals: ownerWithdrawals, budget: budget, warnings: warnings,
      sources: getNguonTienList_(session.maKyHoc, 'TRUNG_TAM').filter(item => item.trangThai === 'ACTIVE'),
      familySources: getNguonTienList_(session.maKyHoc, 'GIA_DINH').filter(item => item.trangThai === 'ACTIVE'),
      sourceSummaries: familySourceSummaries,
      ownerJar: ownerJar,
      config: { thang: context.month, thuNhapDuKien: number_(configRow.ThuNhapKhacDuKien), quyKhanCapMucTieu: number_(configRow.QuyKhanCapMucTieu), ghiChu: String(configRow.GhiChu || '').trim() },
      summary: Object.assign({}, familySource, { unlinkedOwnerDraw: unlinkedOwnerDraw }),
      allocationStatus: jars.allocationStatus,
      people: getFinanceReceivers_()
    }
  });
}

function getQuanLyTaiChinhData(token, yearMonth) {
  const session = requireSession_(token, 'finance.read'); ensureThuChiSheets_(session.maKyHoc); ensureQuanLyTaiChinhSheets_();
  const ym = parseYearMonth_(yearMonth); const month = ym.year + '-' + String(ym.month).padStart(2, '0');
  const monthStart = parseInputDate_(month + '-01'); const monthEnd = new Date(ym.year, ym.month, 0, 23, 59, 59, 999);
  const config = getCauHinhTaiChinhThang_(session.maKyHoc, month);
  const feeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getThuPhiMonthSheetName_(ym.year, ym.month));
  const feeRows = feeSheet ? readObjectsNoCache_(feeSheet.getName()).filter(row => String(row.MaKyHoc || '').trim() === session.maKyHoc && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED') : [];
  const feeSummary = { totalStudents: 0, payingStudents: 0, expected: 0, collected: 0, remaining: 0, classes: [] };
  const classMap = {};
  feeRows.forEach(row => {
    const fee = toBoolean_(row.TamNghi) ? 0 : number_(row.HocPhi); const collected = toBoolean_(row.TamNghi) ? 0 : number_(row.SoTienDaThu);
    if (fee <= 0) return; feeSummary.totalStudents++; feeSummary.payingStudents++; feeSummary.expected += fee; feeSummary.collected += collected; feeSummary.remaining += Math.max(fee - collected, 0);
    const key = String(row.Lop || row.TenLop || 'Chưa xếp lớp').trim();
    if (!classMap[key]) classMap[key] = { lop: key, tenLop: String(row.TenLop || key).trim(), students: 0, expected: 0, collected: 0 };
    classMap[key].students++; classMap[key].expected += fee; classMap[key].collected += collected;
  });
  feeSummary.classes = Object.keys(classMap).map(key => classMap[key]).sort((a, b) => String(a.lop).localeCompare(String(b.lop), 'vi', { numeric: true }));

  const allTransactions = readObjectsNoCache_(SHEET_SOTHUCHI).filter(row => String(row.MaKyHoc || '').trim() === session.maKyHoc).map(mapThuChiTransaction_).filter(item => item.maGiaoDich && item.ngayDate && item.trangThai === 'HOAT_DONG' && item.phamVi === 'TRUNG_TAM');
  const monthTransactions = allTransactions.filter(item => item.ngayDate.getTime() >= monthStart.getTime() && item.ngayDate.getTime() <= monthEnd.getTime() && item.nguonDuLieu !== 'CHUYEN_NOI_BO');
  const cashIncome = monthTransactions.reduce((sum, item) => sum + (item.loai === 'THU' ? item.soTien : 0), 0);
  const businessCashExpense = monthTransactions.reduce((sum, item) => sum + (item.loai === 'CHI' && item.maDanhMuc !== 'CHI_GIA_DINH' ? item.soTien : 0), 0);
  const ownerDraw = monthTransactions.reduce((sum, item) => sum + (item.loai === 'CHI' && item.maDanhMuc === 'CHI_GIA_DINH' ? item.soTien : 0), 0);
  const otherIncome = monthTransactions.reduce((sum, item) => sum + (item.loai === 'THU' && item.nguonDuLieu !== 'HOC_PHI' ? item.soTien : 0), 0);
  const currentCash = getNguonTienList_(session.maKyHoc).reduce((sum, source) => sum + number_(source.soDuBanDau), 0) + allTransactions.filter(item => item.ngayDate.getTime() <= monthEnd.getTime()).reduce((sum, item) => sum + (item.loai === 'THU' ? item.soTien : -item.soTien), 0);
  const planData = buildKeHoachChiThangData_(session.maKyHoc, month, allTransactions, currentCash, feeSummary.expected || cashIncome);
  const businessPlans = planData.items.filter(item => item.nhomChi !== 'GIA_DINH');
  const businessPlanTotal = businessPlans.reduce((sum, item) => sum + item.soTienPhaiChi, 0);
  const businessPlanRemaining = businessPlans.reduce((sum, item) => sum + item.conPhaiChi, 0);
  const revenueForecast = number_(config.doanhThuDuKien) || feeSummary.expected || cashIncome;
  const accruedRevenueEstimate = feeSummary.expected + otherIncome;
  const studentCount = number_(config.soHocSinh) || feeSummary.payingStudents;
  const currentAverageFee = feeSummary.payingStudents > 0 ? feeSummary.expected / feeSummary.payingStudents : (studentCount > 0 ? revenueForecast / studentCount : 0);
  const targetMargin = Math.max(0, Math.min(80, number_(config.tyLeLoiNhuanMucTieu) || 20));
  const variableTotal = businessPlans.filter(item => item.nhomChi === 'BAN_TRU' || item.nhomChi === 'HOC_CU').reduce((sum, item) => sum + item.soTienPhaiChi, 0);
  const pricing = FinanceDomain.calculatePricing({
    studentCount: studentCount,
    currentAverageFee: currentAverageFee,
    plannedExpense: businessPlanTotal,
    variableExpense: variableTotal,
    targetMargin: targetMargin,
    collectionRate: 100
  });
  const performance = FinanceDomain.calculatePerformance({
    cashIncome: cashIncome,
    cashExpense: businessCashExpense,
    currentCash: currentCash,
    revenueForecast: revenueForecast,
    accruedRevenue: accruedRevenueEstimate,
    plannedExpense: businessPlanTotal,
    remainingObligations: businessPlanRemaining,
    reserveTarget: number_(planData.summary.reserveTarget),
    ownerDraw: ownerDraw
  });
  const jars = buildHuTaiChinhData_(session.maKyHoc, month, feeSummary, allTransactions, planData);
  jars.canManageAllocation = session.vaiTro === 'OWNER';
  const ownerJar = (jars.items || []).find(item => item.systemRole === 'OWNER_COMPENSATION') || null;
  const ownerWithdrawals = ownerJar ? monthTransactions.filter(item =>
    item.loai === 'CHI' && item.maDanhMuc === 'CHI_GIA_DINH' && item.maHuTaiChinh === ownerJar.code
  ).map(item => ({
    maGiaoDich: item.maGiaoDich, ngayGiaoDich: item.ngayGiaoDich, ngayDisplay: item.ngayDisplay,
    soTien: item.soTien, maNguonTien: item.maNguonTien, tenNguonTien: item.tenNguonTien,
    nguoiNhan: item.nguoiNopNhan, soPhieu: item.soPhieu, noiDung: item.noiDung, ghiChu: item.ghiChu
  })).sort((a, b) => String(b.ngayGiaoDich).localeCompare(String(a.ngayGiaoDich))) : [];
  const linkedOwnerDraw = ownerWithdrawals.reduce((sum, item) => sum + number_(item.soTien), 0);
  const unlinkedOwnerDraw = Math.max(ownerDraw - linkedOwnerDraw, 0);
  const plannedOwnerIncome = ownerJar ? number_(ownerJar.allocated) : 0;

  const familyCategories = getDanhMucGiaDinhList_();
  const familyRows = getFamilyTransactions_(session.maKyHoc, month);
  const activeFamilyRows = familyRows.filter(item => item.trangThai === 'HOAT_DONG');
  const familySourceSummaries = buildFamilySourceSummaries_(session.maKyHoc, month);
  const familyConfigRow = readObjectsNoCache_(SHEET_CAUHINH_GIADINH_THANG).find(row => String(row.Thang || '').trim() === month) || {};
  const manualFamilyIncome = activeFamilyRows.reduce((sum, item) => sum + (item.loai === 'THU' && item.nguonDuLieu !== 'GIA_DINH_RUT_CHU' ? item.soTien : 0), 0);
  const familyExpense = activeFamilyRows.reduce((sum, item) => sum + (item.loai === 'CHI' ? item.soTien : 0), 0);
  const familySaving = Math.max(0, activeFamilyRows.reduce((sum, item) => {
    if (item.loai === 'TIET_KIEM') return sum + item.soTien;
    if (item.nguonDuLieu === 'GIA_DINH_CHUYEN' && item.maNguonTien === 'GD_TIET_KIEM') return sum + (item.ledgerType === 'THU' ? item.soTien : -item.soTien);
    return sum;
  }, 0));
  const plannedOtherIncome = Math.max(0, number_(familyConfigRow.ThuNhapKhacDuKien));
  const familySource = FinanceDomain.calculateFamilySource({
    plannedOwnerIncome: plannedOwnerIncome, ownerReceived: linkedOwnerDraw,
    otherIncome: manualFamilyIncome, plannedOtherIncome: plannedOtherIncome,
    totalExpense: familyExpense, totalSaving: familySaving
  });
  const familyIncome = familySource.totalIncome;
  const budgetIncome = familySource.budgetIncome;
  const familyActualMap = activeFamilyRows.reduce((map, item) => { map[item.maDanhMucGiaDinh] = (map[item.maDanhMucGiaDinh] || 0) + item.soTien; return map; }, {});
  const familyWarnings = []; const activeBudgetCategories = familyCategories.filter(item => item.trangThai === 'ACTIVE' && item.loai !== 'THU');
  if (!ownerJar) familyWarnings.push({ level: 'danger', message: 'Chưa xác định Hũ lương chủ trung tâm nên chưa thể lập nguồn thu gia đình tự động.' });
  if (jars.allocationStatus && !jars.allocationStatus.locked) familyWarnings.push({ level: 'warning', message: 'Phương án phân bổ hũ chưa chốt; nguồn thu gia đình theo kế hoạch vẫn có thể thay đổi.' });
  if (unlinkedOwnerDraw > 0) familyWarnings.push({ level: 'warning', message: 'Có ' + formatMoneyText_(unlinkedOwnerDraw) + ' khoản chi gia đình chưa liên kết đúng Hũ lương chủ trung tâm nên không được tính vào tiền gia đình thực nhận.' });
  const ratioTotal = activeBudgetCategories.reduce((sum, item) => sum + item.tyLeMacDinh, 0);
  if (Math.abs(ratioTotal - 100) > 0.01) familyWarnings.push({ level: 'warning', message: 'Tổng tỷ lệ phân bổ gia đình hiện là ' + ratioTotal.toFixed(1) + '%, cần điều chỉnh về 100%.' });
  const familyBudget = activeBudgetCategories.map(item => {
    const budget = budgetIncome * item.tyLeMacDinh / 100; const actual = number_(familyActualMap[item.maDanhMucGiaDinh]);
    if (item.kieuGioiHan === 'MAX' && actual > budget && budget > 0) familyWarnings.push({ level: 'danger', message: item.tenDanhMuc + ' đã vượt ngân sách ' + formatMoneyText_(actual - budget) + '.' });
    if (item.kieuGioiHan === 'MIN' && actual < budget) familyWarnings.push({ level: 'warning', message: item.tenDanhMuc + ' còn thiếu ' + formatMoneyText_(budget - actual) + ' so với mục tiêu.' });
    return Object.assign({}, item, { nganSach: budget, thucTe: actual, conLai: item.kieuGioiHan === 'MIN' ? Math.max(budget - actual, 0) : budget - actual });
  });
  if (familyExpense + familySaving > familyIncome && familyIncome > 0) familyWarnings.push({ level: 'danger', message: 'Tổng chi và tiết kiệm gia đình đang lớn hơn thu nhập thực tế trong tháng.' });

  const receivers = readObjectsNoCache_(SHEET_DOITUONG_THUCHI).filter(row => {
    const type = String(row.Loai || 'BOTH').trim().toUpperCase();
    const status = String(row.TrangThai || 'ACTIVE').trim().toUpperCase();
    return (type === 'CHI' || type === 'BOTH') && status === 'ACTIVE';
  }).map(row => ({
    maDoiTuong: String(row.MaDoiTuong || '').trim(),
    loai: String(row.Loai || 'BOTH').trim().toUpperCase(),
    tenDoiTuong: String(row.TenDoiTuong || '').trim(),
    soDienThoai: String(row.SoDienThoai || '').trim(),
    diaChi: String(row.DiaChi || '').trim(),
    ghiChu: String(row.GhiChu || '').trim()
  })).filter(item => item.maDoiTuong && item.tenDoiTuong)
    .sort((a, b) => a.tenDoiTuong.localeCompare(b.tenDoiTuong, 'vi'));

  return jsonResponse_({
    month: month, config: config, staff: getNhanSuTaiChinhList_(session.maKyHoc), recurringExpenses: getKhoanChiDinhKyList_(session.maKyHoc, month),
    people: receivers,
    categories: readObjectsNoCache_(SHEET_DANHMUC_THUCHI).filter(row => String(row.Loai || '').trim().toUpperCase() === 'CHI' && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() === 'ACTIVE').map(row => ({ maDanhMuc: String(row.MaDanhMuc || '').trim(), tenDanhMuc: String(row.TenDanhMuc || '').trim() })),
    fee: feeSummary, plan: planData, jars: jars,
    periodLock: GovernanceService.periodStatus(session, month),
    performance: performance,
    pricing: pricing,
    family: { categories: familyCategories, transactions: familyRows, withdrawals: ownerWithdrawals, budget: familyBudget, warnings: familyWarnings,
      sources: getNguonTienList_(session.maKyHoc, 'TRUNG_TAM').filter(item => item.trangThai === 'ACTIVE'),
      familySources: getNguonTienList_(session.maKyHoc, 'GIA_DINH').filter(item => item.trangThai === 'ACTIVE'),
      sourceSummaries: familySourceSummaries,
      ownerJar: ownerJar,
      config: { thang: month, thuNhapDuKien: number_(familyConfigRow.ThuNhapKhacDuKien), quyKhanCapMucTieu: number_(familyConfigRow.QuyKhanCapMucTieu), ghiChu: String(familyConfigRow.GhiChu || '').trim() },
      summary: Object.assign({}, familySource, { unlinkedOwnerDraw: unlinkedOwnerDraw }) }
  });
}

function getQuanLyThuChiData(token, filters) {
  const session = requireSession_(token, 'finance.read');
  ensureThuChiSheets_(session.maKyHoc);

  filters = filters || {};

  const range = resolveThuChiDateRange_(filters.fromDate, filters.toDate);
  const loaiFilter = String(filters.loai || 'ALL').trim().toUpperCase();
  const danhMucFilter = String(filters.maDanhMuc || '').trim();
  const sourceFilter = String(filters.maNguonTien || 'ALL').trim().toUpperCase();
  const scopeFilter = normalizeFinanceScope_(filters.phamVi || 'ALL', 'ALL');
  const trangThaiFilter = String(filters.trangThai || 'HOAT_DONG').trim().toUpperCase();
  const keyword = normalizeText_(filters.keyword || '');
  const planMonth = range.fromText.slice(0, 7);

  const cacheKey = buildCacheKey_(
    'thuchi_v6_' + session.maKyHoc + '_' +
    range.fromText + '_' + range.toText + '_' +
    loaiFilter + '_' + danhMucFilter + '_' +
    sourceFilter + '_' + scopeFilter + '_' + trangThaiFilter + '_' + keyword
  );

  const cached = cacheGetString_(cacheKey);
  if (cached) return cached;

  const categoryRows = readObjects_(SHEET_DANHMUC_THUCHI);
  const transactionRows = readObjects_(SHEET_SOTHUCHI);
  const peopleRows = readObjects_(SHEET_DOITUONG_THUCHI);
  const sources = getNguonTienList_(session.maKyHoc, scopeFilter);

  const people = peopleRows
    .filter(row => String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED')
    .map(row => ({
      maDoiTuong: String(row.MaDoiTuong || '').trim(),
      loai: String(row.Loai || 'BOTH').trim().toUpperCase(),
      tenDoiTuong: String(row.TenDoiTuong || '').trim(),
      soDienThoai: String(row.SoDienThoai || '').trim(),
      diaChi: String(row.DiaChi || '').trim(),
      ghiChu: String(row.GhiChu || '').trim()
    }))
    .filter(item => item.maDoiTuong && item.tenDoiTuong)
    .sort((a, b) => a.tenDoiTuong.localeCompare(b.tenDoiTuong, 'vi'));

  const centerCategories = categoryRows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED')
    .map(row => ({
      maDanhMuc: String(row.MaDanhMuc || '').trim(),
      loai: String(row.Loai || '').trim().toUpperCase(),
      tenDanhMuc: String(row.TenDanhMuc || '').trim(),
      maHuMacDinh: normalizeMaHuTaiChinh_(row.MaHuMacDinh),
      thuTu: number_(row.ThuTu) || 999,
      trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase()
    }))
    .filter(item => item.maDanhMuc && (item.loai === 'THU' || item.loai === 'CHI'))
    .sort((a, b) => {
      if (a.loai !== b.loai) return a.loai.localeCompare(b.loai);
      if (a.thuTu !== b.thuTu) return a.thuTu - b.thuTu;
      return a.tenDanhMuc.localeCompare(b.tenDanhMuc, 'vi');
    });

  const familyCategories = getDanhMucGiaDinhList_()
    .filter(item => item.trangThai !== 'DELETED')
    .map(item => ({
      maDanhMuc: item.maDanhMucGiaDinh,
      loai: item.loai === 'THU' ? 'THU' : 'CHI',
      loaiGiaDinh: item.loai,
      tenDanhMuc: item.tenDanhMuc,
      maHuMacDinh: '',
      thuTu: 1000 + item.thuTu,
      trangThai: item.trangThai,
      phamVi: 'GIA_DINH'
    }));
  const categories = centerCategories
    .map(item => Object.assign({ phamVi: 'TRUNG_TAM' }, item))
    .concat(familyCategories);

  const allSessionTransactions = transactionRows
    .filter(row => String(row.MaKyHoc || '').trim() === session.maKyHoc)
    .map(mapThuChiTransaction_)
    .filter(item => item.maGiaoDich && item.ngayDate);

  const scopeTransactions = allSessionTransactions.filter(item => scopeFilter === 'ALL' || item.phamVi === scopeFilter);
  const activeAll = scopeTransactions.filter(item => item.trangThai === 'HOAT_DONG');
  const sourceMap = sources.reduce((map, item) => {
    map[item.maNguonTien] = item;
    return map;
  }, {});

  function signedAmount(item) {
    return item.loai === 'THU' ? item.soTien : -item.soTien;
  }

  function summarizeSource(source) {
    const sourceTransactions = activeAll.filter(item => item.maNguonTien === source.maNguonTien);
    const opening = number_(source.soDuBanDau) + sourceTransactions.reduce((sum, item) => {
      return item.ngayDate.getTime() < range.fromDate.getTime()
        ? sum + signedAmount(item)
        : sum;
    }, 0);

    const inPeriod = sourceTransactions.filter(item => {
      const time = item.ngayDate.getTime();
      return time >= range.fromDate.getTime() && time <= range.toDate.getTime();
    });

    const operational = inPeriod.filter(item => item.nguonDuLieu !== 'CHUYEN_NOI_BO' && item.nguonDuLieu !== 'GIA_DINH_CHUYEN');
    const transfers = inPeriod.filter(item => item.nguonDuLieu === 'CHUYEN_NOI_BO' || item.nguonDuLieu === 'GIA_DINH_CHUYEN');

    const totalThu = operational.reduce((sum, item) => {
      return sum + (item.loai === 'THU' ? item.soTien : 0);
    }, 0);

    const totalChi = operational.reduce((sum, item) => {
      return sum + (item.loai === 'CHI' ? item.soTien : 0);
    }, 0);

    const chuyenVao = transfers.reduce((sum, item) => {
      return sum + (item.loai === 'THU' ? item.soTien : 0);
    }, 0);

    const chuyenRa = transfers.reduce((sum, item) => {
      return sum + (item.loai === 'CHI' ? item.soTien : 0);
    }, 0);

    return {
      maNguonTien: source.maNguonTien,
      tenNguonTien: source.tenNguonTien,
      thuTu: source.thuTu,
      soDuBanDau: number_(source.soDuBanDau),
      openingBalance: opening,
      totalThu: totalThu,
      totalChi: totalChi,
      chuyenVao: chuyenVao,
      chuyenRa: chuyenRa,
      closingBalance: opening + totalThu - totalChi + chuyenVao - chuyenRa,
      transactionCount: inPeriod.length
    };
  }

  const sourceSummaries = sources.map(summarizeSource);

  const unassignedSource = {
    maNguonTien: '',
    tenNguonTien: 'Chưa phân loại',
    thuTu: 999,
    soDuBanDau: 0
  };

  const unassignedTransactions = activeAll.filter(item => !item.maNguonTien);
  const hasUnassigned = unassignedTransactions.length > 0;
  const unassignedSummary = summarizeSource(unassignedSource);

  let dateSourceTransactions = scopeTransactions.filter(item => {
    const time = item.ngayDate.getTime();
    if (time < range.fromDate.getTime() || time > range.toDate.getTime()) return false;

    if (sourceFilter === 'UNASSIGNED') return !item.maNguonTien;
    if (sourceFilter !== 'ALL') return item.maNguonTien === sourceFilter;

    return true;
  });

  let periodTransactions = dateSourceTransactions.slice();

  if (loaiFilter === 'THU' || loaiFilter === 'CHI') {
    periodTransactions = periodTransactions.filter(item => item.loai === loaiFilter);
  }

  if (danhMucFilter) {
    periodTransactions = periodTransactions.filter(item => item.maDanhMuc === danhMucFilter);
  }

  if (trangThaiFilter !== 'ALL') {
    periodTransactions = periodTransactions.filter(item => item.trangThai === trangThaiFilter);
  }

  if (keyword) {
    periodTransactions = periodTransactions.filter(item => {
      const text = normalizeText_(
        item.noiDung + ' ' +
        item.tenDanhMuc + ' ' +
        item.tenNguonTien + ' ' +
        item.nguoiNopNhan + ' ' +
        item.soChungTu + ' ' +
        item.ghiChu
      );

      return text.indexOf(keyword) !== -1;
    });
  }

  const activeFiltered = periodTransactions.filter(item => item.trangThai === 'HOAT_DONG');
  const activeOperational = activeFiltered.filter(item => item.nguonDuLieu !== 'CHUYEN_NOI_BO' && item.nguonDuLieu !== 'GIA_DINH_CHUYEN');

  const totalThu = activeOperational.reduce((sum, item) => {
    return sum + (item.loai === 'THU' ? item.soTien : 0);
  }, 0);

  const totalChi = activeOperational.reduce((sum, item) => {
    return sum + (item.loai === 'CHI' ? item.soTien : 0);
  }, 0);

  const totalHocPhi = activeOperational.reduce((sum, item) => {
    return sum + (
      item.nguonDuLieu === 'HOC_PHI' && item.loai === 'THU'
        ? item.soTien
        : 0
    );
  }, 0);

  let openingBalance = 0;
  let closingBalance = 0;

  if (sourceFilter === 'UNASSIGNED') {
    openingBalance = unassignedSummary.openingBalance;
    closingBalance = unassignedSummary.closingBalance;
  } else if (sourceFilter !== 'ALL' && sourceMap[sourceFilter]) {
    const selected = sourceSummaries.find(item => item.maNguonTien === sourceFilter);
    openingBalance = selected ? selected.openingBalance : 0;
    closingBalance = selected ? selected.closingBalance : 0;
  } else {
    openingBalance = sourceSummaries.reduce((sum, item) => sum + item.openingBalance, 0) +
      unassignedSummary.openingBalance;

    closingBalance = sourceSummaries.reduce((sum, item) => sum + item.closingBalance, 0) +
      unassignedSummary.closingBalance;
  }

  const centerSources = getNguonTienList_(session.maKyHoc, 'TRUNG_TAM');
  const centerActiveAll = allSessionTransactions.filter(item => item.trangThai === 'HOAT_DONG' && item.phamVi === 'TRUNG_TAM');
  const allClosingBalance = centerSources.reduce((sum, source) => {
    return sum + number_(source.soDuBanDau) + centerActiveAll
      .filter(item => item.maNguonTien === source.maNguonTien)
      .reduce((sourceSum, item) => sourceSum + signedAmount(item), 0);
  }, 0);
  const monthOperational = centerActiveAll.filter(item => {
    return item.nguonDuLieu !== 'CHUYEN_NOI_BO' &&
      String(item.ngayGiaoDich || '').slice(0, 7) === planMonth;
  });
  const monthRevenue = monthOperational.reduce((sum, item) => {
    return sum + (item.loai === 'THU' ? item.soTien : 0);
  }, 0);
  const monthlyFinance = buildKeHoachChiThangData_(
    session.maKyHoc,
    planMonth,
    centerActiveAll,
    allClosingBalance,
    monthRevenue
  );

  const sortedDesc = periodTransactions.slice().sort((a, b) => {
    const dateCompare = b.ngayDate.getTime() - a.ngayDate.getTime();
    if (dateCompare !== 0) return dateCompare;
    return String(b.updatedAtRaw || '').localeCompare(String(a.updatedAtRaw || ''));
  });

  const cashbookActive = dateSourceTransactions
    .filter(item => item.trangThai === 'HOAT_DONG')
    .sort((a, b) => {
      const dateCompare = a.ngayDate.getTime() - b.ngayDate.getTime();
      if (dateCompare !== 0) return dateCompare;
      return String(a.createdAtRaw || '').localeCompare(String(b.createdAtRaw || ''));
    });

  let runningBalance = openingBalance;
  const cashbook = cashbookActive.map(item => {
    runningBalance += signedAmount(item);

    return Object.assign({}, stripThuChiPrivateFields_(item), {
      soDu: runningBalance
    });
  });

  const reportYear = range.fromDate.getFullYear();
  const yearTransactions = scopeTransactions.filter(item => {
    if (!item.ngayDate || item.ngayDate.getFullYear() !== reportYear) return false;
    if (sourceFilter === 'UNASSIGNED') return !item.maNguonTien;
    if (sourceFilter !== 'ALL') return item.maNguonTien === sourceFilter;
    return true;
  });

  const monthly = buildThuChiMonthlyReport_(
    yearTransactions.filter(item => item.nguonDuLieu !== 'CHUYEN_NOI_BO' && item.nguonDuLieu !== 'GIA_DINH_CHUYEN'),
    reportYear
  );

  const categoryReport = buildThuChiCategoryReport_(activeOperational);
  const paymentMethods = buildThuChiPaymentMethodReport_(activeOperational);

  const data = {
    session: session,
    range: {
      fromDate: range.fromText,
      toDate: range.toText,
      openingBalance: openingBalance,
      closingBalance: closingBalance
    },
    summary: {
      totalThu: totalThu,
      totalChi: totalChi,
      net: totalThu - totalChi,
      openingBalance: openingBalance,
      closingBalance: closingBalance,
      totalHocPhi: totalHocPhi,
      transactionCount: activeFiltered.length,
      unassignedCount: unassignedTransactions.length
    },
    scope: scopeFilter,
    categories: categories,
    financialJars: getDanhMucHuTaiChinhList_(false),
    people: people,
    staff: getNhanSuTaiChinhList_(session.maKyHoc),
    sources: sources,
    centerSources: getNguonTienList_(session.maKyHoc, 'TRUNG_TAM'),
    familySources: getNguonTienList_(session.maKyHoc, 'GIA_DINH'),
    sourceSummaries: sourceSummaries,
    unassignedSummary: hasUnassigned ? unassignedSummary : null,
    transactions: sortedDesc.map(stripThuChiPrivateFields_),
    recentTransactions: sortedDesc.slice(0, 10).map(stripThuChiPrivateFields_),
    cashbook: cashbook,
    monthlyFinance: monthlyFinance,
    reports: {
      year: reportYear,
      monthly: monthly,
      categories: categoryReport,
      paymentMethods: paymentMethods,
      sources: sourceSummaries
    }
  };

  const json = jsonResponse_(data);
  cachePutString_(cacheKey, json, CACHE_SECONDS);
  return json;
}

function saveDoiTuongThuChi(token, data) {
  requireSession_(token, 'finance.write');
  data = data || {};
  const tenDoiTuong = String(data.tenDoiTuong || '').trim();
  const loai = String(data.loai || 'BOTH').trim().toUpperCase();
  const soDienThoai = String(data.soDienThoai || '').trim();
  const diaChi = String(data.diaChi || '').trim();
  const ghiChu = String(data.ghiChu || '').trim();

  if (!tenDoiTuong) throw new Error('Vui lòng nhập tên người nộp/người nhận.');
  if (['THU', 'CHI', 'BOTH'].indexOf(loai) === -1) throw new Error('Loại đối tượng không hợp lệ.');
  if (soDienThoai && !/^0\d{9}$/.test(soDienThoai)) {
    throw new Error('Số điện thoại phải gồm 10 số và bắt đầu bằng số 0.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Hệ thống đang cập nhật danh mục người nộp/người nhận.');
  let saved = null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureSheet_(ss, SHEET_DOITUONG_THUCHI, getDoiTuongThuChiHeaders_());
    const rows = readObjectsNoCache_(SHEET_DOITUONG_THUCHI);
    const duplicate = rows.some(row => {
      return normalizeText_(row.TenDoiTuong) === normalizeText_(tenDoiTuong) &&
        String(row.Loai || 'BOTH').trim().toUpperCase() === loai &&
        String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED';
    });
    if (duplicate) throw new Error('Tên này đã có trong danh sách.');
    const now = new Date();
    saved = {
      maDoiTuong: 'DT_' + Utilities.getUuid().slice(0, 10).toUpperCase(),
      loai: loai,
      tenDoiTuong: tenDoiTuong,
      soDienThoai: soDienThoai,
      diaChi: diaChi,
      ghiChu: ghiChu
    };
    appendObjectsToSheet_(sheet, [{
      MaDoiTuong: saved.maDoiTuong, Loai: loai, TenDoiTuong: tenDoiTuong,
      SoDienThoai: soDienThoai, DiaChi: diaChi, GhiChu: ghiChu,
      TrangThai: 'ACTIVE', CreatedAt: now, UpdatedAt: now
    }], getDoiTuongThuChiHeaders_());
  } finally {
    lock.releaseLock();
  }
  bumpDataVersion_();
  return jsonResponse_({ success: true, item: saved, message: 'Đã thêm người nộp/người nhận.' });
}

function saveThuChiGiaoDich(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);

  const command = CashbookDomain.transaction(Object.assign({}, data || {}, { soTien: number_(data && data.soTien) }));
  const maGiaoDich = command.maGiaoDich;
  const loai = command.loai;
  const maDanhMuc = command.maDanhMuc;
  let maKeHoachChi = command.maKeHoachChi;
  let maNhanSu = command.maNhanSu;
  const requestedJar = command.maHuTaiChinh;
  const maNguonTien = command.maNguonTien;
  const ngayText = command.ngayGiaoDich;
  const noiDung = command.noiDung;
  const soTien = command.soTien;
  let nguoiNopNhan = command.nguoiNopNhan;
  const soChungTu = command.soChungTu;
  const ghiChu = command.ghiChu;
  const chungTuImage = command.chungTuImage;

  assertFinancePeriodOpen_(session, ngayText.slice(0, 7));

  if (!maNguonTien || !getNguonTienDefinition_(maNguonTien)) {
    throw new Error('Vui lòng chọn đúng nguồn tiền.');
  }

  const categories = readObjectsNoCache_(SHEET_DANHMUC_THUCHI);
  const category = categories.find(row => {
    return String(row.MaDanhMuc || '').trim() === maDanhMuc &&
      String(row.Loai || '').trim().toUpperCase() === loai &&
      String(row.TrangThai || '').trim().toUpperCase() === 'ACTIVE';
  });

  if (!category) {
    throw new Error('Danh mục không tồn tại, đã ngừng sử dụng hoặc không đúng loại giao dịch.');
  }
  let selectedStaff = null;
  if (loai === 'CHI' && maDanhMuc === 'CHI_LUONG') {
    selectedStaff = getNhanSuTaiChinhList_(session.maKyHoc).find(item => item.maNhanSu === maNhanSu && (item.trangThai === 'ACTIVE' || !!maGiaoDich)) || null;
    if (!selectedStaff) throw new Error('Vui lòng chọn người nhận từ danh sách nhân sự đang làm việc.');
    nguoiNopNhan = selectedStaff.hoTen;
    if (!maKeHoachChi) {
      const month = ngayText.slice(0, 7);
      const staffPlans = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG).filter(row =>
        String(row.MaKyHoc || '').trim() === session.maKyHoc && String(row.Thang || '').trim() === month &&
        String(row.MaDanhMuc || '').trim() === 'CHI_LUONG' && String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED' &&
        (String(row.MaNhanSu || '').trim() === maNhanSu || String(row.MaThamChieu || '').trim() === 'NHANSU|' + maNhanSu)
      );
      const baseSalaryPlan = staffPlans.find(row => String(row.NguonKeHoach || '').trim().toUpperCase() === 'NHANSU') || (staffPlans.length === 1 ? staffPlans[0] : null);
      if (baseSalaryPlan) maKeHoachChi = String(baseSalaryPlan.MaKeHoachChi || '').trim();
    }
  } else {
    maNhanSu = '';
  }
  const activeJarCodes = getDanhMucHuTaiChinhList_(false).reduce((map, item) => { map[item.code] = true; return map; }, {});
  if (loai === 'CHI' && requestedJar && !activeJarCodes[normalizeMaHuTaiChinh_(requestedJar)]) {
    throw new Error('Hũ tài chính của phiếu chi không hợp lệ.');
  }
  let maHuTaiChinh = loai === 'CHI'
    ? (normalizeMaHuTaiChinh_(requestedJar) || normalizeMaHuTaiChinh_(category.MaHuMacDinh) || inferHuTaiChinhForCategory_(maDanhMuc, category.TenDanhMuc))
    : '';

  if (selectedStaff) {
    const jars = getDanhMucHuTaiChinhList_(false);
    const selectedJar = jars.find(item => item.code === maHuTaiChinh) || null;
    const isOwner = normalizeText_(selectedStaff.vaiTro || '').indexOf('chu co so') !== -1;
    if (!isOwner && selectedJar && selectedJar.systemRole === 'OWNER_COMPENSATION') {
      const staffJar = jars.find(item => item.code === 'VAN_HANH') || jars.find(item => {
        const text = normalizeText_(item.name || '');
        return item.systemRole !== 'OWNER_COMPENSATION' && (text.indexOf('luong trung tam') !== -1 || text.indexOf('luong nhan vien') !== -1 || text.indexOf('van hanh') !== -1);
      });
      if (!staffJar) throw new Error('Hãy cấu hình một hũ chi lương nhân viên hoặc hũ vận hành trước khi lập phiếu chi lương.');
      maHuTaiChinh = staffJar.code;
    }
  }

  if (maKeHoachChi) {
    if (loai !== 'CHI') throw new Error('Chỉ phiếu chi mới được liên kết với kế hoạch chi tháng.');
    const plan = readObjectsNoCache_(SHEET_KEHOACH_CHI_THANG).find(row => {
      return String(row.MaKeHoachChi || '').trim() === maKeHoachChi &&
        String(row.MaKyHoc || '').trim() === session.maKyHoc &&
        String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED';
    });
    if (!plan) throw new Error('Khoản kế hoạch chi không còn tồn tại.');
    if (String(plan.MaDanhMuc || '').trim() !== maDanhMuc) {
      throw new Error('Danh mục phiếu chi phải trùng với danh mục của khoản kế hoạch.');
    }
    const planReference = String(plan.MaThamChieu || '').trim();
    let planStaffId = String(plan.MaNhanSu || '').trim() || (planReference.indexOf('NHANSU|') === 0 ? planReference.slice(8) : '');
    if (!planStaffId && planReference.indexOf('DINHKY|') === 0) {
      const recurringPlan = readObjectsNoCache_(SHEET_KHOANCHI_DINHKY).find(row =>
        String(row.MaKyHoc || '').trim() === session.maKyHoc && String(row.MaKhoanDinhKy || '').trim() === planReference.slice(7)
      );
      planStaffId = recurringPlan ? String(recurringPlan.MaNhanSu || '').trim() : '';
    }
    if (maDanhMuc === 'CHI_LUONG' && planStaffId && planStaffId !== maNhanSu) {
      throw new Error('Người nhận phiếu chi không khớp với nhân sự của khoản kế hoạch.');
    }
  }

  const sourceList = getNguonTienList_(session.maKyHoc);
  const source = sourceList.find(item => item.maNguonTien === maNguonTien);

  if (!source || source.trangThai !== 'ACTIVE') {
    throw new Error('Nguồn tiền không tồn tại hoặc đã ngừng sử dụng.');
  }

  const uploadedAttachment = chungTuImage && chungTuImage.dataUrl
    ? saveThuChiAttachment_(chungTuImage, loai, ngayText)
    : null;

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang có người cập nhật sổ thu chi. Vui lòng thao tác lại.');
  }

  let savedId = '';
  let savedSoPhieu = '';

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureSheet_(ss, SHEET_SOTHUCHI, getSoThuChiHeaders_());
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    const now = new Date();

    let targetIndex = -1;
    let id = maGiaoDich;

    if (id) {
      for (let i = 1; i < values.length; i++) {
        if (
          String(values[i][index.MaGiaoDich] || '').trim() === id &&
          String(values[i][index.MaKyHoc] || '').trim() === session.maKyHoc
        ) {
          targetIndex = i;
          break;
        }
      }

      if (targetIndex === -1) {
        throw new Error('Không tìm thấy giao dịch cần sửa.');
      }

      const sourceType = String(
        values[targetIndex][index.NguonDuLieu] || ''
      ).trim().toUpperCase();
      const existingType = String(
        values[targetIndex][index.LoaiGiaoDich] || ''
      ).trim().toUpperCase();
      const existingStatus = String(
        values[targetIndex][index.TrangThai] || 'HOAT_DONG'
      ).trim().toUpperCase();

      if (sourceType !== 'NHAP_TAY') {
        throw new Error('Giao dịch tự động chỉ được xem và in; vui lòng điều chỉnh tại màn hình dữ liệu nguồn.');
      }

      if (existingStatus === 'DA_HUY') {
        throw new Error('Phiếu đã huỷ không thể sửa. Vui lòng lập phiếu mới nếu cần điều chỉnh.');
      }

      if (existingType && existingType !== loai) {
        throw new Error('Không thể đổi loại phiếu sau khi đã phát sinh số phiếu.');
      }
    } else {
      id = 'TC_' + Utilities.getUuid().slice(0, 12).toUpperCase();
    }

    const currentRow = targetIndex >= 0
      ? values[targetIndex].slice()
      : new Array(headers.length).fill('');
    const existingSoPhieu = targetIndex >= 0 && index.SoPhieu !== undefined
      ? String(currentRow[index.SoPhieu] || '').trim()
      : '';
    const soPhieu = existingSoPhieu || generateNextSoPhieuFromRows_(
      loai,
      ngayText,
      values.slice(1),
      index
    );

    savedId = id;
    savedSoPhieu = soPhieu;

    currentRow[index.MaGiaoDich] = id;
    currentRow[index.NgayGiaoDich] = parseInputDate_(ngayText);
    currentRow[index.MaKyHoc] = session.maKyHoc;
    currentRow[index.LoaiGiaoDich] = loai;
    currentRow[index.MaDanhMuc] = maDanhMuc;
    currentRow[index.MaKeHoachChi] = loai === 'CHI' ? maKeHoachChi : '';
    currentRow[index.MaNhanSu] = loai === 'CHI' ? maNhanSu : '';
    currentRow[index.MaHuTaiChinh] = maHuTaiChinh;
    currentRow[index.TenDanhMuc] = String(category.TenDanhMuc || '').trim();
    currentRow[index.NoiDung] = noiDung;
    currentRow[index.SoTien] = soTien;
    currentRow[index.HinhThuc] = getHinhThucByNguon_(maNguonTien);
    currentRow[index.MaNguonTien] = maNguonTien;
    currentRow[index.TenNguonTien] = source.tenNguonTien;
    currentRow[index.MaNguonDoiUng] = '';
    currentRow[index.MaNhomChuyen] = '';
    currentRow[index.NguoiNopNhan] = nguoiNopNhan;
    currentRow[index.SoPhieu] = soPhieu;
    currentRow[index.SoChungTu] = soChungTu;
    if (uploadedAttachment) {
      currentRow[index.ChungTuFileId] = uploadedAttachment.fileId;
      currentRow[index.ChungTuUrl] = uploadedAttachment.url;
    }
    currentRow[index.GhiChu] = ghiChu;
    currentRow[index.NguonDuLieu] = 'NHAP_TAY';
    currentRow[index.MaThamChieu] = '';
    currentRow[index.PhamVi] = 'TRUNG_TAM';
    currentRow[index.MaDanhMucGiaDinh] = '';
    currentRow[index.LoaiGiaDinh] = '';
    currentRow[index.MaGiaoDichLienKet] = '';
    currentRow[index.TrangThai] = 'HOAT_DONG';
    currentRow[index.NguoiTao] = session.tenDangNhap || session.hoTen || session.maNguoiDung;
    currentRow[index.CreatedAt] = targetIndex >= 0
      ? (currentRow[index.CreatedAt] || now)
      : now;
    currentRow[index.UpdatedAt] = now;

    if (targetIndex >= 0) {
      sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([currentRow]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([currentRow]);
    }
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, maGiaoDich ? 'UPDATE' : 'CREATE', 'SO_THU_CHI', savedId, null, {
    soPhieu: savedSoPhieu, loai: loai, soTien: soTien, ngayGiaoDich: ngayText, maDanhMuc: maDanhMuc, maHuTaiChinh: maHuTaiChinh
  });

  return jsonResponse_({
    success: true,
    maGiaoDich: savedId,
    soPhieu: savedSoPhieu,
    chungTuFileId: uploadedAttachment ? uploadedAttachment.fileId : '',
    chungTuUrl: uploadedAttachment ? uploadedAttachment.url : '',
    message: maGiaoDich
      ? 'Đã cập nhật ' + (loai === 'CHI' ? 'phiếu chi ' : 'phiếu thu ') + savedSoPhieu + '.'
      : 'Đã tạo ' + (loai === 'CHI' ? 'phiếu chi ' : 'phiếu thu ') + savedSoPhieu + '.'
  });
}

function saveThuChiAttachment_(imageData, loai, ngayText) {
  const dataUrl = String(imageData.dataUrl || '');
  const match = dataUrl.match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error('Ảnh chứng từ không hợp lệ. Chỉ hỗ trợ JPG, PNG hoặc WEBP.');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('Ảnh chứng từ sau khi xử lý không được vượt quá 5 MB.');

  const props = PropertiesService.getScriptProperties();
  let folder = null;
  const folderId = props.getProperty('THUCHI_ATTACHMENT_FOLDER_ID');
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (error) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder('ThienAnEduApp_ChungTuThuChi');
    props.setProperty('THUCHI_ATTACHMENT_FOLDER_ID', folder.getId());
  }

  const extension = match[1] === 'image/png' ? 'png' : (match[1] === 'image/webp' ? 'webp' : 'jpg');
  const safeDate = String(ngayText || '').replace(/[^\d]/g, '') || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const fileName = String(loai || 'CT') + '_' + safeDate + '_' + Utilities.getUuid().slice(0, 8) + '.' + extension;
  const blob = Utilities.newBlob(bytes, match[1], fileName);
  const file = folder.createFile(blob);
  if (String(props.getProperty('ALLOW_PUBLIC_ATTACHMENT_LINKS') || 'FALSE').toUpperCase() === 'TRUE') try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    // Một số Google Workspace không cho phép chia sẻ công khai; file vẫn giữ quyền riêng tư.
  }
  return { fileId: file.getId(), url: file.getUrl(), name: fileName };
}

function saveChuyenNguonTien(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);

  const command = CashbookDomain.transfer(Object.assign({}, data || {}, { soTien: number_(data && data.soTien) }));
  const ngayText = command.ngayGiaoDich;
  const maNguonDi = command.maNguonDi;
  const maNguonDen = command.maNguonDen;
  const soTien = command.soTien;
  const noiDung = command.noiDung;
  const soChungTu = command.soChungTu;
  const ghiChu = command.ghiChu;
  const scope = normalizeFinanceScope_(data && data.phamVi || 'TRUNG_TAM');
  const familyCategoryId = scope === 'GIA_DINH' ? String(data && data.maDanhMucGiaDinh || '').trim() : '';
  const familySavingCategory = familyCategoryId ? getDanhMucGiaDinhList_().find(item => item.maDanhMucGiaDinh === familyCategoryId && item.loai === 'TIET_KIEM' && item.trangThai === 'ACTIVE') : null;
  if (familyCategoryId && !familySavingCategory) throw new Error('Mục tiêu tiết kiệm gia đình không hợp lệ.');
  if (scope === 'GIA_DINH' && maNguonDen === 'GD_TIET_KIEM' && !familySavingCategory) throw new Error('Vui lòng chọn mục tiêu tiết kiệm khi chuyển tiền vào tài khoản tiết kiệm.');

  assertFinancePeriodOpen_(session, ngayText.slice(0, 7));

  const sourceList = getNguonTienList_(session.maKyHoc, scope);
  const sourceFrom = sourceList.find(item => item.maNguonTien === maNguonDi);
  const sourceTo = sourceList.find(item => item.maNguonTien === maNguonDen);

  if (!sourceFrom || !sourceTo) {
    throw new Error('Nguồn tiền chuyển hoặc nhận không hợp lệ.');
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang cập nhật nguồn tiền. Vui lòng thao tác lại.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureSheet_(ss, SHEET_SOTHUCHI, getSoThuChiHeaders_());
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    const now = new Date();
    const groupId = 'CN_' + Utilities.getUuid().slice(0, 12).toUpperCase();

    const outRow = new Array(headers.length).fill('');
    outRow[index.MaGiaoDich] = 'TC_RA_' + Utilities.getUuid().slice(0, 10).toUpperCase();
    outRow[index.NgayGiaoDich] = parseInputDate_(ngayText);
    outRow[index.MaKyHoc] = session.maKyHoc;
    outRow[index.LoaiGiaoDich] = 'CHI';
    outRow[index.MaDanhMuc] = 'CHI_CHUYEN_NOI_BO';
    outRow[index.MaHuTaiChinh] = '';
    outRow[index.TenDanhMuc] = 'Chuyển tiền nội bộ';
    outRow[index.NoiDung] = noiDung + ' - ' + sourceFrom.tenNguonTien + ' → ' + sourceTo.tenNguonTien;
    outRow[index.SoTien] = soTien;
    outRow[index.HinhThuc] = 'Chuyển nội bộ';
    outRow[index.MaNguonTien] = sourceFrom.maNguonTien;
    outRow[index.TenNguonTien] = sourceFrom.tenNguonTien;
    outRow[index.MaNguonDoiUng] = sourceTo.maNguonTien;
    outRow[index.MaNhomChuyen] = groupId;
    outRow[index.NguoiNopNhan] = sourceTo.tenNguonTien;
    outRow[index.SoChungTu] = soChungTu;
    outRow[index.GhiChu] = ghiChu;
    outRow[index.NguonDuLieu] = scope === 'GIA_DINH' ? 'GIA_DINH_CHUYEN' : 'CHUYEN_NOI_BO';
    outRow[index.MaThamChieu] = groupId;
    outRow[index.PhamVi] = scope;
    outRow[index.LoaiGiaDinh] = scope === 'GIA_DINH' ? 'CHUYEN' : '';
    outRow[index.MaGiaoDichLienKet] = groupId;
    outRow[index.TrangThai] = 'HOAT_DONG';
    outRow[index.NguoiTao] = session.tenDangNhap || session.hoTen || session.maNguoiDung;
    outRow[index.CreatedAt] = now;
    outRow[index.UpdatedAt] = now;

    const inRow = new Array(headers.length).fill('');
    inRow[index.MaGiaoDich] = 'TC_VAO_' + Utilities.getUuid().slice(0, 10).toUpperCase();
    inRow[index.NgayGiaoDich] = parseInputDate_(ngayText);
    inRow[index.MaKyHoc] = session.maKyHoc;
    inRow[index.LoaiGiaoDich] = 'THU';
    inRow[index.MaDanhMuc] = 'THU_CHUYEN_NOI_BO';
    inRow[index.MaHuTaiChinh] = '';
    inRow[index.TenDanhMuc] = familySavingCategory ? familySavingCategory.tenDanhMuc : 'Nhận chuyển tiền nội bộ';
    inRow[index.NoiDung] = noiDung + ' - ' + sourceFrom.tenNguonTien + ' → ' + sourceTo.tenNguonTien;
    inRow[index.SoTien] = soTien;
    inRow[index.HinhThuc] = 'Chuyển nội bộ';
    inRow[index.MaNguonTien] = sourceTo.maNguonTien;
    inRow[index.TenNguonTien] = sourceTo.tenNguonTien;
    inRow[index.MaNguonDoiUng] = sourceFrom.maNguonTien;
    inRow[index.MaNhomChuyen] = groupId;
    inRow[index.NguoiNopNhan] = sourceFrom.tenNguonTien;
    inRow[index.SoChungTu] = soChungTu;
    inRow[index.GhiChu] = ghiChu;
    inRow[index.NguonDuLieu] = scope === 'GIA_DINH' ? 'GIA_DINH_CHUYEN' : 'CHUYEN_NOI_BO';
    inRow[index.MaThamChieu] = groupId;
    inRow[index.PhamVi] = scope;
    inRow[index.MaDanhMucGiaDinh] = familySavingCategory ? familySavingCategory.maDanhMucGiaDinh : '';
    inRow[index.LoaiGiaDinh] = scope === 'GIA_DINH' ? 'CHUYEN' : '';
    inRow[index.MaGiaoDichLienKet] = groupId;
    inRow[index.TrangThai] = 'HOAT_DONG';
    inRow[index.NguoiTao] = session.tenDangNhap || session.hoTen || session.maNguoiDung;
    inRow[index.CreatedAt] = now;
    inRow[index.UpdatedAt] = now;

    sheet
      .getRange(sheet.getLastRow() + 1, 1, 2, headers.length)
      .setValues([outRow, inRow]);
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'CREATE', 'CHUYEN_NGUON_TIEN', ngayText, null, { maNguonDi: maNguonDi, maNguonDen: maNguonDen, soTien: soTien });

  return jsonResponse_({
    success: true,
    message: 'Đã chuyển tiền giữa hai nguồn. Giao dịch không làm tăng tổng thu hoặc tổng chi.'
  });
}

function cancelChuyenNguonTien(token, maNhomChuyen) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);

  const groupId = String(maNhomChuyen || '').trim();
  if (!groupId) throw new Error('Thiếu mã nhóm chuyển tiền.');
  const currentTransfer = readObjectsNoCache_(SHEET_SOTHUCHI).find(row =>
    String(row.MaNhomChuyen || '').trim() === groupId && String(row.MaKyHoc || '').trim() === session.maKyHoc
  );
  if (currentTransfer) assertFinancePeriodOpen_(session, formatDateForInput_(currentTransfer.NgayGiaoDich).slice(0, 7));

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang cập nhật nguồn tiền. Vui lòng thao tác lại.');
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SOTHUCHI);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    const targets = [];

    for (let i = 1; i < values.length; i++) {
      const rowGroup = String(values[i][index.MaNhomChuyen] || '').trim();
      const rowKyHoc = String(values[i][index.MaKyHoc] || '').trim();
      const rowSource = String(values[i][index.NguonDuLieu] || '').trim().toUpperCase();

      if (
        rowGroup === groupId &&
        rowKyHoc === session.maKyHoc &&
        rowSource === 'CHUYEN_NOI_BO'
      ) {
        targets.push(i);
      }
    }

    if (targets.length !== 2) {
      throw new Error('Không tìm thấy đầy đủ hai bút toán của giao dịch chuyển tiền.');
    }

    const now = new Date();

    targets.forEach(arrayIndex => {
      const row = values[arrayIndex].slice();
      row[index.TrangThai] = 'DA_HUY';
      row[index.UpdatedAt] = now;

      sheet
        .getRange(arrayIndex + 1, 1, 1, headers.length)
        .setValues([row]);
    });
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'CANCEL', 'CHUYEN_NGUON_TIEN', groupId, currentTransfer, null);

  return jsonResponse_({
    success: true,
    message: 'Đã huỷ giao dịch chuyển tiền nội bộ.'
  });
}

function saveNguonTienConfig(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);

  data = data || {};

  const maNguonTien = String(data.maNguonTien || '').trim().toUpperCase();
  const soDuBanDau = number_(data.soDuBanDau);

  if (!getNguonTienDefinition_(maNguonTien)) {
    throw new Error('Nguồn tiền không hợp lệ.');
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang cập nhật cấu hình nguồn tiền. Vui lòng thao tác lại.');
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NGUONTIEN);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    let targetIndex = -1;

    for (let i = 1; i < values.length; i++) {
      if (
        String(values[i][index.MaKyHoc] || '').trim() === session.maKyHoc &&
        String(values[i][index.MaNguonTien] || '').trim().toUpperCase() === maNguonTien
      ) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex === -1) {
      throw new Error('Không tìm thấy cấu hình nguồn tiền.');
    }

    const row = values[targetIndex].slice();
    row[index.SoDuBanDau] = soDuBanDau;
    row[index.UpdatedAt] = new Date();

    sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'UPDATE', 'NGUON_TIEN', maNguonTien, null, { soDuBanDau: soDuBanDau });

  return jsonResponse_({
    success: true,
    message: 'Đã cập nhật số dư ban đầu của ' + getNguonTienName_(maNguonTien) + '.'
  });
}


function cancelThuChiGiaoDich(token, maGiaoDich) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_();

  const id = String(maGiaoDich || '').trim();
  if (!id) throw new Error('Thiếu mã giao dịch.');
  const currentTransaction = readObjectsNoCache_(SHEET_SOTHUCHI).find(row =>
    String(row.MaGiaoDich || '').trim() === id && String(row.MaKyHoc || '').trim() === session.maKyHoc
  );
  if (currentTransaction) assertFinancePeriodOpen_(session, formatDateForInput_(currentTransaction.NgayGiaoDich).slice(0, 7));

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang có người cập nhật sổ thu chi. Vui lòng thao tác lại.');
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SOTHUCHI);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    let targetIndex = -1;

    for (let i = 1; i < values.length; i++) {
      if (
        String(values[i][index.MaGiaoDich] || '').trim() === id &&
        String(values[i][index.MaKyHoc] || '').trim() === session.maKyHoc
      ) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex === -1) throw new Error('Không tìm thấy giao dịch cần huỷ.');

    const source = String(values[targetIndex][index.NguonDuLieu] || '').trim().toUpperCase();
    if (source !== 'NHAP_TAY') {
      throw new Error('Khoản thu học phí tự động phải được điều chỉnh tại trang Quản lý thu phí.');
    }

    const row = values[targetIndex].slice();
    row[index.TrangThai] = 'DA_HUY';
    row[index.UpdatedAt] = new Date();
    sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'CANCEL', 'SO_THU_CHI', id, currentTransaction, null);

  return jsonResponse_({
    success: true,
    message: 'Đã huỷ giao dịch. Dữ liệu vẫn được giữ để đối chiếu.'
  });
}

function saveDanhMucThuChi(token, data) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_();

  data = data || {};

  const inputId = String(data.maDanhMuc || '').trim();
  const loai = String(data.loai || '').trim().toUpperCase();
  const tenDanhMuc = String(data.tenDanhMuc || '').trim();
  const maHuMacDinh = loai === 'CHI' ? normalizeMaHuTaiChinh_(data.maHuMacDinh) : '';
  const thuTu = number_(data.thuTu) || 999;

  if (loai !== 'THU' && loai !== 'CHI') throw new Error('Loại danh mục không hợp lệ.');
  if (!tenDanhMuc) throw new Error('Vui lòng nhập tên danh mục.');
  if (loai === 'CHI' && !getDanhMucHuTaiChinhList_(false).some(item => item.code === maHuMacDinh)) throw new Error('Vui lòng chọn hũ tài chính mặc định đang hoạt động cho danh mục chi.');

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang cập nhật danh mục. Vui lòng thao tác lại.');
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DANHMUC_THUCHI);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    let targetIndex = -1;

    if (inputId) {
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][index.MaDanhMuc] || '').trim() === inputId) {
          targetIndex = i;
          break;
        }
      }

      if (targetIndex === -1) throw new Error('Không tìm thấy danh mục cần sửa.');

      const oldType = String(values[targetIndex][index.Loai] || '').trim().toUpperCase();
      if (oldType !== loai) {
        throw new Error('Không thể đổi loại Thu/Chi của danh mục đã tạo.');
      }
    }

    const duplicated = values.slice(1).some((row, rowIndex) => {
      if (targetIndex === rowIndex + 1) return false;
      return String(row[index.Loai] || '').trim().toUpperCase() === loai &&
        normalizeText_(row[index.TenDanhMuc]) === normalizeText_(tenDanhMuc) &&
        String(row[index.TrangThai] || '').trim().toUpperCase() !== 'DELETED';
    });

    if (duplicated) throw new Error('Tên danh mục đã tồn tại trong nhóm này.');

    const now = new Date();
    const id = inputId || ('DM_' + loai + '_' + Utilities.getUuid().slice(0, 8).toUpperCase());
    const row = targetIndex >= 0
      ? values[targetIndex].slice()
      : new Array(headers.length).fill('');

    row[index.MaDanhMuc] = id;
    row[index.Loai] = loai;
    row[index.TenDanhMuc] = tenDanhMuc;
    row[index.MaHuMacDinh] = maHuMacDinh;
    row[index.ThuTu] = thuTu;
    row[index.TrangThai] = targetIndex >= 0
      ? (row[index.TrangThai] || 'ACTIVE')
      : 'ACTIVE';
    row[index.CreatedAt] = targetIndex >= 0 ? (row[index.CreatedAt] || now) : now;
    row[index.UpdatedAt] = now;

    if (targetIndex >= 0) {
      sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
    }
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, inputId ? 'UPDATE' : 'CREATE', 'DANH_MUC_THU_CHI', inputId || tenDanhMuc, null, { loai: loai, tenDanhMuc: tenDanhMuc, maHuMacDinh: maHuMacDinh });

  return jsonResponse_({
    success: true,
    message: inputId ? 'Đã cập nhật danh mục.' : 'Đã thêm danh mục.'
  });
}

function toggleDanhMucThuChi(token, maDanhMuc, enabled) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_();

  const id = String(maDanhMuc || '').trim();
  if (!id) throw new Error('Thiếu mã danh mục.');
  if (id === 'THU_HOC_PHI' && !toBoolean_(enabled)) {
    throw new Error('Danh mục Thu học phí là danh mục hệ thống, không thể ngừng sử dụng.');
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang cập nhật danh mục. Vui lòng thao tác lại.');
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DANHMUC_THUCHI);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header || '').trim());
    const index = buildHeaderIndex_(headers);
    let targetIndex = -1;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][index.MaDanhMuc] || '').trim() === id) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex === -1) throw new Error('Không tìm thấy danh mục.');

    const row = values[targetIndex].slice();
    row[index.TrangThai] = toBoolean_(enabled) ? 'ACTIVE' : 'INACTIVE';
    row[index.UpdatedAt] = new Date();
    sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'STATUS', 'DANH_MUC_THU_CHI', id, null, { enabled: toBoolean_(enabled) });

  return jsonResponse_({
    success: true,
    message: toBoolean_(enabled) ? 'Đã kích hoạt danh mục.' : 'Đã ngừng sử dụng danh mục.'
  });
}

function syncHocPhiToThuChi(token, yearMonth) {
  const session = requireSession_(token, 'finance.write');
  ensureThuChiSheets_(session.maKyHoc);

  const ym = parseYearMonth_(yearMonth);
  const monthKey = ym.year + '-' + String(ym.month).padStart(2, '0');
  assertFinancePeriodOpen_(session, monthKey);
  const feeSheetName = getThuPhiMonthSheetName_(ym.year, ym.month);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feeSheet = ss.getSheetByName(feeSheetName);

  if (!feeSheet) {
    throw new Error('Chưa có dữ liệu thu phí tại sheet ' + feeSheetName + '.');
  }

  ensureSheet_(ss, feeSheetName, getThuPhiMonthHeaders_());

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Hệ thống đang đồng bộ dữ liệu tài chính. Vui lòng thao tác lại.');
  }

  let createdCount = 0;
  let updatedCount = 0;
  let cancelledCount = 0;
  let unassignedCount = 0;
  let receiptCount = 0;

  try {
    const feeValues = feeSheet.getDataRange().getValues();
    if (!feeValues || feeValues.length < 2) {
      throw new Error('Sheet thu phí tháng chưa có dữ liệu học sinh.');
    }

    const feeHeaders = feeValues[0].map(header => String(header || '').trim());
    const feeIndex = buildHeaderIndex_(feeHeaders);

    const ledgerSheet = ss.getSheetByName(SHEET_SOTHUCHI);
    const ledgerValues = ledgerSheet.getDataRange().getValues();
    const ledgerHeaders = ledgerValues[0].map(header => String(header || '').trim());
    const ledgerIndex = buildHeaderIndex_(ledgerHeaders);
    const ledgerRows = ledgerValues.slice(1).map(row => row.slice());

    const referenceMap = {};
    ledgerRows.forEach((row, rowIndex) => {
      const ref = String(row[ledgerIndex.MaThamChieu] || '').trim();
      const ky = String(row[ledgerIndex.MaKyHoc] || '').trim();
      if (ref && ky === session.maKyHoc) referenceMap[ref] = rowIndex;
    });

    const now = new Date();
    let feeChanged = false;

    feeValues.slice(1).forEach(feeRow => {
      const rowKyHoc = String(feeRow[feeIndex.MaKyHoc] || '').trim();
      const rowStatus = String(feeRow[feeIndex.TrangThai] || '').trim().toUpperCase();
      const maHocSinh = String(feeRow[feeIndex.MaHocSinh] || '').trim();

      if (rowKyHoc !== session.maKyHoc || rowStatus === 'DELETED' || !maHocSinh) return;

      const amount = number_(feeRow[feeIndex.SoTienDaThu]);
      const tamNghi = toBoolean_(feeRow[feeIndex.TamNghi]);
      const reference = buildHocPhiReference_(session.maKyHoc, monthKey, maHocSinh);
      const existingIndex = referenceMap[reference];
      const shouldBeActive = !tamNghi && amount > 0;
      const savedSource = feeIndex.NguonTienThu !== undefined
        ? String(feeRow[feeIndex.NguonTienThu] || '').trim().toUpperCase()
        : '';
      const sourceCode = savedSource || inferNguonTienFromLegacy_(
        feeRow[feeIndex.HinhThucThu],
        ''
      );

      if (shouldBeActive && !sourceCode) unassignedCount++;

      if (existingIndex !== undefined) {
        const row = ledgerRows[existingIndex];
        const oldStatus = String(row[ledgerIndex.TrangThai] || '').trim().toUpperCase();
        const existingReceipt = String(row[ledgerIndex.SoPhieu] || '').trim() ||
          String(feeRow[feeIndex.SoPhieu] || '').trim();

        if (!shouldBeActive) {
          row[ledgerIndex.TrangThai] = 'DA_HUY';
          row[ledgerIndex.GhiChu] = String(feeRow[feeIndex.GhiChu] || '').trim();
          row[ledgerIndex.UpdatedAt] = now;
          if (existingReceipt && !String(feeRow[feeIndex.SoPhieu] || '').trim()) {
            feeRow[feeIndex.SoPhieu] = existingReceipt;
            feeChanged = true;
          }
          if (oldStatus === 'HOAT_DONG') cancelledCount++;
          else updatedCount++;
          return;
        }

        const receipt = existingReceipt || generateNextSoPhieuFromRows_(
          'HOC_PHI',
          feeRow[feeIndex.NgayDong] || now,
          ledgerRows,
          ledgerIndex
        );

        if (!existingReceipt) receiptCount++;
        if (String(feeRow[feeIndex.SoPhieu] || '').trim() !== receipt) {
          feeRow[feeIndex.SoPhieu] = receipt;
          feeChanged = true;
        }

        fillHocPhiLedgerRow_(row, ledgerIndex, {
          maKyHoc: session.maKyHoc,
          maHocSinh: maHocSinh,
          yearMonth: monthKey,
          hoTen: String(feeRow[feeIndex.HoTen] || '').trim(),
          ngayThu: feeRow[feeIndex.NgayDong] || '',
          soTien: amount,
          maNguonTien: sourceCode,
          ghiChu: String(feeRow[feeIndex.GhiChu] || '').trim(),
          active: true,
          createdAt: row[ledgerIndex.CreatedAt] || now,
          updatedAt: now,
          reference: reference,
          maGiaoDich: row[ledgerIndex.MaGiaoDich],
          soPhieu: receipt
        });

        updatedCount++;
        return;
      }

      if (!shouldBeActive) return;

      const receipt = String(feeRow[feeIndex.SoPhieu] || '').trim() ||
        generateNextSoPhieuFromRows_(
          'HOC_PHI',
          feeRow[feeIndex.NgayDong] || now,
          ledgerRows,
          ledgerIndex
        );

      if (!String(feeRow[feeIndex.SoPhieu] || '').trim()) receiptCount++;
      feeRow[feeIndex.SoPhieu] = receipt;
      feeChanged = true;

      const newRow = new Array(ledgerHeaders.length).fill('');
      fillHocPhiLedgerRow_(newRow, ledgerIndex, {
        maKyHoc: session.maKyHoc,
        maHocSinh: maHocSinh,
        yearMonth: monthKey,
        hoTen: String(feeRow[feeIndex.HoTen] || '').trim(),
        ngayThu: feeRow[feeIndex.NgayDong] || '',
        soTien: amount,
        maNguonTien: sourceCode,
        ghiChu: String(feeRow[feeIndex.GhiChu] || '').trim(),
        active: true,
        createdAt: now,
        updatedAt: now,
        reference: reference,
        maGiaoDich: 'TC_HP_' + Utilities.getUuid().slice(0, 10).toUpperCase(),
        soPhieu: receipt
      });

      referenceMap[reference] = ledgerRows.length;
      ledgerRows.push(newRow);
      createdCount++;
    });

    if (ledgerRows.length) {
      ledgerSheet.getRange(2, 1, ledgerRows.length, ledgerHeaders.length).setValues(ledgerRows);
    }

    if (feeChanged) {
      feeSheet.getRange(2, 1, feeValues.length - 1, feeHeaders.length)
        .setValues(feeValues.slice(1));
    }
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();
  safeWriteAuditLog_(session, 'SYNC', 'THU_PHI_SO_THU_CHI', monthKey, null, {
    createdCount: createdCount, updatedCount: updatedCount, cancelledCount: cancelledCount,
    receiptCount: receiptCount, unassignedCount: unassignedCount
  });

  return jsonResponse_({
    success: true,
    createdCount: createdCount,
    updatedCount: updatedCount,
    cancelledCount: cancelledCount,
    receiptCount: receiptCount,
    unassignedCount: unassignedCount,
    message: 'Đã đồng bộ học phí tháng ' + String(ym.month).padStart(2, '0') + '/' + ym.year +
      ': thêm ' + createdCount + ', cập nhật ' + updatedCount + ', huỷ ' + cancelledCount +
      ', cấp mới ' + receiptCount + ' số phiếu' +
      (unassignedCount > 0 ? '. Còn ' + unassignedCount + ' khoản chưa chọn nguồn tiền.' : '.')
  });
}

function upsertThuChiHocPhiNoLock_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SOTHUCHI);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(header => String(header || '').trim());
  const index = buildHeaderIndex_(headers);
  const reference = buildHocPhiReference_(data.maKyHoc, data.yearMonth, data.maHocSinh);
  let targetIndex = -1;

  for (let i = 1; i < values.length; i++) {
    if (
      String(values[i][index.MaKyHoc] || '').trim() === data.maKyHoc &&
      String(values[i][index.MaThamChieu] || '').trim() === reference
    ) {
      targetIndex = i;
      break;
    }
  }

  const now = new Date();
  const active = !data.tamNghi && number_(data.soTien) > 0;

  if (targetIndex === -1 && !active) {
    return {
      soPhieu: String(data.soPhieu || '').trim(),
      maGiaoDich: ''
    };
  }

  const row = targetIndex >= 0
    ? values[targetIndex].slice()
    : new Array(headers.length).fill('');

  const existingSoPhieu = targetIndex >= 0 && index.SoPhieu !== undefined
    ? String(row[index.SoPhieu] || '').trim()
    : '';
  const soPhieu = existingSoPhieu || String(data.soPhieu || '').trim() || (
    active
      ? generateNextSoPhieuFromRows_('HOC_PHI', data.ngayThu || now, values.slice(1), index)
      : ''
  );

  if (targetIndex >= 0 && !active) {
    row[index.TrangThai] = 'DA_HUY';
    row[index.GhiChu] = String(data.ghiChu || '').trim();
    if (index.SoPhieu !== undefined && soPhieu) row[index.SoPhieu] = soPhieu;
    row[index.UpdatedAt] = now;
    sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);

    return {
      soPhieu: soPhieu,
      maGiaoDich: String(row[index.MaGiaoDich] || '').trim()
    };
  }

  const maGiaoDich = targetIndex >= 0
    ? String(row[index.MaGiaoDich] || '').trim()
    : ('TC_HP_' + Utilities.getUuid().slice(0, 10).toUpperCase());

  fillHocPhiLedgerRow_(row, index, {
    maKyHoc: data.maKyHoc,
    maHocSinh: data.maHocSinh,
    yearMonth: data.yearMonth,
    hoTen: data.hoTen,
    ngayThu: data.ngayThu,
    soTien: data.soTien,
    maNguonTien: data.maNguonTien,
    ghiChu: data.ghiChu,
    active: active,
    createdAt: targetIndex >= 0 ? (row[index.CreatedAt] || now) : now,
    updatedAt: now,
    reference: reference,
    maGiaoDich: maGiaoDich,
    soPhieu: soPhieu
  });

  if (targetIndex >= 0) {
    sheet.getRange(targetIndex + 1, 1, 1, headers.length).setValues([row]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
  }

  return {
    soPhieu: soPhieu,
    maGiaoDich: maGiaoDich
  };
}

function fillHocPhiLedgerRow_(row, index, data) {
  const monthParts = String(data.yearMonth || '').split('-');
  const monthText = monthParts.length === 2
    ? monthParts[1] + '/' + monthParts[0]
    : String(data.yearMonth || '');
  const maNguonTien = String(data.maNguonTien || '').trim().toUpperCase();
  const sourceName = getNguonTienName_(maNguonTien);

  row[index.MaGiaoDich] = data.maGiaoDich;
  row[index.NgayGiaoDich] = data.ngayThu || data.updatedAt || new Date();
  row[index.MaKyHoc] = data.maKyHoc;
  row[index.LoaiGiaoDich] = 'THU';
  row[index.MaDanhMuc] = 'THU_HOC_PHI';
  row[index.MaHuTaiChinh] = '';
  row[index.TenDanhMuc] = 'Thu học phí';
  row[index.NoiDung] = 'Thu học phí tháng ' + monthText + ' - ' + String(data.hoTen || '').trim();
  row[index.SoTien] = number_(data.soTien);
  row[index.HinhThuc] = getHinhThucByNguon_(maNguonTien);
  row[index.MaNguonTien] = maNguonTien;
  row[index.TenNguonTien] = sourceName;
  row[index.MaNguonDoiUng] = '';
  row[index.MaNhomChuyen] = '';
  row[index.NguoiNopNhan] = String(data.hoTen || '').trim();
  row[index.SoPhieu] = String(data.soPhieu || '').trim();
  row[index.SoChungTu] = '';
  row[index.GhiChu] = String(data.ghiChu || '').trim();
  row[index.NguonDuLieu] = 'HOC_PHI';
  row[index.MaThamChieu] = data.reference;
  row[index.PhamVi] = 'TRUNG_TAM';
  row[index.MaDanhMucGiaDinh] = '';
  row[index.LoaiGiaDinh] = '';
  row[index.MaGiaoDichLienKet] = '';
  row[index.TrangThai] = data.active ? 'HOAT_DONG' : 'DA_HUY';
  row[index.NguoiTao] = 'Hệ thống';
  row[index.CreatedAt] = data.createdAt;
  row[index.UpdatedAt] = data.updatedAt;
}

function buildHocPhiReference_(maKyHoc, yearMonth, maHocSinh) {
  return 'HP|' + String(maKyHoc || '').trim() + '|' +
    String(yearMonth || '').trim() + '|' + String(maHocSinh || '').trim();
}

function mapThuChiTransaction_(row) {
  const date = toDateOnly_(row.NgayGiaoDich);
  const source = String(row.NguonDuLieu || 'NHAP_TAY').trim().toUpperCase();
  const maNguonTien = String(row.MaNguonTien || '').trim().toUpperCase();
  const inferredSource = maNguonTien || inferNguonTienFromLegacy_(
    row.HinhThuc,
    row.TenNguonTien
  );
  const sourceName = String(row.TenNguonTien || '').trim() ||
    getNguonTienName_(inferredSource) ||
    'Chưa phân loại';
  const scope = normalizeFinanceScope_(row.PhamVi || (source.indexOf('GIA_DINH') === 0 ? 'GIA_DINH' : 'TRUNG_TAM'));

  return {
    maGiaoDich: String(row.MaGiaoDich || '').trim(),
    ngayDate: date,
    ngayGiaoDich: formatDateForInput_(date),
    ngayDisplay: formatDateDisplay_(date),
    loai: String(row.LoaiGiaoDich || '').trim().toUpperCase(),
    maDanhMuc: String(row.MaDanhMuc || '').trim(),
    maKeHoachChi: String(row.MaKeHoachChi || '').trim(),
    maNhanSu: String(row.MaNhanSu || '').trim(),
    maHuTaiChinh: normalizeMaHuTaiChinh_(row.MaHuTaiChinh),
    tenDanhMuc: String(row.TenDanhMuc || '').trim(),
    noiDung: String(row.NoiDung || '').trim(),
    soTien: number_(row.SoTien),
    hinhThuc: String(row.HinhThuc || '').trim(),
    maNguonTien: inferredSource,
    tenNguonTien: sourceName,
    maNguonDoiUng: String(row.MaNguonDoiUng || '').trim().toUpperCase(),
    maNhomChuyen: String(row.MaNhomChuyen || '').trim(),
    nguoiNopNhan: String(row.NguoiNopNhan || '').trim(),
    soPhieu: String(row.SoPhieu || '').trim(),
    soChungTu: String(row.SoChungTu || '').trim(),
    chungTuFileId: String(row.ChungTuFileId || '').trim(),
    chungTuUrl: String(row.ChungTuUrl || '').trim(),
    ghiChu: String(row.GhiChu || '').trim(),
    nguonDuLieu: source,
    maThamChieu: String(row.MaThamChieu || '').trim(),
    phamVi: scope,
    maDanhMucGiaDinh: String(row.MaDanhMucGiaDinh || '').trim(),
    loaiGiaDinh: String(row.LoaiGiaDinh || '').trim().toUpperCase(),
    maGiaoDichLienKet: String(row.MaGiaoDichLienKet || '').trim(),
    trangThai: String(row.TrangThai || 'HOAT_DONG').trim().toUpperCase(),
    nguoiTao: String(row.NguoiTao || '').trim(),
    createdAtRaw: row.CreatedAt || '',
    updatedAtRaw: row.UpdatedAt || '',
    editable: source === 'NHAP_TAY' && scope === 'TRUNG_TAM' &&
      String(row.TrangThai || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG',
    transfer: source === 'CHUYEN_NOI_BO' || source === 'GIA_DINH_CHUYEN',
    canCancelTransfer: (source === 'CHUYEN_NOI_BO' || source === 'GIA_DINH_CHUYEN') &&
      String(row.TrangThai || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG'
  };
}

function stripThuChiPrivateFields_(item) {
  return {
    maGiaoDich: item.maGiaoDich,
    ngayGiaoDich: item.ngayGiaoDich,
    ngayDisplay: item.ngayDisplay,
    loai: item.loai,
    maDanhMuc: item.maDanhMuc,
    maKeHoachChi: item.maKeHoachChi,
    maNhanSu: item.maNhanSu,
    maHuTaiChinh: item.maHuTaiChinh,
    tenDanhMuc: item.tenDanhMuc,
    noiDung: item.noiDung,
    soTien: item.soTien,
    hinhThuc: item.hinhThuc,
    maNguonTien: item.maNguonTien,
    tenNguonTien: item.tenNguonTien,
    maNguonDoiUng: item.maNguonDoiUng,
    maNhomChuyen: item.maNhomChuyen,
    nguoiNopNhan: item.nguoiNopNhan,
    soPhieu: item.soPhieu,
    soChungTu: item.soChungTu,
    chungTuFileId: item.chungTuFileId,
    chungTuUrl: item.chungTuUrl,
    ghiChu: item.ghiChu,
    nguonDuLieu: item.nguonDuLieu,
    maThamChieu: item.maThamChieu,
    phamVi: item.phamVi,
    maDanhMucGiaDinh: item.maDanhMucGiaDinh,
    loaiGiaDinh: item.loaiGiaDinh,
    maGiaoDichLienKet: item.maGiaoDichLienKet,
    trangThai: item.trangThai,
    nguoiTao: item.nguoiTao,
    editable: item.editable,
    transfer: item.transfer,
    canCancelTransfer: item.canCancelTransfer
  };
}

function buildThuChiMonthlyReport_(transactions, year) {
  const map = {};

  for (let month = 1; month <= 12; month++) {
    const key = year + '-' + String(month).padStart(2, '0');
    map[key] = { month: month, label: String(month).padStart(2, '0') + '/' + year, thu: 0, chi: 0, net: 0 };
  }

  transactions.forEach(item => {
    if (item.trangThai !== 'HOAT_DONG' || !item.ngayDate || item.ngayDate.getFullYear() !== year) return;
    const key = year + '-' + String(item.ngayDate.getMonth() + 1).padStart(2, '0');
    if (!map[key]) return;
    if (item.loai === 'THU') map[key].thu += item.soTien;
    if (item.loai === 'CHI') map[key].chi += item.soTien;
    map[key].net = map[key].thu - map[key].chi;
  });

  return Object.keys(map).sort().map(key => map[key]);
}

function buildThuChiCategoryReport_(transactions) {
  const map = {};

  transactions.forEach(item => {
    const key = item.loai + '|' + item.maDanhMuc;
    if (!map[key]) {
      map[key] = {
        loai: item.loai,
        maDanhMuc: item.maDanhMuc,
        tenDanhMuc: item.tenDanhMuc || item.maDanhMuc,
        soTien: 0,
        soGiaoDich: 0
      };
    }
    map[key].soTien += item.soTien;
    map[key].soGiaoDich++;
  });

  return Object.keys(map)
    .map(key => map[key])
    .sort((a, b) => b.soTien - a.soTien);
}

function buildThuChiPaymentMethodReport_(transactions) {
  const map = {};

  transactions.forEach(item => {
    const name = item.hinhThuc || 'Chưa xác định';
    if (!map[name]) map[name] = { hinhThuc: name, thu: 0, chi: 0 };
    if (item.loai === 'THU') map[name].thu += item.soTien;
    if (item.loai === 'CHI') map[name].chi += item.soTien;
  });

  return Object.keys(map)
    .map(key => map[key])
    .sort((a, b) => (b.thu + b.chi) - (a.thu + a.chi));
}

function resolveThuChiDateRange_(fromText, toText) {
  let fromDate = toDateOnly_(fromText);
  let toDate = toDateOnly_(toText);
  const now = new Date();

  if (!fromDate) fromDate = new Date(now.getFullYear(), now.getMonth(), 1);
  if (!toDate) toDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  if (fromDate.getTime() > toDate.getTime()) {
    const temp = fromDate;
    fromDate = toDate;
    toDate = temp;
  }

  return {
    fromDate: fromDate,
    toDate: toDate,
    fromText: formatDateForInput_(fromDate),
    toText: formatDateForInput_(toDate)
  };
}

function buildHeaderIndex_(headers) {
  return headers.reduce((map, header, index) => {
    if (header) map[header] = index;
    return map;
  }, {});
}


/* =========================================================
   COMMON SHEET / CACHE HELPERS
========================================================= */

function readObjects_(sheetName) {
  if (!isDataCacheEnabled_()) {
    return readObjectsNoCache_(sheetName);
  }

  const cacheKey = buildCacheKey_('sheet_' + sheetName);
  const cached = cacheGetString_(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const objects = readObjectsNoCache_(sheetName);
  cachePutString_(cacheKey, JSON.stringify(objects), CACHE_SECONDS);

  return objects;
}

function readObjectsNoCache_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) return [];

  const headers = values[0].map(h => String(h || '').trim());

  if (headers.filter(h => h).length === 0) return [];

  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        if (header) {
          obj[header] = row[index];
        }
      });

      return obj;
    });
}

function writeObjectsToSheet_(sheetName, objects, requiredHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, sheetName, requiredHeaders);

  const currentValues = sheet.getDataRange().getValues();
  const currentHeaders = currentValues[0]
    .map(h => String(h || '').trim())
    .filter(h => h);

  let headers = currentHeaders.length ? currentHeaders : requiredHeaders.slice();

  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
    }
  });

  const values = [headers].concat(
    objects.map(obj => {
      return headers.map(header => {
        return Object.prototype.hasOwnProperty.call(obj, header) ? obj[header] : '';
      });
    })
  );

  sheet.getRange(1, 1, values.length, headers.length).setValues(values);

  if (currentValues.length > values.length) {
    sheet.getRange(
      values.length + 1,
      1,
      currentValues.length - values.length,
      Math.max(1, sheet.getLastColumn())
    ).clearContent();
  }

  if (!currentHeaders.length || headers.length !== currentHeaders.length) {
    formatHeader_(sheet);
  }
}

function saveChuyenNguonTienGiaDinh(token, data) {
  return saveChuyenNguonTien(token, Object.assign({}, data || {}, { phamVi: 'GIA_DINH' }));
}

/**
 * Cáº­p nháº­t Ä‘Ãºng má»™t dÃ²ng theo khoÃ¡, khÃ´ng xoÃ¡ vÃ  ghi láº¡i toÃ n bá»™ sheet.
 */
function updateObjectRowById_(sheetName, idHeader, idValue, object, requiredHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ensureSheet_(ss, sheetName, requiredHeaders);
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  if (lastColumn <= 0 || lastRow <= 0) {
    throw new Error('Sheet ' + sheetName + ' chÆ°a cÃ³ cáº¥u trÃºc dá»¯ liá»‡u.');
  }

  let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(header => String(header || '').trim());
  let headersChanged = false;
  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
      headersChanged = true;
    }
  });
  if (headersChanged) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(sheet);
  }
  const idColumnIndex = headers.indexOf(idHeader);

  if (idColumnIndex === -1) {
    throw new Error('Sheet ' + sheetName + ' thiáº¿u cá»™t ' + idHeader + '.');
  }

  const ids = lastRow > 1
    ? sheet.getRange(2, idColumnIndex + 1, lastRow - 1, 1).getValues()
    : [];
  const targetOffset = ids.findIndex(row => String(row[0] || '').trim() === String(idValue || '').trim());

  if (targetOffset === -1) {
    throw new Error('KhÃ´ng tÃ¬m tháº¥y dá»¯ liá»‡u cáº§n cáº­p nháº­t trong sheet ' + sheetName + '.');
  }

  const rowValues = headers.map(header => {
    return Object.prototype.hasOwnProperty.call(object, header) ? object[header] : '';
  });

  sheet.getRange(targetOffset + 2, 1, 1, headers.length).setValues([rowValues]);
}

function compareStudentSort_(a, b) {
  const sortA = number_(a.sapXep || a.SapXep);
  const sortB = number_(b.sapXep || b.SapXep);
  const finalA = sortA > 0 ? sortA : 999999;
  const finalB = sortB > 0 ? sortB : 999999;

  const lopA = String(a.lop || a.Lop || '');
  const lopB = String(b.lop || b.Lop || '');
  const lopMatchA = lopA.match(/([1-9])/);
  const lopMatchB = lopB.match(/([1-9])/);
  const classNumberA = lopMatchA ? Number(lopMatchA[1]) : Math.floor(finalA / 100);
  const classNumberB = lopMatchB ? Number(lopMatchB[1]) : Math.floor(finalB / 100);

  // Lớp luôn tăng từ 1 đến 9.
  if (classNumberA !== classNumberB) return classNumberA - classNumberB;

  if (lopA !== lopB) return lopA.localeCompare(lopB, 'vi');

  // Cấp 2: THCS Phước Thái trước, THCS Long Phước sau.
  if (classNumberA >= 6 && classNumberA <= 9) {
    const schoolA = normalizeText_(a.truong || a.Truong || '');
    const schoolB = normalizeText_(b.truong || b.Truong || '');
    const schoolOrderA = schoolA === 'thcs phuoc thai' ? 1 : (schoolA === 'thcs long phuoc' ? 2 : 3);
    const schoolOrderB = schoolB === 'thcs phuoc thai' ? 1 : (schoolB === 'thcs long phuoc' ? 2 : 3);

    if (schoolOrderA !== schoolOrderB) return schoolOrderA - schoolOrderB;
  }

  // Trong cùng lớp/nhóm trường: Nam trước, Nữ sau.
  const genderA = normalizeText_(a.gioiTinh || a.GioiTinh || '');
  const genderB = normalizeText_(b.gioiTinh || b.GioiTinh || '');
  const genderOrderA = genderA === 'nam' ? 1 : (genderA === 'nu' ? 2 : 3);
  const genderOrderB = genderB === 'nam' ? 1 : (genderB === 'nu' ? 2 : 3);

  if (genderOrderA !== genderOrderB) return genderOrderA - genderOrderB;

  // SapXep là tiêu chí cố định cuối cùng trước tên.
  if (finalA !== finalB) return finalA - finalB;

  return String(a.hoTen || a.HoTen || '').localeCompare(String(b.hoTen || b.HoTen || ''), 'vi');
}

function parseYearMonth_(yearMonth) {
  const text = String(yearMonth || '').trim();

  if (!/^\d{4}-\d{2}$/.test(text)) {
    const now = new Date();

    return {
      year: now.getFullYear(),
      month: now.getMonth() + 1
    };
  }

  const parts = text.split('-');

  return {
    year: Number(parts[0]),
    month: Number(parts[1])
  };
}

function parseInputDate_(dateText) {
  if (!dateText) return '';

  const parts = String(dateText).split('-');

  if (parts.length !== 3) return dateText;

  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function formatDateForInput_(value) {
  const date = toDateOnly_(value);

  if (!date) return '';

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateDisplay_(value) {
  const date = toDateOnly_(value);

  if (!date) return '';

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function toDateOnly_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = String(value).trim();

  if (!text) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const parts = text.split('-');
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
    const parts = text.split('/');
    return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  }

  const parsed = new Date(text);

  if (isNaN(parsed.getTime())) return null;

  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function number_(value) {
  if (value === null || value === undefined || value === '') return 0;

  const number = Number(String(value).replace(/[^\d.-]/g, ''));

  return isNaN(number) ? 0 : number;
}

function moneyNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value).trim();
  const negative = text.charAt(0) === '-';
  const digits = text.replace(/\D/g, '');
  if (!digits) return 0;
  const parsed = Number(digits);
  return isNaN(parsed) ? 0 : (negative ? -parsed : parsed);
}

function toBoolean_(value) {
  if (value === true) return true;

  const text = String(value || '').trim().toUpperCase();
  return text === 'TRUE' || text === 'CÓ' || text === 'CO' || text === 'YES' || text === '1';
}

function defaultHocPhiByKhoi_(khoi) {
  const numberKhoi = Number(khoi);

  if (numberKhoi >= 1 && numberKhoi <= 5) return 1800000;

  return 2000000;
}

function normalizeText_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function jsonResponse_(data) {
  return JSON.stringify(data);
}

function getDataVersion_() {
  if (DATA_VERSION_MEMO_) return DATA_VERSION_MEMO_;

  const props = PropertiesService.getScriptProperties();
  let version = props.getProperty('DATA_VERSION');

  if (!version) {
    version = String(Date.now());
    props.setProperty('DATA_VERSION', version);
  }

  DATA_VERSION_MEMO_ = version;
  return version;
}

function bumpDataVersion_() {
  DATA_VERSION_MEMO_ = String(Date.now());
  PropertiesService.getScriptProperties().setProperty('DATA_VERSION', DATA_VERSION_MEMO_);
}

function buildCacheKey_(name) {
  return CACHE_PREFIX + getDataVersion_() + '_' + hashString_(name);
}

function hashString_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    String(value),
    Utilities.Charset.UTF_8
  );

  return bytes.map(byte => {
    const v = byte < 0 ? byte + 256 : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function cachePutString_(key, value, seconds) {
  if (!isDataCacheEnabled_()) return;

  const cache = CacheService.getScriptCache();
  const text = String(value || '');
  const count = Math.max(1, Math.ceil(text.length / CACHE_CHUNK_SIZE));

  cache.put(key + '_count', String(count), seconds);

  for (let i = 0; i < count; i++) {
    const part = text.slice(i * CACHE_CHUNK_SIZE, (i + 1) * CACHE_CHUNK_SIZE);
    cache.put(key + '_part_' + i, part, seconds);
  }
}

function cacheGetString_(key) {
  if (!isDataCacheEnabled_()) return null;

  const cache = CacheService.getScriptCache();
  const countText = cache.get(key + '_count');

  if (!countText) return null;

  const count = Number(countText);

  if (!count || count < 1) return null;

  const parts = [];

  for (let i = 0; i < count; i++) {
    const part = cache.get(key + '_part_' + i);

    if (part === null) return null;

    parts.push(part);
  }

  return parts.join('');
}


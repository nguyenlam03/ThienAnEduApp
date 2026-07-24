const APP_PASSWORD = 'ng0ctr@m';

const SHEET_KYHOC = 'KyHoc';
const SHEET_KHOI = 'Khoi';
const SHEET_LOP = 'Lop';
const SHEET_HOCSINH = 'HocSinh';
const SHEET_HOCSINH_KYHOC = 'HocSinhKyHoc';
const SHEET_THUPHI = 'ThuPhi';
const SHEET_DANHMUC_THUCHI = 'DanhMucThuChi';
const SHEET_SOTHUCHI = 'SoThuChi';
const SHEET_NGUONTIEN = 'NguonTien';

const CACHE_LOGIN_PREFIX = 'LOGIN_TOKEN_';
const CACHE_PREFIX = 'TA_CACHE_';
const CACHE_SECONDS = 1800;
const CACHE_CHUNK_SIZE = 90000;
const DATA_CACHE_ENABLED_PROPERTY = 'DATA_CACHE_ENABLED';

/**
 * Include file HTML dùng chung.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Điều hướng trang.
 */
function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : 'Login';

  const protectedPages = [
    'Index',
    'QuanLyHocSinh',
    'QuanLyThuPhu',
    'QuanLyThuChi'
  ];

  if (protectedPages.indexOf(page) !== -1) {
    const token = e.parameter.token || '';
    const session = getSessionFromToken_(token);

    if (!session.valid) {
      return HtmlService.createTemplateFromFile('Login')
        .evaluate()
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
    template.cacheEnabled = isDataCacheEnabled_();

    return template.evaluate()
      .setTitle('Thiên Ân Education')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .setTitle('Đăng nhập')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
   SETUP DATABASE
========================================================= */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

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

function checkLogin(password, maKyHoc) {
  const inputPassword = String(password || '').trim();
  const inputKyHoc = String(maKyHoc || '').trim();

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

  if (inputPassword !== APP_PASSWORD) {
    return jsonResponse_({
      success: false,
      message: 'Mật khẩu không đúng. Vui lòng kiểm tra lại.'
    });
  }

  const kyHocList = getKyHocArray_();
  const found = kyHocList.find(item => item.maKyHoc === inputKyHoc);

  if (!found) {
    return jsonResponse_({
      success: false,
      message: 'Kỳ học không hợp lệ hoặc chưa được kích hoạt.'
    });
  }

  const token = createLoginToken_(found.maKyHoc);
  const webAppUrl = ScriptApp.getService().getUrl();

  return jsonResponse_({
    success: true,
    message: 'Đăng nhập thành công.',
    maKyHoc: found.maKyHoc,
    tenKyHoc: found.tenKyHoc,
    token: token,
    redirectUrl: webAppUrl + '?page=Index&token=' + encodeURIComponent(token)
  });
}

function logout(token) {
  if (token) {
    CacheService.getScriptCache().remove(CACHE_LOGIN_PREFIX + token);
  }

  return jsonResponse_({
    success: true,
    message: 'Đã đăng xuất.'
  });
}

/**
 * Bật/tắt cache dữ liệu nghiệp vụ. Mặc định tắt.
 * Cache token đăng nhập vẫn hoạt động độc lập để duy trì phiên đăng nhập.
 */
function setDataCacheEnabled(token, enabled) {
  requireSession_(token);

  const isEnabled = enabled === true || String(enabled).toLowerCase() === 'true';
  PropertiesService.getScriptProperties().setProperty(
    DATA_CACHE_ENABLED_PROPERTY,
    isEnabled ? 'TRUE' : 'FALSE'
  );

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
  return String(
    PropertiesService.getScriptProperties().getProperty(DATA_CACHE_ENABLED_PROPERTY) || 'FALSE'
  ).toUpperCase() === 'TRUE';
}

/**
 * Xoá bộ nhớ tạm bằng cách đổi DATA_VERSION.
 * Cache cũ sẽ tự hết hạn, cache mới sẽ được tạo lại từ Google Sheets.
 */
function clearCacheAndReload(token) {
  const session = requireSession_(token);

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
  const token = Utilities.getUuid();

  CacheService.getScriptCache().put(
    CACHE_LOGIN_PREFIX + token,
    maKyHoc,
    21600
  );

  return token;
}

function getSessionFromToken_(token) {
  if (!token) {
    return {
      valid: false
    };
  }

  const maKyHoc = CacheService.getScriptCache().get(CACHE_LOGIN_PREFIX + token);

  if (!maKyHoc) {
    return {
      valid: false
    };
  }

  const kyHocList = getKyHocArray_();
  const found = kyHocList.find(item => item.maKyHoc === maKyHoc);

  if (!found) {
    return {
      valid: false
    };
  }

  return {
    valid: true,
    maKyHoc: found.maKyHoc,
    tenKyHoc: found.tenKyHoc
  };
}

function requireSession_(token) {
  const session = getSessionFromToken_(token);

  if (!session.valid) {
    throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
  }

  return session;
}

/* =========================================================
   QUẢN LÝ HỌC SINH
========================================================= */

function getInitialHocSinhData(token) {
  const session = requireSession_(token);
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

  return rows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() === 'ACTIVE')
    .map(row => ({
      maLop: String(row.MaLop || '').trim(),
      tenLop: String(row.TenLop || '').trim(),
      khoi: String(row.Khoi || '').trim(),
      thuTu: Number(row.ThuTu) || 999
    }))
    .sort((a, b) => {
      if (a.khoi !== b.khoi) return Number(a.khoi) - Number(b.khoi);
      return a.thuTu - b.thuTu;
    });
}

function getHocSinhList(token, filters) {
  const session = requireSession_(token);

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
        ngaySinh: formatDateForInput_(ngayVao),
        ngaySinhDisplay: formatDateDisplay_(ngayVao),
        gioiTinh: String(row.GioiTinh || '').trim(),
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

function saveHocSinhOrder(token, orderedStudentIds) {
  const session = requireSession_(token);
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

    const studentGroupMap = values.slice(1).reduce((map, row) => {
      const id = String(row[index.MaHocSinh] || '').trim();
      if (!id || !currentIdMap[id]) return map;

      const khoi = index.Khoi === undefined ? '' : String(row[index.Khoi] || '').trim();
      const lop = index.Lop === undefined ? '' : String(row[index.Lop] || '').trim();
      map[id] = {
        groupKey: khoi + '|' + lop,
        khoi: Number(khoi)
      };
      return map;
    }, {});
    const nextOrderByGroup = {};
    const orderMap = orderedIds.reduce((map, id) => {
      const group = studentGroupMap[id];

      if (!group || group.khoi < 1 || group.khoi > 9) {
        throw new Error('Học sinh có khối không hợp lệ. Chỉ hỗ trợ khối 1 đến khối 9.');
      }

      const position = number_(nextOrderByGroup[group.groupKey]);

      if (position > 99) {
        throw new Error('Mỗi lớp chỉ hỗ trợ tối đa 100 vị trí sắp xếp.');
      }

      map[id] = group.khoi * 100 + position;
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
  const session = requireSession_(token);

  hocSinh = hocSinh || {};

  const maHocSinh = String(hocSinh.maHocSinh || '').trim();
  const kyHocIds = Array.isArray(hocSinh.kyHocIds) ? hocSinh.kyHocIds : [];
  let sapXep = String(hocSinh.sapXep || '').trim();
  const khoi = String(hocSinh.khoi || '').trim();
  const lop = String(hocSinh.lop || '').trim();
  const hoTen = String(hocSinh.hoTen || '').trim();
  const ngayVao = String(hocSinh.ngaySinh || '').trim();
  const gioiTinh = String(hocSinh.gioiTinh || '').trim();
  const sdtPhuHuynh = String(hocSinh.sdtPhuHuynh || '').trim();
  const diaChi = String(hocSinh.diaChi || '').trim();
  const ghiChu = String(hocSinh.ghiChu || '').trim();
  const hocPhi = String(hocSinh.hocPhi || '').replace(/[^\d]/g, '');
  const capNhatThuPhi = toBoolean_(hocSinh.capNhatThuPhi);
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

  if (sdtPhuHuynh && !/^0\d{9}$/.test(sdtPhuHuynh)) {
    throw new Error('Số điện thoại phụ huynh phải gồm 10 số và bắt đầu bằng số 0.');
  }

  if (capNhatThuPhi && kyHocIds.indexOf(session.maKyHoc) === -1) {
    throw new Error('Muốn cập nhật thu phí tháng hiện tại, học sinh phải thuộc kỳ học đang đăng nhập.');
  }

  if (capNhatThuPhi && !/^\d{4}-\d{2}$/.test(thuPhiYearMonth)) {
    throw new Error('Tháng cập nhật thu phí không hợp lệ.');
  }

  const validLop = getLopList_().find(item => item.maLop === lop && item.khoi === khoi);

  if (!validLop) {
    throw new Error('Lớp không thuộc khối đã chọn.');
  }

  const now = new Date();
  const rows = readObjectsNoCache_(SHEET_HOCSINH);
  let newId = maHocSinh;

  const khoiNumber = Number(khoi);
  const sortRangeStart = khoiNumber * 100;
  const sortRangeEnd = sortRangeStart + 99;
  const requestedSort = number_(sapXep);
  const requestedSortUsed = rows.some(row => {
    return String(row.MaHocSinh || '').trim() !== maHocSinh &&
      String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED' &&
      String(row.Khoi || '').trim() === khoi &&
      String(row.Lop || '').trim() === lop &&
      number_(row.SapXep) === requestedSort;
  });

  if (khoiNumber < 1 || khoiNumber > 9) {
    throw new Error('Chỉ hỗ trợ sắp xếp học sinh từ khối 1 đến khối 9.');
  }

  if (requestedSort < sortRangeStart || requestedSort > sortRangeEnd || requestedSortUsed) {
    const usedSortValues = rows.reduce((map, row) => {
      const rowId = String(row.MaHocSinh || '').trim();
      const rowStatus = String(row.TrangThai || '').trim().toUpperCase();

      if (
        rowId !== maHocSinh &&
        rowStatus !== 'DELETED' &&
        String(row.Khoi || '').trim() === khoi &&
        String(row.Lop || '').trim() === lop
      ) {
        const value = number_(row.SapXep);
        if (value >= sortRangeStart && value <= sortRangeEnd) map[value] = true;
      }

      return map;
    }, {});
    let availableSort = 0;

    for (let value = sortRangeStart; value <= sortRangeEnd; value++) {
      if (!usedSortValues[value]) {
        availableSort = value;
        break;
      }
    }

    if (!availableSort) {
      throw new Error('Lớp đã sử dụng hết dải SapXep từ ' + sortRangeStart + ' đến ' + sortRangeEnd + '.');
    }

    sapXep = String(availableSort);
  }

  const existed = rows.some(row => {
    return String(row.MaHocSinh || '').trim() === maHocSinh &&
      String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
  });

  const updatedRows = rows.map(row => {
    if (
      String(row.MaHocSinh || '').trim() === maHocSinh &&
      String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED'
    ) {
      return Object.assign({}, row, {
        MaKyHoc: session.maKyHoc,
        SapXep: sapXep,
        Khoi: khoi,
        Lop: lop,
        HoTen: hoTen,
        NgayVao: parseInputDate_(ngayVao),
        NgaySinh: parseInputDate_(ngayVao),
        GioiTinh: gioiTinh,
        SDTPhuHuynh: sdtPhuHuynh,
        DiaChi: diaChi,
        GhiChu: ghiChu,
        TrangThai: 'ACTIVE',
        UpdatedAt: now
      });
    }

    return row;
  });

  if (!existed) {
    newId = 'HS_' + Utilities.getUuid().slice(0, 8).toUpperCase();

    updatedRows.push({
      MaHocSinh: newId,
      MaKyHoc: session.maKyHoc,
      SapXep: sapXep,
      Khoi: khoi,
      Lop: lop,
      HoTen: hoTen,
      NgayVao: parseInputDate_(ngayVao),
      NgaySinh: parseInputDate_(ngayVao),
      GioiTinh: gioiTinh,
      SDTPhuHuynh: sdtPhuHuynh,
      DiaChi: diaChi,
      GhiChu: ghiChu,
      TrangThai: 'ACTIVE',
      CreatedAt: now,
      UpdatedAt: now
    });
  }

  writeObjectsToSheet_(SHEET_HOCSINH, updatedRows, getHocSinhHeaders_());
  saveHocSinhKyHoc_(newId, kyHocIds, hocPhi);

  // Đổi phiên bản trước khi đọc lại dữ liệu để bảo đảm không dùng cache cũ.
  bumpDataVersion_();

  let thuPhiResult = null;

  if (!existed && capNhatThuPhi) {
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

  return jsonResponse_({
    success: true,
    maHocSinh: newId,
    thuPhiUpdated: !!(thuPhiResult && (thuPhiResult.added || thuPhiResult.alreadyExists)),
    thuPhiSheetName: thuPhiResult ? thuPhiResult.sheetName : '',
    message: message
  });
}

function deleteHocSinh(token, maHocSinh) {
  requireSession_(token);

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

  return jsonResponse_({
    success: true,
    message: 'Đã xoá học sinh và xoá bộ nhớ tạm.'
  });
}

function getHocSinhHeaders_() {
  return [
    'MaHocSinh',
    'MaKyHoc',
    'SapXep',
    'Khoi',
    'Lop',
    'HoTen',
    'NgayVao',
    'NgaySinh',
    'GioiTinh',
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
  const session = requireSession_(token);
  const ym = parseYearMonth_(yearMonth);

  ensureThuChiSheets_(session.maKyHoc);

  const cacheKey = buildCacheKey_(
    'thuphi_snapshot_v3_' + session.maKyHoc + '_' + ym.year + '_' + ym.month
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

  const rows = snapshot.rows
    .filter(row => {
      return String(row.MaKyHoc || '').trim() === session.maKyHoc &&
        String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
    });

  const resultStudents = rows
    .map(row => {
      const tamNghi = toBoolean_(row.TamNghi) ||
        String(row.TrangThaiThu || '').trim() === 'Tạm nghỉ';
      const rowHocPhi = number_(row.HocPhi);
      const hocPhiGoc = number_(row.HocPhiGoc) || rowHocPhi;
      const hocPhi = tamNghi ? 0 : rowHocPhi;
      const daThu = tamNghi ? 0 : number_(row.SoTienDaThu);
      const conLai = tamNghi ? 0 : Math.max(hocPhi - daThu, 0);
      const savedSource = String(row.NguonTienThu || '').trim().toUpperCase();
      const inferredSource = savedSource || inferNguonTienFromLegacy_(
        row.HinhThucThu,
        ''
      );

      return {
        maHocSinh: String(row.MaHocSinh || '').trim(),
        sapXep: number_(row.SapXep),

        hoTen: String(row.HoTen || '').trim(),
        khoi: String(row.Khoi || '').trim(),
        tenKhoi: String(row.TenKhoi || row.Khoi || '').trim(),
        lop: String(row.Lop || '').trim(),
        tenLop: String(row.TenLop || row.Lop || '').trim(),
        sdtPhuHuynh: String(row.SDTPhuHuynh || '').trim(),

        hocPhiGoc: hocPhiGoc,
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

/**
 * Ghi nhận hoặc chỉnh sửa thông tin thu phí của một học sinh.
 * Chỉ cập nhật đúng một dòng trong sheet tháng; không ghi lại toàn bộ sheet.
 */
function saveThuPhiHocSinh(token, data) {
  const session = requireSession_(token);

  data = data || {};

  const maHocSinh = String(data.maHocSinh || '').trim();
  const yearMonth = String(data.yearMonth || '').trim();
  const hocPhiInput = number_(data.hocPhi);
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

  if (!tamNghi && hocPhiInput < 0) {
    throw new Error('Học phí tháng không được nhỏ hơn 0.');
  }

  if (!tamNghi && soTienDaThuInput < 0) {
    throw new Error('Số tiền đã thu không được nhỏ hơn 0.');
  }

  if (!tamNghi && soTienDaThuInput > 0 && !ngayThuText) {
    throw new Error('Vui lòng nhập ngày đóng học phí.');
  }

  if (!tamNghi && soTienDaThuInput > 0 && !getNguonTienDefinition_(maNguonTien)) {
    throw new Error('Vui lòng chọn nguồn nhận tiền học phí.');
  }

  const ym = parseYearMonth_(yearMonth);
  const snapshot = ensureThuPhiMonthSnapshot_(
    session.maKyHoc,
    ym.year,
    ym.month
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

    if (tamNghi) {
      row[headerIndex.HocPhiGoc] = oldHocPhiGoc;
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
    row[headerIndex.SoPhieu] = savedSoPhieu;

    sheet.getRange(targetArrayIndex + 1, 1, 1, headers.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();

  return jsonResponse_({
    success: true,
    soPhieu: savedSoPhieu,
    message: tamNghi
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
function ensureThuPhiMonthSnapshot_(maKyHoc, year, month) {
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
    const rows = readObjectsNoCache_(sheetName);

    const hasExistingSnapshot = rows.some(row => {
      return String(row.MaKyHoc || '').trim() === maKyHoc &&
        String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED';
    });

    if (markerMatchesSheet || hasExistingSnapshot) {
      if (!markerMatchesSheet) {
        props.setProperty(
          propertyKey,
          String(sheet.getSheetId()) + '|' + String(Date.now())
        );
      }

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
        SDTPhuHuynh: student.sdtPhuHuynh,
        NgayVao: toDateOnly_(student.ngayVaoRaw) || student.ngayVaoRaw || '',
        HocPhi: hocPhi,
        HocPhiGoc: hocPhi,
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
function addHocSinhToThuPhiMonth_(maKyHoc, yearMonth, maHocSinh) {
  const ym = parseYearMonth_(yearMonth);
  const snapshot = ensureThuPhiMonthSnapshot_(maKyHoc, ym.year, ym.month);
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
      SDTPhuHuynh: student.sdtPhuHuynh,
      NgayVao: toDateOnly_(student.ngayVaoRaw) || student.ngayVaoRaw || '',
      HocPhi: hocPhi,
      HocPhiGoc: hocPhi,
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
    'SDTPhuHuynh',
    'NgayVao',
    'HocPhi',
    'HocPhiGoc',
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
    'GhiChu',
    'NguonDuLieu',
    'MaThamChieu',
    'TrangThai',
    'NguoiTao',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function getNguonTienHeaders_() {
  return [
    'MaKyHoc',
    'MaNguonTien',
    'TenNguonTien',
    'ThuTu',
    'SoDuBanDau',
    'TrangThai',
    'CreatedAt',
    'UpdatedAt'
  ];
}

function getDefaultNguonTien_() {
  return [
    ['VIETCOMBANK', 'Vietcombank', 1],
    ['VIETINBANK', 'VietinBank', 2],
    ['BIDV', 'BIDV', 3],
    ['TIEN_MAT', 'Tiền mặt', 4],
    ['KET', 'Két', 5]
  ];
}

function getNguonTienDefinition_(maNguonTien) {
  const code = String(maNguonTien || '').trim().toUpperCase();
  const found = getDefaultNguonTien_().find(item => item[0] === code);

  return found
    ? { maNguonTien: found[0], tenNguonTien: found[1], thuTu: found[2] }
    : null;
}

function getNguonTienName_(maNguonTien) {
  const found = getNguonTienDefinition_(maNguonTien);
  return found ? found.tenNguonTien : '';
}

function getHinhThucByNguon_(maNguonTien) {
  const code = String(maNguonTien || '').trim().toUpperCase();
  return code === 'VIETCOMBANK' || code === 'VIETINBANK' || code === 'BIDV'
    ? 'Chuyển khoản'
    : (code === 'TIEN_MAT' || code === 'KET' ? 'Tiền mặt' : '');
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

function getNguonTienList_(maKyHoc) {
  const rows = readObjectsNoCache_(SHEET_NGUONTIEN);

  return rows
    .filter(row => {
      return String(row.MaKyHoc || '').trim() === String(maKyHoc || '').trim() &&
        String(row.TrangThai || 'ACTIVE').trim().toUpperCase() !== 'DELETED';
    })
    .map(row => ({
      maNguonTien: String(row.MaNguonTien || '').trim().toUpperCase(),
      tenNguonTien: String(row.TenNguonTien || '').trim(),
      thuTu: number_(row.ThuTu) || 999,
      soDuBanDau: number_(row.SoDuBanDau),
      trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase()
    }))
    .filter(item => !!getNguonTienDefinition_(item.maNguonTien))
    .sort((a, b) => a.thuTu - b.thuTu);
}


function getDefaultDanhMucThuChi_() {
  return [
    ['THU_HOC_PHI', 'THU', 'Thu học phí', 1],
    ['THU_GHI_DANH', 'THU', 'Phí ghi danh', 2],
    ['THU_BAN_HANG', 'THU', 'Bán hàng, giáo trình', 3],
    ['THU_TAI_TRO', 'THU', 'Tài trợ', 4],
    ['THU_HOAN_PHI', 'THU', 'Hoàn phí', 5],
    ['THU_KHAC', 'THU', 'Thu khác', 99],
    ['CHI_LUONG', 'CHI', 'Lương giáo viên, nhân viên', 1],
    ['CHI_THUE', 'CHI', 'Thuê mặt bằng', 2],
    ['CHI_DIEN_NUOC', 'CHI', 'Điện, nước, Internet', 3],
    ['CHI_HOC_CU', 'CHI', 'Học cụ, tài liệu', 4],
    ['CHI_MARKETING', 'CHI', 'Marketing, quảng cáo', 5],
    ['CHI_SUA_CHUA', 'CHI', 'Sửa chữa, bảo trì', 6],
    ['CHI_KHAC', 'CHI', 'Chi khác', 99]
  ];
}

function ensureThuChiSheets_(maKyHoc) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const schemaVersion = '4';
  const schemaKey = 'THUCHI_SCHEMA_VERSION';
  const sourceKey = maKyHoc
    ? 'THUCHI_NGUON_INIT_' + String(maKyHoc || '').trim()
    : '';

  const schemaReady =
    props.getProperty(schemaKey) === schemaVersion &&
    ss.getSheetByName(SHEET_DANHMUC_THUCHI) &&
    ss.getSheetByName(SHEET_SOTHUCHI) &&
    ss.getSheetByName(SHEET_NGUONTIEN);

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

    props.setProperty(schemaKey, schemaVersion);
  } finally {
    lock.releaseLock();
  }
}

function getQuanLyThuChiData(token, filters) {
  const session = requireSession_(token);
  ensureThuChiSheets_(session.maKyHoc);

  filters = filters || {};

  const range = resolveThuChiDateRange_(filters.fromDate, filters.toDate);
  const loaiFilter = String(filters.loai || 'ALL').trim().toUpperCase();
  const danhMucFilter = String(filters.maDanhMuc || '').trim();
  const sourceFilter = String(filters.maNguonTien || 'ALL').trim().toUpperCase();
  const trangThaiFilter = String(filters.trangThai || 'HOAT_DONG').trim().toUpperCase();
  const keyword = normalizeText_(filters.keyword || '');

  const cacheKey = buildCacheKey_(
    'thuchi_v2_' + session.maKyHoc + '_' +
    range.fromText + '_' + range.toText + '_' +
    loaiFilter + '_' + danhMucFilter + '_' +
    sourceFilter + '_' + trangThaiFilter + '_' + keyword
  );

  const cached = cacheGetString_(cacheKey);
  if (cached) return cached;

  const categoryRows = readObjects_(SHEET_DANHMUC_THUCHI);
  const transactionRows = readObjects_(SHEET_SOTHUCHI);
  const sources = getNguonTienList_(session.maKyHoc);

  const categories = categoryRows
    .filter(row => String(row.TrangThai || '').trim().toUpperCase() !== 'DELETED')
    .map(row => ({
      maDanhMuc: String(row.MaDanhMuc || '').trim(),
      loai: String(row.Loai || '').trim().toUpperCase(),
      tenDanhMuc: String(row.TenDanhMuc || '').trim(),
      thuTu: number_(row.ThuTu) || 999,
      trangThai: String(row.TrangThai || 'ACTIVE').trim().toUpperCase()
    }))
    .filter(item => item.maDanhMuc && (item.loai === 'THU' || item.loai === 'CHI'))
    .sort((a, b) => {
      if (a.loai !== b.loai) return a.loai.localeCompare(b.loai);
      if (a.thuTu !== b.thuTu) return a.thuTu - b.thuTu;
      return a.tenDanhMuc.localeCompare(b.tenDanhMuc, 'vi');
    });

  const allSessionTransactions = transactionRows
    .filter(row => String(row.MaKyHoc || '').trim() === session.maKyHoc)
    .map(mapThuChiTransaction_)
    .filter(item => item.maGiaoDich && item.ngayDate);

  const activeAll = allSessionTransactions.filter(item => item.trangThai === 'HOAT_DONG');
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

    const operational = inPeriod.filter(item => item.nguonDuLieu !== 'CHUYEN_NOI_BO');
    const transfers = inPeriod.filter(item => item.nguonDuLieu === 'CHUYEN_NOI_BO');

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

  let dateSourceTransactions = allSessionTransactions.filter(item => {
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
  const activeOperational = activeFiltered.filter(item => item.nguonDuLieu !== 'CHUYEN_NOI_BO');

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
  const yearTransactions = allSessionTransactions.filter(item => {
    if (!item.ngayDate || item.ngayDate.getFullYear() !== reportYear) return false;
    if (sourceFilter === 'UNASSIGNED') return !item.maNguonTien;
    if (sourceFilter !== 'ALL') return item.maNguonTien === sourceFilter;
    return true;
  });

  const monthly = buildThuChiMonthlyReport_(
    yearTransactions.filter(item => item.nguonDuLieu !== 'CHUYEN_NOI_BO'),
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
    categories: categories,
    sources: sources,
    sourceSummaries: sourceSummaries,
    unassignedSummary: hasUnassigned ? unassignedSummary : null,
    transactions: sortedDesc.map(stripThuChiPrivateFields_),
    recentTransactions: sortedDesc.slice(0, 10).map(stripThuChiPrivateFields_),
    cashbook: cashbook,
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

function saveThuChiGiaoDich(token, data) {
  const session = requireSession_(token);
  ensureThuChiSheets_(session.maKyHoc);

  data = data || {};

  const maGiaoDich = String(data.maGiaoDich || '').trim();
  const loai = String(data.loai || '').trim().toUpperCase();
  const maDanhMuc = String(data.maDanhMuc || '').trim();
  const maNguonTien = String(data.maNguonTien || '').trim().toUpperCase();
  const ngayText = String(data.ngayGiaoDich || '').trim();
  const noiDung = String(data.noiDung || '').trim();
  const soTien = number_(data.soTien);
  const nguoiNopNhan = String(data.nguoiNopNhan || '').trim();
  const soChungTu = String(data.soChungTu || '').trim();
  const ghiChu = String(data.ghiChu || '').trim();

  if (loai !== 'THU' && loai !== 'CHI') {
    throw new Error('Loại giao dịch không hợp lệ.');
  }

  if (!ngayText || !toDateOnly_(ngayText)) {
    throw new Error('Vui lòng nhập ngày giao dịch hợp lệ.');
  }

  if (!maDanhMuc) {
    throw new Error('Vui lòng chọn danh mục thu chi.');
  }

  if (!maNguonTien || !getNguonTienDefinition_(maNguonTien)) {
    throw new Error('Vui lòng chọn đúng nguồn tiền.');
  }

  if (!noiDung) {
    throw new Error('Vui lòng nhập nội dung giao dịch.');
  }

  if (soTien <= 0) {
    throw new Error('Số tiền phải lớn hơn 0.');
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

  const sourceList = getNguonTienList_(session.maKyHoc);
  const source = sourceList.find(item => item.maNguonTien === maNguonTien);

  if (!source || source.trangThai !== 'ACTIVE') {
    throw new Error('Nguồn tiền không tồn tại hoặc đã ngừng sử dụng.');
  }

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
    currentRow[index.GhiChu] = ghiChu;
    currentRow[index.NguonDuLieu] = 'NHAP_TAY';
    currentRow[index.MaThamChieu] = '';
    currentRow[index.TrangThai] = 'HOAT_DONG';
    currentRow[index.NguoiTao] = 'Admin';
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

  return jsonResponse_({
    success: true,
    maGiaoDich: savedId,
    soPhieu: savedSoPhieu,
    message: maGiaoDich
      ? 'Đã cập nhật ' + (loai === 'CHI' ? 'phiếu chi ' : 'phiếu thu ') + savedSoPhieu + '.'
      : 'Đã tạo ' + (loai === 'CHI' ? 'phiếu chi ' : 'phiếu thu ') + savedSoPhieu + '.'
  });
}

function saveChuyenNguonTien(token, data) {
  const session = requireSession_(token);
  ensureThuChiSheets_(session.maKyHoc);

  data = data || {};

  const ngayText = String(data.ngayGiaoDich || '').trim();
  const maNguonDi = String(data.maNguonDi || '').trim().toUpperCase();
  const maNguonDen = String(data.maNguonDen || '').trim().toUpperCase();
  const soTien = number_(data.soTien);
  const noiDung = String(data.noiDung || '').trim() || 'Chuyển tiền nội bộ';
  const soChungTu = String(data.soChungTu || '').trim();
  const ghiChu = String(data.ghiChu || '').trim();

  if (!ngayText || !toDateOnly_(ngayText)) {
    throw new Error('Vui lòng nhập ngày chuyển tiền hợp lệ.');
  }

  if (!maNguonDi || !maNguonDen || maNguonDi === maNguonDen) {
    throw new Error('Nguồn chuyển và nguồn nhận phải khác nhau.');
  }

  if (soTien <= 0) {
    throw new Error('Số tiền chuyển phải lớn hơn 0.');
  }

  const sourceList = getNguonTienList_(session.maKyHoc);
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
    outRow[index.NguonDuLieu] = 'CHUYEN_NOI_BO';
    outRow[index.MaThamChieu] = groupId;
    outRow[index.TrangThai] = 'HOAT_DONG';
    outRow[index.NguoiTao] = 'Admin';
    outRow[index.CreatedAt] = now;
    outRow[index.UpdatedAt] = now;

    const inRow = new Array(headers.length).fill('');
    inRow[index.MaGiaoDich] = 'TC_VAO_' + Utilities.getUuid().slice(0, 10).toUpperCase();
    inRow[index.NgayGiaoDich] = parseInputDate_(ngayText);
    inRow[index.MaKyHoc] = session.maKyHoc;
    inRow[index.LoaiGiaoDich] = 'THU';
    inRow[index.MaDanhMuc] = 'THU_CHUYEN_NOI_BO';
    inRow[index.TenDanhMuc] = 'Nhận chuyển tiền nội bộ';
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
    inRow[index.NguonDuLieu] = 'CHUYEN_NOI_BO';
    inRow[index.MaThamChieu] = groupId;
    inRow[index.TrangThai] = 'HOAT_DONG';
    inRow[index.NguoiTao] = 'Admin';
    inRow[index.CreatedAt] = now;
    inRow[index.UpdatedAt] = now;

    sheet
      .getRange(sheet.getLastRow() + 1, 1, 2, headers.length)
      .setValues([outRow, inRow]);
  } finally {
    lock.releaseLock();
  }

  bumpDataVersion_();

  return jsonResponse_({
    success: true,
    message: 'Đã chuyển tiền giữa hai nguồn. Giao dịch không làm tăng tổng thu hoặc tổng chi.'
  });
}

function cancelChuyenNguonTien(token, maNhomChuyen) {
  const session = requireSession_(token);
  ensureThuChiSheets_(session.maKyHoc);

  const groupId = String(maNhomChuyen || '').trim();
  if (!groupId) throw new Error('Thiếu mã nhóm chuyển tiền.');

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

  return jsonResponse_({
    success: true,
    message: 'Đã huỷ giao dịch chuyển tiền nội bộ.'
  });
}

function saveNguonTienConfig(token, data) {
  const session = requireSession_(token);
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

  return jsonResponse_({
    success: true,
    message: 'Đã cập nhật số dư ban đầu của ' + getNguonTienName_(maNguonTien) + '.'
  });
}


function cancelThuChiGiaoDich(token, maGiaoDich) {
  const session = requireSession_(token);
  ensureThuChiSheets_();

  const id = String(maGiaoDich || '').trim();
  if (!id) throw new Error('Thiếu mã giao dịch.');

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

  return jsonResponse_({
    success: true,
    message: 'Đã huỷ giao dịch. Dữ liệu vẫn được giữ để đối chiếu.'
  });
}

function saveDanhMucThuChi(token, data) {
  requireSession_(token);
  ensureThuChiSheets_();

  data = data || {};

  const inputId = String(data.maDanhMuc || '').trim();
  const loai = String(data.loai || '').trim().toUpperCase();
  const tenDanhMuc = String(data.tenDanhMuc || '').trim();
  const thuTu = number_(data.thuTu) || 999;

  if (loai !== 'THU' && loai !== 'CHI') throw new Error('Loại danh mục không hợp lệ.');
  if (!tenDanhMuc) throw new Error('Vui lòng nhập tên danh mục.');

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

  return jsonResponse_({
    success: true,
    message: inputId ? 'Đã cập nhật danh mục.' : 'Đã thêm danh mục.'
  });
}

function toggleDanhMucThuChi(token, maDanhMuc, enabled) {
  requireSession_(token);
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

  return jsonResponse_({
    success: true,
    message: toBoolean_(enabled) ? 'Đã kích hoạt danh mục.' : 'Đã ngừng sử dụng danh mục.'
  });
}

function syncHocPhiToThuChi(token, yearMonth) {
  const session = requireSession_(token);
  ensureThuChiSheets_(session.maKyHoc);

  const ym = parseYearMonth_(yearMonth);
  const monthKey = ym.year + '-' + String(ym.month).padStart(2, '0');
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

  return {
    maGiaoDich: String(row.MaGiaoDich || '').trim(),
    ngayDate: date,
    ngayGiaoDich: formatDateForInput_(date),
    ngayDisplay: formatDateDisplay_(date),
    loai: String(row.LoaiGiaoDich || '').trim().toUpperCase(),
    maDanhMuc: String(row.MaDanhMuc || '').trim(),
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
    ghiChu: String(row.GhiChu || '').trim(),
    nguonDuLieu: source,
    maThamChieu: String(row.MaThamChieu || '').trim(),
    trangThai: String(row.TrangThai || 'HOAT_DONG').trim().toUpperCase(),
    nguoiTao: String(row.NguoiTao || '').trim(),
    createdAtRaw: row.CreatedAt || '',
    updatedAtRaw: row.UpdatedAt || '',
    editable: source === 'NHAP_TAY' &&
      String(row.TrangThai || 'HOAT_DONG').trim().toUpperCase() === 'HOAT_DONG',
    transfer: source === 'CHUYEN_NOI_BO',
    canCancelTransfer: source === 'CHUYEN_NOI_BO' &&
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
    ghiChu: item.ghiChu,
    nguonDuLieu: item.nguonDuLieu,
    maThamChieu: item.maThamChieu,
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

  sheet.clearContents();
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  formatHeader_(sheet);
  sheet.autoResizeColumns(1, headers.length);
}

function compareStudentSort_(a, b) {
  const sortA = number_(a.sapXep || a.SapXep);
  const sortB = number_(b.sapXep || b.SapXep);
  const finalA = sortA > 0 ? sortA : 999999;
  const finalB = sortB > 0 ? sortB : 999999;

  // Mọi danh sách học sinh ưu tiên SapXep tăng dần: 100-199, 200-299, ... 900-999.
  if (finalA !== finalB) return finalA - finalB;

  const khoiA = Number(a.khoi || a.Khoi || 999);
  const khoiB = Number(b.khoi || b.Khoi || 999);

  if (khoiA !== khoiB) return khoiA - khoiB;

  const lopA = String(a.lop || a.Lop || '');
  const lopB = String(b.lop || b.Lop || '');

  if (lopA !== lopB) return lopA.localeCompare(lopB, 'vi');

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
  const props = PropertiesService.getScriptProperties();
  let version = props.getProperty('DATA_VERSION');

  if (!version) {
    version = String(Date.now());
    props.setProperty('DATA_VERSION', version);
  }

  return version;
}

function bumpDataVersion_() {
  PropertiesService.getScriptProperties().setProperty('DATA_VERSION', String(Date.now()));
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


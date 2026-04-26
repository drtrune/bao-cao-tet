/**
 * 🛠️ CONFIG BACKEND (Google Apps Script)
 * ----------------------------------------------------------------------
 * FILE: Config_Backend.gs
 * MỤC ĐÍCH: 
 *   1. Đọc dữ liệu cấu hình từ các Sheet (Config_*)
 *   2. Cung cấp hàm publish để đẩy cấu hình lên GitHub (tạo ra file config.js)
 *   3. Quản trị các tham số hệ thống (PAT, Spreadsheet ID)
 * 
 * LƯU Ý: Đây là "nguồn phát". Mọi thay đổi về logic cấu hình thực hiện tại đây.
 * ----------------------------------------------------------------------
 */

// ======================================================================
// QUẢN TRỊ SPREADSHEET & QUYỀN
// ======================================================================

/**
 * HÀM NÀY ĐỂ KÍCH HOẠT QUYỀN DRIVE (Chạy thủ công 1 lần nếu gặp lỗi)
 */
function FORCE_AUTH() {
  DriveApp.getRootFolder();
  console.log("Đã cấp quyền Drive thành công!");
}

/**
 * Lấy Spreadsheet từ ID trong Script Properties
 * Nếu chưa có ID, fallback về getActiveSpreadsheet()
 */
function getSpreadsheet() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {
      console.warn("Không mở được Spreadsheet bằng ID từ Properties: " + e.message);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Lưu Spreadsheet ID vào Script Properties
 */
function setupSpreadsheetId() {
  const ui = SpreadsheetApp.getUi();
  const currentId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const autoId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const masked = currentId ? ('...' + currentId.slice(-6)) : '(chưa cài đặt)';

  const response = ui.prompt(
    '📎 Cài đặt Spreadsheet ID',
    'ID hiện tại: ' + masked + '\n' +
    'ID của Sheet này: ' + autoId + '\n\n' +
    'Nhấn OK để dùng ID của Sheet này, hoặc nhập ID khác:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const inputId = response.getResponseText().trim();
    const finalId = inputId || autoId;
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', finalId);
    ui.alert('✅ Đã lưu Spreadsheet ID!\n\nID: ' + finalId);
  }
}

// ======================================================================
// MENU TUỲ CHỈNH TRÊN GOOGLE SHEET
// ======================================================================

/**
 * Tạo custom menu khi mở spreadsheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Config')
    .addItem('🚀 Publish Config lên GitHub', 'publishConfigToGitHub')
    .addSeparator()
    .addItem('🔑 Cài đặt GitHub PAT', 'setupGitHubPAT')
    .addItem('📎 Cài đặt Spreadsheet ID', 'setupSpreadsheetId')
    .addItem('📋 Kiểm tra cấu hình', 'showCurrentConfig')
    .addToUi();
}

// ======================================================================
// ĐỌC CẤU HÌNH TỪ SHEET (Dành cho Backend)
// ======================================================================

/**
 * Đọc danh sách đối tượng thống kê từ sheet Config_DoiTuong
 */
function getConfigDoiTuong() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Config_DoiTuong');
  if (!sheet || sheet.getLastRow() < 2) {
    return ["KHÁM CHỮA BỆNH CHUNG", "Tai nạn giao thông", "COVID-19", "Các đối tượng người bệnh khác (không gồm các đối tượng trên)"];
  }
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var ten = String(data[i][1] || '').trim();
    if (ten) result.push(ten);
  }
  return result.length > 0 ? result : ["KHÁM CHỮA BỆNH CHUNG", "Tai nạn giao thông", "COVID-19", "Các đối tượng người bệnh khác"];
}

/**
 * Đọc danh sách tất cả khoa phòng từ sheet Config_KhoaPhong
 */
function getConfigKhoaPhong() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Config_KhoaPhong');
  if (!sheet || sheet.getLastRow() < 2) {
    return ["Nội tổng hợp", "Ngoại tổng hợp", "Phụ Sản", "Nhi", "Liên chuyên khoa", "Khám bệnh", "Phòng cấp cứu", "Hồi sức Cấp cứu"];
  }
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var ten = String(data[i][1] || '').trim();
    if (ten) result.push(ten);
  }
  return result.length > 0 ? result : ["Nội tổng hợp", "Ngoại tổng hợp", "Phụ Sản", "Nhi", "Liên chuyên khoa", "Khám bệnh", "Phòng cấp cứu", "Hồi sức Cấp cứu"];
}

// ======================================================================
// GITHUB PUBLISHER LOGIC
// ======================================================================

/**
 * Mở dialog để user nhập GitHub Personal Access Token
 */
function setupGitHubPAT() {
  const ui = SpreadsheetApp.getUi();
  const currentPAT = PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');
  const masked = currentPAT ? ('***' + currentPAT.slice(-4)) : '(chưa cài đặt)';

  const response = ui.prompt(
    '🔑 Cài đặt GitHub Personal Access Token',
    'Token hiện tại: ' + masked + '\n\n' +
    'Nhập GitHub PAT mới (scope: repo).\n' +
    'Tạo tại: https://github.com/settings/tokens',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const token = response.getResponseText().trim();
    if (token.length < 10) {
      ui.alert('❌ Token không hợp lệ (quá ngắn).');
      return;
    }
    PropertiesService.getScriptProperties().setProperty('GITHUB_PAT', token);
    ui.alert('✅ Đã lưu GitHub PAT thành công!\nMasked: ***' + token.slice(-4));
  }
}

/**
 * Đọc dữ liệu từ 1 sheet, trả về mảng object (header = key)
 */
function readSheetAsObjects(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h).trim(); });
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    let hasValue = false;
    headers.forEach(function(key, idx) {
      if (key) {
        row[key] = data[i][idx];
        if (data[i][idx] !== '' && data[i][idx] !== null) hasValue = true;
      }
    });
    if (hasValue) rows.push(row);
  }
  return rows;
}

/**
 * Đọc sheet key-value (2 cột: Key, Value)
 */
function readSheetAsKeyValue(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (key) result[key] = data[i][1];
  }
  return result;
}

/**
 * Đọc toàn bộ config từ các sheet Config_* và build object APP_CONFIG
 */
function readAllConfig() {
  const khoaPhongRaw = readSheetAsObjects('Config_KhoaPhong');
  const khoaPhong = khoaPhongRaw.map(function(r) {
    return {
      ma: String(r['Mã khoa'] || '').trim(),
      ten: String(r['Tên khoa'] || '').trim(),
      nhom: String(r['Nhóm'] || '').trim(),
      sections: String(r['Sections hiển thị'] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean),
      disableCols: String(r['Cột disable'] || '').split(',').map(function(s) { return parseInt(s.trim()); }).filter(function(n) { return !isNaN(n); })
    };
  });

  const doiTuongRaw = readSheetAsObjects('Config_DoiTuong');
  const doiTuong = doiTuongRaw.map(function(r) {
    return {
      stt: parseInt(r['STT']) || 0,
      ten: String(r['Tên đối tượng'] || '').trim(),
      laDongTong: String(r['Là dòng tổng']).toUpperCase() === 'TRUE'
    };
  });

  const bangTKRaw = readSheetAsObjects('Config_BangThongKe');
  const colCodes = [];
  const headers = [];
  const computedCols = {};

  bangTKRaw.forEach(function(r) {
    const maCot = String(r['Mã cột'] || '').trim();
    const code = parseInt(String(r['Code nhập'] || '0').trim());
    const congThuc = String(r['Công thức'] || '').trim();
    const kieuHT = String(r['Kiểu hiển thị'] || 'input').trim();

    if (code && kieuHT === 'input') colCodes.push(code);
    if (congThuc) computedCols[maCot] = congThuc;

    headers.push({
      maCot: maCot,
      header1: String(r['Header cấp 1'] || '').trim(),
      header2: String(r['Header cấp 2'] || '').trim(),
      colspan: parseInt(r['colspan']) || 0,
      rowspan: parseInt(r['rowspan']) || 0,
      code: code,
      kieuHT: kieuHT
    });
  });

  const chiSoRaw = readSheetAsObjects('Config_ChiSoKhac');
  const chiSoKhac = chiSoRaw.map(function(r) {
    return {
      ma: String(r['Mã'] || '').trim(),
      ten: String(r['Tên chỉ số'] || '').trim(),
      inputId: String(r['ID input'] || '').trim(),
      section: String(r['Section'] || '').trim(),
      icon: String(r['Icon'] || '').trim(),
      mauNen: String(r['Màu nền'] || '').trim(),
      donVi: String(r['Đơn vị'] || '').trim()
    };
  });

  const danhMucRaw = readSheetAsObjects('Config_DanhMuc');
  const danhMuc = {};
  danhMucRaw.forEach(function(r) {
    const nhom = String(r['Nhóm'] || '').trim();
    const ma = String(r['Mã'] || '').trim();
    if (!nhom || !ma) return;
    if (!danhMuc[nhom]) danhMuc[nhom] = {};
    danhMuc[nhom][ma] = {
      label: String(r['Nhãn hiển thị'] || '').trim(),
      short: String(r['Nhãn ngắn'] || '').trim()
    };
  });

  const app = readSheetAsKeyValue('Config_App');

  return {
    version: new Date().toISOString(),
    khoaPhong: khoaPhong,
    doiTuong: doiTuong,
    bangThongKe: { colCodes: colCodes, headers: headers, computedCols: computedCols },
    chiSoKhac: chiSoKhac,
    danhMuc: danhMuc,
    app: app
  };
}

/**
 * Đọc config → build file config.js → push lên GitHub Pages
 */
function publishConfigToGitHub() {
  const ui = SpreadsheetApp.getUi();
  const pat = PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');
  if (!pat) {
    ui.alert('❌ Chưa cài đặt GitHub PAT!\n\nVui lòng vào menu ⚙️ Config → 🔑 Cài đặt GitHub PAT');
    return;
  }

  const appConfig = readSheetAsKeyValue('Config_App');
  const repo = String(appConfig['GITHUB_REPO'] || '').trim();
  const configPath = String(appConfig['GITHUB_CONFIG_PATH'] || 'config.js').trim();
  const branch = String(appConfig['GITHUB_BRANCH'] || 'main').trim();

  if (!repo) {
    ui.alert('❌ Thiếu GITHUB_REPO trong sheet Config_App!');
    return;
  }

  try {
    const config = readAllConfig();
    const timestamp = Utilities.formatDate(new Date(), 'GMT+7', "yyyy-MM-dd'T'HH:mm:ssXXX");
    const jsContent = '/**\n' +
      ' * AUTO-GENERATED CONFIG — Không chỉnh sửa trực tiếp\n' +
      ' * Cập nhật lúc: ' + timestamp + '\n' +
      ' */\n' +
      'window.APP_CONFIG = ' + JSON.stringify(config, null, 2) + ';\n';

    const currentSHA = getGitHubFileSHA(pat, repo, configPath, branch);
    var payload = {
      message: '🔄 Cập nhật config - ' + timestamp,
      content: Utilities.base64Encode(Utilities.newBlob(jsContent).getBytes()),
      branch: branch
    };
    if (currentSHA) payload.sha = currentSHA;

    var apiUrl = 'https://api.github.com/repos/' + repo + '/contents/' + configPath;
    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'PUT',
      headers: {
        'Authorization': 'token ' + pat,
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
      ui.alert('✅ Publish thành công!\n\nGitHub Pages sẽ cập nhật sau vài phút.');
    } else {
      ui.alert('❌ Lỗi GitHub API: ' + response.getContentText());
    }
  } catch (err) {
    ui.alert('❌ Lỗi: ' + err.toString());
  }
}

/**
 * Lấy SHA hiện tại của file trên GitHub
 */
function getGitHubFileSHA(pat, repo, filePath, branch) {
  try {
    var apiUrl = 'https://api.github.com/repos/' + repo + '/contents/' + filePath + '?ref=' + branch;
    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'GET',
      headers: { 'Authorization': 'token ' + pat, 'Accept': 'application/vnd.github.v3+json' },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) return JSON.parse(response.getContentText()).sha;
    return null;
  } catch (e) { return null; }
}

/**
 * Xem trước config hiện tại (debug)
 */
function showCurrentConfig() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = readAllConfig();
    const summary = '📊 Config hiện tại:\n\n' +
      '🏥 Khoa phòng: ' + config.khoaPhong.length + '\n' +
      '📋 Đối tượng: ' + config.doiTuong.length + '\n' +
      '⚙️ Version: ' + config.version;
    ui.alert(summary);
  } catch (err) { ui.alert('❌ Lỗi: ' + err.toString()); }
}

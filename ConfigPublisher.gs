/**
 * Config Publisher — Đọc config từ Sheet và publish lên GitHub Pages
 * 
 * Workflow: Sheet Config_* → readAllConfig() → publishConfigToGitHub() → config.js trên GitHub Pages
 */

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
    .addItem('📋 Kiểm tra cấu hình', 'showCurrentConfig')
    .addToUi();
}

// ======================================================================
// CÀI ĐẶT GITHUB PAT
// ======================================================================

/**
 * Mở dialog để user nhập GitHub Personal Access Token
 * Token được lưu an toàn trong Script Properties (không lưu trong Sheet)
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

// ======================================================================
// ĐỌC TOÀN BỘ CONFIG TỪ SHEETS
// ======================================================================

/**
 * Đọc dữ liệu từ 1 sheet, trả về mảng object (header = key)
 */
function readSheetAsObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  // 1. Khoa phòng
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

  // 2. Đối tượng (dòng trong bảng thống kê)
  const doiTuongRaw = readSheetAsObjects('Config_DoiTuong');
  const doiTuong = doiTuongRaw.map(function(r) {
    return {
      stt: parseInt(r['STT']) || 0,
      ten: String(r['Tên đối tượng'] || '').trim(),
      laDongTong: String(r['Là dòng tổng']).toUpperCase() === 'TRUE'
    };
  });

  // 3. Bảng thống kê (header 2 cấp + cột)
  const bangTKRaw = readSheetAsObjects('Config_BangThongKe');
  const colCodes = [];
  const headers = [];
  const computedCols = {};

  bangTKRaw.forEach(function(r) {
    const maCot = String(r['Mã cột'] || '').trim();
    const code = parseInt(String(r['Code nhập'] || '0').trim());
    const congThuc = String(r['Công thức'] || '').trim();
    const kieuHT = String(r['Kiểu hiển thị'] || 'input').trim();

    if (code && kieuHT === 'input') {
      colCodes.push(code);
    }

    if (congThuc) {
      computedCols[maCot] = congThuc;
    }

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

  // 4. Chỉ số khác
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

  // 5. Danh mục (nhóm theo cột Nhóm)
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

  // 6. Cấu hình ứng dụng (key-value)
  const app = readSheetAsKeyValue('Config_App');

  return {
    version: new Date().toISOString(),
    khoaPhong: khoaPhong,
    doiTuong: doiTuong,
    bangThongKe: {
      colCodes: colCodes,
      headers: headers,
      computedCols: computedCols
    },
    chiSoKhac: chiSoKhac,
    danhMuc: danhMuc,
    app: app
  };
}

// ======================================================================
// PUBLISH LÊN GITHUB
// ======================================================================

/**
 * Đọc config → build file config.js → push lên GitHub Pages
 */
function publishConfigToGitHub() {
  const ui = SpreadsheetApp.getUi();

  // Kiểm tra PAT
  const pat = PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');
  if (!pat) {
    ui.alert('❌ Chưa cài đặt GitHub PAT!\n\nVui lòng vào menu ⚙️ Config → 🔑 Cài đặt GitHub PAT');
    return;
  }

  // Đọc thông tin repo từ Config_App
  const appConfig = readSheetAsKeyValue('Config_App');
  const repo = String(appConfig['GITHUB_REPO'] || '').trim();
  const configPath = String(appConfig['GITHUB_CONFIG_PATH'] || 'config.js').trim();
  const branch = String(appConfig['GITHUB_BRANCH'] || 'main').trim();

  if (!repo) {
    ui.alert('❌ Thiếu GITHUB_REPO trong sheet Config_App!\n\nVí dụ: drtrune/bao-cao-tet');
    return;
  }

  try {
    // Đọc toàn bộ config
    const config = readAllConfig();

    // Build nội dung file config.js
    const timestamp = Utilities.formatDate(new Date(), 'GMT+7', "yyyy-MM-dd'T'HH:mm:ssXXX");
    const jsContent = '/**\n' +
      ' * AUTO-GENERATED CONFIG — Không chỉnh sửa trực tiếp\n' +
      ' * Nguồn: Google Sheet Config → GAS publishConfig → GitHub API\n' +
      ' * Cập nhật lúc: ' + timestamp + '\n' +
      ' */\n' +
      'window.APP_CONFIG = ' + JSON.stringify(config, null, 2) + ';\n';

    // Lấy SHA của file hiện tại (nếu đã tồn tại, cần SHA để update)
    const currentSHA = getGitHubFileSHA(pat, repo, configPath, branch);

    // Push lên GitHub
    var payload = {
      message: '🔄 Cập nhật config - ' + timestamp,
      content: Utilities.base64Encode(Utilities.newBlob(jsContent).getBytes()),
      branch: branch
    };
    if (currentSHA) {
      payload.sha = currentSHA;
    }

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

    var responseCode = response.getResponseCode();
    if (responseCode === 200 || responseCode === 201) {
      var result = JSON.parse(response.getContentText());
      var htmlUrl = result.content ? result.content.html_url : '';
      ui.alert(
        '✅ Publish thành công!\n\n' +
        '📁 File: ' + configPath + '\n' +
        '🕐 Thời gian: ' + timestamp + '\n' +
        '🔗 URL: ' + htmlUrl + '\n\n' +
        '⏳ GitHub Pages sẽ cập nhật trong 1-10 phút.'
      );
    } else {
      var errorBody = response.getContentText();
      ui.alert('❌ Lỗi GitHub API!\n\nHTTP ' + responseCode + ':\n' + errorBody);
    }
  } catch (err) {
    ui.alert('❌ Lỗi: ' + err.toString());
  }
}

/**
 * Lấy SHA hiện tại của file trên GitHub (cần cho lệnh update)
 */
function getGitHubFileSHA(pat, repo, filePath, branch) {
  try {
    var apiUrl = 'https://api.github.com/repos/' + repo + '/contents/' + filePath + '?ref=' + branch;
    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Authorization': 'token ' + pat,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText()).sha;
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Xem trước config hiện tại (debug)
 */
function showCurrentConfig() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = readAllConfig();
    const summary = 
      '📊 Config hiện tại:\n\n' +
      '🏥 Khoa phòng: ' + config.khoaPhong.length + ' khoa\n' +
      '📋 Đối tượng: ' + config.doiTuong.length + ' dòng\n' +
      '📊 Cột bảng: ' + config.bangThongKe.colCodes.length + ' cột\n' +
      '📈 Chỉ số khác: ' + config.chiSoKhac.length + ' chỉ số\n' +
      '📁 Danh mục: ' + Object.keys(config.danhMuc).join(', ') + '\n' +
      '⚙️ App settings: ' + Object.keys(config.app).length + ' keys\n\n' +
      '🕐 Version: ' + config.version;

    ui.alert(summary);
  } catch (err) {
    ui.alert('❌ Lỗi đọc config: ' + err.toString());
  }
}

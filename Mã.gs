/**
 * Backend Google Apps Script - Lễ 30/4-1/5
 * Chứa logic nghiệp vụ xử lý dữ liệu và API Router
 */

/**
 * Xử lý giao diện Web App
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile("index");
  // Truyền URL WebApp xuống để frontend gọi fetch
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle("Báo cáo Tết 2026 - Bệnh viện Bình Định")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Hàm include file HTML module
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ======================================================================
// PHẦN 1: LƯU TRỮ DỮ LIỆU
// ======================================================================

/**
 * Đánh dấu dữ liệu cũ của 1 khoa + 1 ngày là "Lịch sử" và trả về số phiên bản mới
 * (Thay thế cho cơ chế xoá cứng để đảm bảo lưu vết dữ liệu y khoa)
 */
function deactivateExistingReport(khoa, ngay) {
  const ss = getSpreadsheet();
  let nextVersion = 1;

  // Xử lý trong Data_ThongKe
  const sheetTK = ss.getSheetByName("Data_ThongKe");
  if (sheetTK && sheetTK.getLastRow() > 1) {
    const dataTK = sheetTK.getDataRange().getValues();
    for (let i = dataTK.length - 1; i >= 1; i--) {
      const rowDate = Utilities.formatDate(new Date(dataTK[i][1]), "GMT+7", "yyyy-MM-dd");
      if (dataTK[i][2] === khoa && rowDate === ngay) {
        const currentStatus = dataTK[i][22];
        if (currentStatus !== "Lịch sử") {
          const ver = parseInt(dataTK[i][23]) || 1;
          if (ver >= nextVersion) nextVersion = ver + 1;
          
          sheetTK.getRange(i + 1, 23).setValue("Lịch sử"); // Cột W (index 22)
          sheetTK.getRange(i + 1, 25).setValue(new Date()); // Cột Y (index 24)
        }
      }
    }
  }

  // Xử lý trong Data_BenhNhan
  const sheetBN = ss.getSheetByName("Data_BenhNhan");
  if (sheetBN && sheetBN.getLastRow() > 1) {
    const dataBN = sheetBN.getDataRange().getValues();
    for (let i = dataBN.length - 1; i >= 1; i--) {
      const rowDate = Utilities.formatDate(new Date(dataBN[i][1]), "GMT+7", "yyyy-MM-dd");
      if (dataBN[i][2] === khoa && rowDate === ngay) {
        if (dataBN[i][32] !== "Lịch sử") {
          sheetBN.getRange(i + 1, 33).setValue("Lịch sử"); // Cột AG (index 32)
        }
      }
    }
  }
  
  return nextVersion;
}

/**
 * Lưu báo cáo mới
 */
function saveReport(payload) {
  try {
    const ss = getSpreadsheet();
    const sheetTK =
      ss.getSheetByName("Data_ThongKe") || ss.insertSheet("Data_ThongKe");
    const sheetBN =
      ss.getSheetByName("Data_BenhNhan") || ss.insertSheet("Data_BenhNhan");

    const khoa = payload.khoa;
    const ngayStr = payload.ngay;
    const ngay = new Date(ngayStr + "T00:00:00+07:00");
    const sdt = payload.sdtNguoiBaoCao || "";
    const now = new Date();

    // 1. Vô hiệu hóa bản ghi cũ và lấy version tiếp theo
    const phienBan = deactivateExistingReport(khoa, ngayStr);

    // 2. Lưu số liệu thống kê (Data_ThongKe)
    const rowTKBase = [now, ngay, khoa, sdt];
    const thongKeRows = payload.thongKe.map((row) => {
      return [
        ...rowTKBase,
        row.doiTuong,
        row.c1, row.c21, row.c22, row.c3, row.c41, row.c42,
        row.c51, row.c52, row.c61, row.c62, row.c71, row.c72,
        payload.chiSo.phauThuat,
        payload.chiSo.phauThuatCC,
        payload.chiSo.mau,
        payload.chiSo.treSinh,
        payload.chiSo.treSinhMo,
        "Đã nộp",
        phienBan,
        now,
      ];
    });
    if (thongKeRows.length > 0) {
      sheetTK
        .getRange(sheetTK.getLastRow() + 1, 1, thongKeRows.length, 25)
        .setValues(thongKeRows);
    }

    // 3. Lưu danh sách bệnh nhân (Data_BenhNhan)
    if (payload.benhNhan && payload.benhNhan.length > 0) {
      const bnRows = payload.benhNhan.map((bn) => {
        return [
          now, ngay, khoa,
          bn.maHSBA, bn.maBN, bn.ten, bn.ngaySinh, bn.tuoi, bn.gioi, bn.danToc, bn.quocTich,
          bn.loaiGiayTo, bn.soGiayTo, bn.ngayCap, bn.noiCap, bn.bhyt,
          bn.tinh, bn.xa, bn.thon,
          bn.htTinh, bn.htXa, bn.htThon,
          bn.lyDo, bn.phTinh, bn.phXa, bn.phThon,
          bn.ngayTaiNan, bn.ngayVao, bn.ngayRa, bn.ketQua, bn.tinhTrang,
          bn.chuyenTu, bn.chanDoan, bn.chanDoanChiTiet, bn.icd, bn.fileHSBAUrl,
          "Đã nộp", phienBan, now
        ];
      });
      sheetBN
        .getRange(sheetBN.getLastRow() + 1, 1, bnRows.length, 39)
        .setValues(bnRows);
    }

    // 4. Xóa các file cũ trên Drive (nếu có yêu cầu xóa)
    if (payload.filesToDelete && payload.filesToDelete.length > 0) {
       payload.filesToDelete.forEach(url => deleteDriveFile(url));
    }

    // 5. Cập nhật sheet Báo cáo tổng
    refreshPhysicalSummarySheetRange(ngayStr, ngayStr);

    return { success: true, version: phienBan };
  } catch (e) {
    throw new Error("Lỗi lưu báo cáo: " + e.toString());
  }
}

// ======================================================================
// PHẦN 2: TRUY XUẤT DỮ LIỆU ĐÃ LƯU
// ======================================================================

/**
 * Lấy báo cáo đã lưu của 1 khoa vào 1 ngày (Bản ghi mới nhất)
 */
function getExistingReport(khoa, ngay) {
  try {
    const ss = getSpreadsheet();
    let result = { found: false, thongKe: [], benhNhan: [], chiSo: {} };

    // Đọc Data_ThongKe
    const sheetTK = ss.getSheetByName("Data_ThongKe");
    if (sheetTK && sheetTK.getLastRow() > 1) {
      const data = sheetTK.getDataRange().getValues();
      // Tìm các dòng khớp (khoa, ngay) và không phải "Lịch sử"
      const matches = data.filter((row, idx) => {
        if (idx === 0) return false;
        const rowDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
        return row[2] === khoa && rowDate === ngay && row[22] !== "Lịch sử";
      });

      if (matches.length > 0) {
        result.found = true;
        result.sdtNguoiBaoCao = matches[0][3];
        result.chiSo = {
          phauThuat: matches[0][17],
          phauThuatCC: matches[0][18],
          mau: matches[0][19],
          treSinh: matches[0][20],
          treSinhMo: matches[0][21],
        };
        result.thongKe = matches.map((m) => {
          return {
            doiTuong: m[4],
            c1: m[5], c21: m[6], c22: m[7], c3: m[8], c41: m[9], c42: m[10],
            c51: m[11], c52: m[12], c61: m[13], c62: m[14], c71: m[15], c72: m[16],
          };
        });
      }
    }

    // Đọc Data_BenhNhan
    const sheetBN = ss.getSheetByName("Data_BenhNhan");
    if (sheetBN && sheetBN.getLastRow() > 1) {
      const dataBN = sheetBN.getDataRange().getValues();
      result.benhNhan = dataBN
        .filter((row, idx) => {
          if (idx === 0) return false;
          const rowDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
          return row[2] === khoa && rowDate === ngay && row[36] !== "Lịch sử";
        })
        .map((bn) => {
          return {
            maHSBA: bn[3], maBN: bn[4], ten: bn[5], ngaySinh: bn[6], tuoi: bn[7], gioi: bn[8],
            danToc: bn[9], quocTich: bn[10], loaiGiayTo: bn[11], soGiayTo: bn[12], ngayCap: bn[13],
            noiCap: bn[14], bhyt: bn[15], tinh: bn[16], xa: bn[17], thon: bn[18],
            htTinh: bn[19], htXa: bn[20], htThon: bn[21], lyDo: bn[22], phTinh: bn[23], phXa: bn[24], phThon: bn[25],
            ngayTaiNan: bn[26], ngayVao: bn[27], ngayRa: bn[28], ketQua: bn[29], tinhTrang: bn[30],
            chuyenTu: bn[31], chanDoan: bn[32], chanDoanChiTiet: bn[33], icd: bn[34], fileHSBAUrl: bn[35]
          };
        });
    }

    return result;
  } catch (e) {
    throw new Error("Lỗi lấy dữ liệu: " + e.toString());
  }
}

// ======================================================================
// PHẦN 3: TỔNG HỢP BÁO CÁO TOÀN VIỆN
// ======================================================================

/**
 * Lấy dữ liệu tổng hợp từ nhiều khoa trong một khoảng ngày
 */
function getAggregatedReportRange(startDate, endDate) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("Data_ThongKe");
    if (!sheet) return { khoaDaNop: [], thongKe: [], chiSo: {} };

    let data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { khoaDaNop: [], thongKe: [], chiSo: {} };

    // Tìm version mới nhất cho mỗi cặp (khoa, ngày)
    let latestTsTK = {}; // khoa_ngày -> timestamp lớn nhất
    for (let i = 1; i < data.length; i++) {
      const rowDate = Utilities.formatDate(new Date(data[i][1]), "GMT+7", "yyyy-MM-dd");
      const khoa = data[i][2];
      const trangThai = data[i][22];
      const ts = new Date(data[i][0]).getTime();
      const key = khoa + "_" + rowDate;

      if (trangThai !== "Lịch sử") {
        if (!latestTsTK[key] || ts > latestTsTK[key]) {
          latestTsTK[key] = ts;
        }
      }
    }

    // Lọc data: chỉ lấy các bản ghi nộp mới nhất trong khoảng ngày
    data = data.filter((row) => {
      const trangThai = row[22];
      if (trangThai === "Lịch sử") return false;
      const rowDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
      const khoa = row[2];
      const ts = new Date(row[0]).getTime();
      const key = khoa + "_" + rowDate;
      return ts === latestTsTK[key];
    });

    const ROW_NAMES = getConfigDoiTuong();
    let report = {
      thongKe: ROW_NAMES.map((name) => ({
        doiTuong: name,
        vals: Array(12).fill(0),
      })),
      chiSo: { pt: 0, ptcc: 0, mau: 0, tre: 0, mo: 0 },
      khoaDaNop: [],
    };

    let departmentDates = {};

    // === Bước 1: Tìm ngày đầu và cuối THỰC TẾ có dữ liệu trong dải ===
    let minDateFound = "";
    let maxDateFound = "";

    data.forEach((row) => {
      const trangThai = row[22];
      if (trangThai === "Lịch sử") return;

      const rowDate = Utilities.formatDate(
        new Date(row[1]), "GMT+7", "yyyy-MM-dd"
      );
      if (rowDate >= startDate && rowDate <= endDate) {
        if (minDateFound === "" || rowDate < minDateFound) minDateFound = rowDate;
        if (maxDateFound === "" || rowDate > maxDateFound) maxDateFound = rowDate;
      }
    });

    // Nếu không tìm thấy ngày nào có dữ liệu
    if (minDateFound === "") return report;

    // === Bước 2: Duyệt từng ngày TRONG KHOẢNG CÓ DỮ LIỆU để cộng dồn ===
    let currentDate = new Date(minDateFound + "T00:00:00+07:00");
    const lastDate = new Date(maxDateFound + "T00:00:00+07:00");

    while (currentDate <= lastDate) {
      const dStr = Utilities.formatDate(currentDate, "GMT+7", "yyyy-MM-dd");
      
      // Lấy dữ liệu của ngày dStr
      const dayData = data.filter(r => Utilities.formatDate(new Date(r[1]), "GMT+7", "yyyy-MM-dd") === dStr);
      
      // Tập hợp các khoa đã báo cáo trong ngày này
      const khoasInDay = [...new Set(dayData.map(r => r[2]))];
      khoasInDay.forEach(k => {
        if (!departmentDates[k]) departmentDates[k] = [];
        departmentDates[k].push(dStr);
      });

      // Cộng dồn chỉ số (Lưu ý: Chỉ số toàn viện = Tổng chỉ số của tất cả các khoa TRONG TẤT CẢ CÁC NGÀY)
      // Vì mỗi khoa báo cáo "số liệu trong ngày", nên tổng cộng dồn là đúng.
      dayData.forEach(row => {
        // Chỉ cộng chỉ số 1 lần cho mỗi dòng đối tượng đầu tiên của khoa trong ngày
        // (Tránh cộng lặp do 1 khoa có 4 dòng đối tượng)
        if (row[4] === ROW_NAMES[0]) {
          report.chiSo.pt += (parseFloat(row[17]) || 0);
          report.chiSo.ptcc += (parseFloat(row[18]) || 0);
          report.chiSo.mau += (parseFloat(row[19]) || 0);
          report.chiSo.tre += (parseFloat(row[20]) || 0);
          report.chiSo.mo += (parseFloat(row[21]) || 0);
        }

        const idx = ROW_NAMES.indexOf(row[4]);
        if (idx !== -1) {
          for (let j = 0; j < 12; j++) {
            report.thongKe[idx].vals[j] += (parseFloat(row[5 + j]) || 0);
          }
        }
      });

      currentDate.setDate(currentDate.getDate() + 1);
    }

    report.khoaDaNop = Object.keys(departmentDates).map(k => ({
      ten: k,
      days: departmentDates[k]
    }));

    // Bổ sung danh sách bệnh nhân chi tiết
    report.benhNhan = getAggregatedPatients(startDate, endDate);

    return report;
  } catch (e) {
    throw new Error("Lỗi tổng hợp: " + e.toString());
  }
}

/**
 * Lấy danh sách bệnh nhân từ tất cả các khoa trong khoảng ngày
 */
function getAggregatedPatients(startDate, endDate) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Data_BenhNhan");
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  
  // Lấy danh sách (khoa, ngày) -> version mới nhất từ Data_ThongKe để map sang
  const sheetTK = ss.getSheetByName("Data_ThongKe");
  let latestVersions = {}; 
  if (sheetTK && sheetTK.getLastRow() > 1) {
     const tkData = sheetTK.getDataRange().getValues();
     tkData.forEach((row, i) => {
        if (i === 0 || row[22] === "Lịch sử") return;
        const dStr = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
        latestVersions[row[2] + "_" + dStr] = row[23];
     });
  }

  return data.filter((row, idx) => {
    if (idx === 0 || row[36] === "Lịch sử") return false;
    const rowDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
    const key = row[2] + "_" + rowDate;
    return rowDate >= startDate && rowDate <= endDate && row[37] === latestVersions[key];
  }).map(bn => ({
    khoa: bn[2], ten: bn[5], tuoi: bn[7], gioi: bn[8], lyDo: bn[22], ketQua: bn[29], chanDoan: bn[32], icd: bn[34], fileHSBAUrl: bn[35],
    maHSBA: bn[3], maBN: bn[4], ngaySinh: bn[6], danToc: bn[9], quocTich: bn[10], loaiGiayTo: bn[11], soGiayTo: bn[12], ngayCap: bn[13],
    noiCap: bn[14], bhyt: bn[15], tinh: bn[16], xa: bn[17], thon: bn[18],
    htTinh: bn[19], htXa: bn[20], htThon: bn[21], phTinh: bn[23], phXa: bn[24], phThon: bn[25],
    ngayTaiNan: bn[26], ngayVao: bn[27], ngayRa: bn[28], tinhTrang: bn[30],
    chuyenTu: bn[31], chanDoanChiTiet: bn[33]
  }));
}

// ======================================================================
// PHẦN 4: THEO DÕI TRẠNG THÁI BÁO CÁO
// ======================================================================

/**
 * Lấy trạng thái khoa nào đã báo cáo vào ngày nào
 * Trả về ma trận: khoa × ngày → đã nộp / chưa nộp
 */
function getReportStatus(startDate, endDate) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("Data_ThongKe");

    // Danh sách TẤT CẢ các khoa cần theo dõi
    const ALL_DEPARTMENTS = getConfigKhoaPhong();

    // Tạo danh sách ngày trong dải
    let dates = [];
    let current = new Date(startDate + "T00:00:00+07:00");
    const end = new Date(endDate + "T00:00:00+07:00");
    while (current <= end) {
      dates.push(Utilities.formatDate(current, "GMT+7", "yyyy-MM-dd"));
      current.setDate(current.getDate() + 1);
    }

    // Đọc data để tìm các (khoa, ngày) đã có
    let reportedMap = {}; // key: "KHOA_ngày" → true
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const rowDate = Utilities.formatDate(new Date(data[i][1]), "GMT+7", "yyyy-MM-dd");
        if (rowDate >= startDate && rowDate <= endDate && data[i][22] !== "Lịch sử") {
          const khoa = data[i][2];
          reportedMap[khoa + "_" + rowDate] = true;
        }
      }
    }

    // Tạo kết quả: mỗi khoa + danh sách ngày đã nộp
    let departments = ALL_DEPARTMENTS.map(khoa => {
      let reported = [];
      dates.forEach(d => {
        if (reportedMap[khoa + "_" + d]) {
          reported.push(d);
        }
      });
      return { khoa: khoa, reported: reported };
    });

    return { dates: dates, departments: departments };
  } catch (e) {
    throw new Error("Lỗi lấy trạng thái: " + e.toString());
  }
}

/**
 * Hàm làm mới sheet vật lý "Báo cáo tổng" theo khoảng thời gian
 */
function refreshPhysicalSummarySheetRange(startDate, endDate) {
  const ss = getSpreadsheet();
  let sheetSummary =
    ss.getSheetByName("Báo cáo tổng") || ss.insertSheet("Báo cáo tổng");
  sheetSummary.clear();

  const report = getAggregatedReportRange(startDate, endDate);
  const ROW_NAMES = getConfigDoiTuong();

  let headerRows = [
    ["BÁO CÁO TỔNG HỢP TOÀN VIỆN"],
    ["Từ ngày: " + startDate + " - Đến ngày: " + endDate],
    ["Khoa đã nộp: " + report.khoaDaNop.map(k => k.ten + " (" + k.days.length + " ngày)").join(", ")],
    [""],
    ["TT", "Khám, cấp cứu", "BN cũ (1)", "Khám bệnh - Tổng số (2.1)", "Khám bệnh - BHYT (2.2)", "Vào viện (3)", "Chuyển viện - Ngoại trú (4.1)", "Chuyển viện - Nội trú (4.2)", "Ra viện - Tổng số (5.1)", "Ra viện - Nặng xin về (5.2)", "Tử vong - Nội viện (6.1)", "Tử vong - Ngoại viện (6.2)", "BN hiện có - Tổng số (7.1)", "BN hiện có - Ca nặng (7.2)"]
  ];

  let bodyRows = report.thongKe.map((row, idx) => {
    return [idx + 1, row.doiTuong, ...row.vals];
  });

  let footerRows = [
    [""],
    ["CÁC CHỈ SỐ KHÁC"],
    ["1. Số ca phẫu thuật (loại 3 trở lên):", report.chiSo.pt],
    ["2. Trong đó, PT cấp cứu do tai nạn:", report.chiSo.ptcc],
    ["3. TS trẻ sinh tại CSKCB:", report.chiSo.tre],
    ["4. Trong đó, số trẻ sinh mổ đẻ:", report.chiSo.mo],
    ["5. Tổng số lượng máu dự trữ (ml):", report.chiSo.mau]
  ];

  const allRows = [...headerRows, ...bodyRows, ...footerRows];
  sheetSummary.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
  
  // Định dạng
  sheetSummary.getRange("A1").setFontSize(14).setFontWeight("bold");
  sheetSummary.getRange(5, 1, 1, 14).setFontWeight("bold").setBackground("#f3f3f3");
}

// ======================================================================
// PHẦN 5: API ROUTER (POST REQUEST)
// ======================================================================

/**
 * Xử lý POST request (API Router)
 */
function doPost(e) {
  try {
    // Phân tích dữ liệu JSON từ body request
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;

    // --- LUỒNG 1: XỬ LÝ UPLOAD FILE ---
    if (action === 'uploadFile' || requestData.filename) {
      var filename = requestData.filename || (requestData.payload && requestData.payload.filename) || "file_upload";
      var mimetype = requestData.mimetype || (requestData.payload && requestData.payload.mimetype) || "application/octet-stream";
      var base64Data = requestData.data || (requestData.payload && requestData.payload.data) || "";
      
      if (!base64Data) {
         throw new Error("Không tìm thấy dữ liệu file (base64) để upload.");
      }

      var folderName = "Báo cáo tết";
      var folder;
      try {
        var folders = DriveApp.getFoldersByName(folderName);
        folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      } catch (err) {
        folder = DriveApp.getRootFolder();
      }
      
      var decoded = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decoded, mimetype, filename);
      
      var file;
      try {
        file = folder.createFile(blob);
      } catch (e1) {
        try {
          file = DriveApp.createFile(blob);
        } catch (e2) {
          throw "LỖI QUYỀN NGHIÊM TRỌNG: " + e2.toString();
        }
      }
      
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        fileId: file.getId(),
        url: file.getUrl()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- LUỒNG 2: XỬ LÝ CÁC API CALL TỪ FRONTEND ---
    if (action) {
      var payload = requestData.payload;
      var result = null;

      switch (action) {
        case 'saveReport':
          result = saveReport(payload);
          break;
        case 'getExistingReport':
          result = getExistingReport(payload.khoa, payload.ngay);
          break;
        case 'getAggregatedReportRange':
          result = getAggregatedReportRange(payload.startDate, payload.endDate);
          break;
        case 'getReportStatus':
          result = getReportStatus(payload.startDate, payload.endDate);
          break;
        default:
          throw new Error('Action không hợp lệ: ' + action);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: result
      })).setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error('Thiếu action trong payload');
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Xóa file trên Google Drive (chuyển vào thùng rác)
 */
function deleteDriveFile(fileUrl) {
  try {
    if (!fileUrl) return { success: false, message: "URL trống" };
    
    // Extract ID từ URL (dạng /d/FILE_ID/...)
    var match = String(fileUrl).match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      var fileId = match[1];
      DriveApp.getFileById(fileId).setTrashed(true); // Chỉ chuyển vào thùng rác cho an toàn
      return { success: true };
    }
    return { success: false, message: "Không tìm thấy ID file từ URL" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

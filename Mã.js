/**
 * Backend Google Apps Script - Tết Bính Ngọ 2026
 * Cập nhật đầy đủ các hàm để khớp với Frontend.
 */


/**
 * HÀM NÀY ĐỂ KÍCH HOẠT QUYỀN DRIVE (Chạy thủ công 1 lần nếu gặp lỗi)
 * Chọn hàm này trên thanh công cụ và bấm Run > Review Permissions > Allow
 */
function FORCE_AUTH() {
  DriveApp.getRootFolder();
  console.log("Đã cấp quyền Drive thành công!");
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile("index");
  // Truyền URL WebApp xuống để frontend gọi fetch
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle("Báo cáo Tết 2026 - Bệnh viện Bình Định")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ======================================================================
// PHẦN 1: LƯU TRỮ DỮ LIỆU
// ======================================================================

/**
 * Đánh dấu dữ liệu cũ của 1 khoa + 1 ngày là "Lịch sử" và trả về số phiên bản mới
 * (Thay thế cho cơ chế xoá cứng để đảm bảo lưu vết dữ liệu y khoa)
 */
function deactivateExistingReport(khoa, ngay) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
          
          // Đánh dấu là Lịch sử (cột 23 là W), và gán phiên bản nếu chưa có (cột 24 là X)
          sheetTK.getRange(i + 1, 23).setValue("Lịch sử");
          if (!dataTK[i][23]) sheetTK.getRange(i + 1, 24).setValue(ver);
        }
      }
    }
  }

  // Xử lý trong Data_DanhSachCa
  const sheetBN = ss.getSheetByName("Data_DanhSachCa");
  if (sheetBN && sheetBN.getLastRow() > 1) {
    const dataBN = sheetBN.getDataRange().getValues();
    for (let i = dataBN.length - 1; i >= 1; i--) {
      const rowDate = Utilities.formatDate(new Date(dataBN[i][1]), "GMT+7", "yyyy-MM-dd");
      if (dataBN[i][2] === khoa && rowDate === ngay) {
        const currentStatus = dataBN[i][40];
        if (currentStatus !== "Lịch sử") {
          const ver = parseInt(dataBN[i][41]) || 1;
          sheetBN.getRange(i + 1, 41).setValue("Lịch sử");
          if (!dataBN[i][41]) sheetBN.getRange(i + 1, 42).setValue(ver);
        }
      }
    }
  }

  return nextVersion;
}

// ======================================================================
// PHẦN QUẢN LÝ DUYỆT BÁO CÁO
// ======================================================================

function checkApprovalStatus(ngay) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data_DuyetNgay");
    if (!sheet) return false;
    
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let rowDate = '';
      if (data[i][0]) {
         try {
           rowDate = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "yyyy-MM-dd");
         } catch(e) { rowDate = String(data[i][0]); }
      }
      if (rowDate === ngay && data[i][1] === "Đã duyệt") {
        return true;
      }
    }
    return false;
  } catch(e) {
    return false;
  }
}

function toggleApproval(ngay, isApproved) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Data_DuyetNgay");
    if (!sheet) {
      sheet = ss.insertSheet("Data_DuyetNgay");
      sheet.appendRow(["Ngày", "Trạng thái", "Thời gian cập nhật"]);
    }
    
    let data = sheet.getDataRange().getValues();
    let foundIndex = -1;
    for (let i = 1; i < data.length; i++) {
      let rowDate = '';
      if (data[i][0]) {
         try {
           rowDate = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "yyyy-MM-dd");
         } catch(e) { rowDate = String(data[i][0]); }
      }
      if (rowDate === ngay) {
        foundIndex = i + 1;
        break;
      }
    }
    
    let statusText = isApproved ? "Đã duyệt" : "Chưa duyệt";
    if (foundIndex > -1) {
      sheet.getRange(foundIndex, 2).setValue(statusText);
      sheet.getRange(foundIndex, 3).setValue(new Date());
    } else {
      sheet.appendRow([ngay, statusText, new Date()]);
    }
    
    return { success: true, status: isApproved };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Hàm hỗ trợ Trigger: Tự động duyệt báo cáo cho ngày hiện tại.
 * Thiết lập Trigger chạy hàm này vào khung giờ mong muốn (VD: 23:59 hàng ngày).
 */
function autoApproveCurrentDate() {
  const vnNow = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Ho_Chi_Minh' }));
  const yyyy = vnNow.getFullYear();
  const mm = String(vnNow.getMonth() + 1).padStart(2, '0');
  const dd = String(vnNow.getDate()).padStart(2, '0');
  const todayStr = `${yyyy}-${mm}-${dd}`;
  
  Logger.log("Đang tự động duyệt cho ngày: " + todayStr);
  return toggleApproval(todayStr, true);
}

/**
 * Hàm lưu báo cáo — cơ chế UPSERT (xoá cũ rồi ghi mới)
 * Đảm bảo mỗi (khoa, ngày) chỉ có 1 bộ dữ liệu duy nhất
 */
function saveReport(payload) {
  try {
    // Kiểm tra duyệt
    if (checkApprovalStatus(payload.ngay)) {
      throw new Error("Dữ liệu ngày " + payload.ngay + " đã được KHTH duyệt. Không thể thêm/sửa báo cáo!");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTK =
      ss.getSheetByName("Data_ThongKe") || ss.insertSheet("Data_ThongKe");
    const sheetBN =
      ss.getSheetByName("Data_DanhSachCa") || ss.insertSheet("Data_DanhSachCa");

    // Tạo header nếu sheet trống
    if (sheetTK.getLastRow() === 0) {
      sheetTK.appendRow([
        "Timestamp", "Ngày trực", "Khoa", "Đối tượng",
        "(1)", "(2.1)", "(2.2)", "(3)", "(4.1)", "(4.2)",
        "(5.1)", "(5.2)", "(6.1)", "(6.2)", "(7.1)", "(7.2)",
        "PT", "PT_CC", "Mau", "Tre", "Mo", "SDT_NguoiBaoCao",
        "Trạng thái", "Phiên bản", "(4.3)", "ChiTietChuyenKhoa", "(3.2)", "ChiTietNhanKhoa"
      ]);
    }

    if (sheetBN.getLastRow() === 0) {
      sheetBN.appendRow([
        "Timestamp", "Ngày trực", "Khoa",
        "Mã HSBA", "Mã BN", "Họ tên", "Ngày sinh", "Tuổi", "Giới",
        "Dân tộc", "Quốc tịch", "BHYT",
        "Tỉnh", "Huyện", "Xã", "Thôn",
        "Lý do", "Ngày tai nạn", "Ngày vào", "Ngày ra",
        "Kết quả", "Tình trạng", "Chuyển từ", "Nơi phát hiện",
        "Chẩn đoán", "ICD-10", "File HSBA URL",
        "Loại giấy tờ", "Số giấy tờ", "Ngày cấp", "Nơi cấp",
        "HT Tỉnh", "HT Huyện", "HT Xã", "HT Thôn",
        "PH Tỉnh", "PH Huyện", "PH Xã", "PH Thôn",
        "Chẩn đoán chi tiết",
        "Trạng thái", "Phiên bản"
      ]);
    }

    // Đánh dấu dữ liệu cũ là "Lịch sử" thay vì xoá (Cơ chế lưu vết)
    const currentVersion = deactivateExistingReport(payload.khoa, payload.ngay);

    // Ghi dữ liệu mới
    const timestamp = new Date();
    payload.thongKe.forEach((row) => {
      sheetTK.appendRow([
        timestamp,
        payload.ngay,
        payload.khoa,
        row.doiTuong,
        row.c1, row.c21, row.c22, row.c3,
        row.c41, row.c42, row.c51, row.c52,
        row.c61, row.c62, row.c71, row.c72,
        payload.chiSo.phauThuat,
        payload.chiSo.phauThuatCC,
        payload.chiSo.mau,
        payload.chiSo.treSinh,
        payload.chiSo.treSinhMo,
        "'" + (payload.sdtNguoiBaoCao || ''),
        "Mới nhất",
        currentVersion,
        row.c43 || 0,
        row.chuyenKhoaChiTiet || "[]",
        row.c32 || 0,
        row.nhanKhoaChiTiet || "[]"
      ]);
    });

    if (payload.benhNhan && payload.benhNhan.length > 0) {
      payload.benhNhan.forEach((bn) => {
        if (bn.ten && bn.ten.trim() !== "") {
          // Xử lý file HSBA: Payload giờ sẽ gửi URL file đã upload trước đó
          let fileUrl = bn.fileHSBAUrl || '';
          
          // Tạo summary cho nơi phát hiện cũ
          const noiPhatHienSummary = [bn.phThon, bn.phXa, bn.phHuyen, bn.phTinh].filter(x => x).join(', ');

          sheetBN.appendRow([
            timestamp,
            payload.ngay,
            payload.khoa,
            bn.maHSBA || '', bn.maBN || '', bn.ten,
            bn.ngaySinh || '', bn.tuoi || '', bn.gioi || '',
            bn.danToc || '', bn.quocTich || '', bn.bhyt || '',
            bn.tinh || '', bn.huyen || '', bn.xa || '', bn.thon || '',
            bn.lyDo || '',
            bn.ngayTaiNan || '',
            bn.ngayVao || '', bn.ngayRa || '',
            bn.ketQua || '', bn.tinhTrang || '',
            bn.chuyenTu || '', noiPhatHienSummary || '',
            bn.chanDoan || '', bn.icd || '',
            fileUrl,
            // New columns:
            bn.loaiGiayTo || '', bn.soGiayTo || '', bn.ngayCap || '', bn.noiCap || '',
            bn.htTinh || '', bn.htHuyen || '', bn.htXa || '', bn.htThon || '',
            bn.phTinh || '', bn.phHuyen || '', bn.phXa || '', bn.phThon || '',
            bn.chanDoanChiTiet || '',
            "Mới nhất",
            currentVersion
          ]);
        }
      });
    }

    // Tự động làm mới Sheet vật lý cho ngày báo cáo này
    refreshPhysicalSummarySheetRange(payload.ngay, payload.ngay);

    // KHÔNG XÓA FILE CŨ TRÊN DRIVE NỮA (Để đảm bảo lưu vết - Audit Trail)
    // Các file cũ vẫn được giữ lại để các phiên bản Lịch sử có thể truy cập được.
    /* 
    if (payload.filesToDelete && Array.isArray(payload.filesToDelete)) {
      payload.filesToDelete.forEach(function(url) {
        deleteDriveFile(url);
      });
    } 
    */

    return { success: true };
  } catch (e) {
    throw new Error(e.toString());
  }
}

// ======================================================================

/**
 * Hàm upload file độc lập (được gọi từ Frontend trước khi lưu báo cáo)
 */
function uploadSingleFile(fileData) {
  try {
    const folderName = "Báo cáo tết";
    let folder;
    try {
      const folders = DriveApp.getFoldersByName(folderName);
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    } catch (e) {
      folder = DriveApp.getRootFolder();
    }

    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.data),
      fileData.mimeType,
      fileData.name
    );
    const file = folder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      file.setAccess(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    return {
      success: true,
      url: 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing'
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ======================================================================
// PHẦN 2: ĐỌC DỮ LIỆU ĐÃ GỬI (CHO TÍNH NĂNG LOAD + SỬA)
// ======================================================================

/**
 * Lấy dữ liệu đã gửi trước đó của 1 khoa + 1 ngày
 * Dùng để load vào form khi user chọn khoa + ngày
 */
function getExistingReport(khoa, ngay) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let result = { found: false, thongKe: [], benhNhan: [], chiSo: {}, isApproved: checkApprovalStatus(ngay) };

    // Đọc Data_ThongKe
    const sheetTK = ss.getSheetByName("Data_ThongKe");
    if (sheetTK && sheetTK.getLastRow() > 1) {
      const dataTK = sheetTK.getDataRange().getValues();
      
      const homNayDate = new Date(ngay);
      const homQuaDate = new Date(homNayDate.getTime() - 24 * 60 * 60 * 1000);
      const homQuaStr = Utilities.formatDate(homQuaDate, "GMT+7", "yyyy-MM-dd");
      const autoCheck = (str) => (str || "").replace(/\*/g, "").trim();

      // 1. Tìm báo cáo ngày hôm nay và Lấy báo cáo tồn hôm qua
      let prevDataMap = {};
      let expectedNhans = {};
      let expectedChuyens = {};
      let transferDetails = [];

      for (let i = 1; i < dataTK.length; i++) {
        const rowDate = Utilities.formatDate(new Date(dataTK[i][1]), "GMT+7", "yyyy-MM-dd");
        const rowKhoa = dataTK[i][2];
        const doiTuong = dataTK[i][3];
        const trangThai = dataTK[i][22];

        if (trangThai === "Lịch sử") continue;

        // Báo cáo hôm nay của khoa
        if (rowKhoa === khoa && rowDate === ngay) {
          result.found = true;
          // Lưu ý: Cột 7.1 là index 14 (cột O) trong sheet Data_ThongKe
          result.thongKe.push({
            doiTuong: doiTuong,
            c1: dataTK[i][4], c21: dataTK[i][5], c22: dataTK[i][6],
            c3: dataTK[i][7], c41: dataTK[i][8], c42: dataTK[i][9],
            c51: dataTK[i][10], c52: dataTK[i][11],
            c61: dataTK[i][12], c62: dataTK[i][13],
            c71: dataTK[i][14], c72: dataTK[i][15],
            c43: dataTK[i][24] || 0,
            chuyenKhoaChiTiet: dataTK[i][25] || "[]",
            c32: dataTK[i][26] || 0,
            nhanKhoaChiTiet: dataTK[i][27] || "[]"
          });
          // Lấy chiSo từ dòng đầu tiên tìm thấy
          if (!result.chiSo || !result.chiSo.phauThuat) {
            result.chiSo = {
              phauThuat: dataTK[i][16] || 0,
              phauThuatCC: dataTK[i][17] || 0,
              mau: dataTK[i][18] || 0,
              treSinh: dataTK[i][19] || 0,
              treSinhMo: dataTK[i][20] || 0,
            };
            result.sdtNguoiBaoCao = dataTK[i][21] || '';
          }
        }

        // Báo cáo hôm qua của khoa (để lấy tồn c1)
        if (rowKhoa === khoa && rowDate === homQuaStr) {
            prevDataMap[autoCheck(doiTuong)] = dataTK[i][14] || 0; 
        }

        // Báo cáo chuyển/nhận khoa liên quan đến khoa hiện tại trong hôm nay
        if (rowKhoa !== khoa && rowDate === ngay) {
            // 1. Khoa khác báo CHUYỂN ĐẾN khoa này -> Khoa này cần NHẬN
            const dtDetails = dataTK[i][25]; // ChuyenKhoaChiTiet của khoa khác
            if (dtDetails && dtDetails !== "[]") {
               try {
                 const arr = JSON.parse(dtDetails);
                 const targetKhoa = (khoa || "").trim().toUpperCase();
                 arr.forEach(item => {
                    const itemTarget = (item.khoa || "").trim().toUpperCase();
                    if (itemTarget === targetKhoa) {
                        const soLuong = parseInt(item.soLuong) || 0;
                        if (soLuong > 0) {
                            if (!expectedNhans[doiTuong]) expectedNhans[doiTuong] = 0;
                            expectedNhans[doiTuong] += soLuong;
                            transferDetails.push({
                                type: "Nhận từ khoa khác",
                                fromKhoa: rowKhoa,
                                doiTuong: doiTuong,
                                soLuong: soLuong
                            });
                        }
                    }
                 });
               } catch(e) {}
            }
            
            // 2. Khoa khác báo NHẬN TỪ khoa này -> Khoa này cần CHUYỂN
            const nkDetails = dataTK[i][27]; // NhanKhoaChiTiet của khoa khác
            if (nkDetails && nkDetails !== "[]") {
               try {
                 const arr = JSON.parse(nkDetails);
                 const targetKhoa = (khoa || "").trim().toUpperCase();
                 arr.forEach(item => {
                    const itemTarget = (item.khoa || "").trim().toUpperCase();
                    if (itemTarget === targetKhoa) {
                        const soLuong = parseInt(item.soLuong) || 0;
                        if (soLuong > 0) {
                            if (!expectedChuyens[doiTuong]) expectedChuyens[doiTuong] = 0;
                            expectedChuyens[doiTuong] += soLuong;
                            transferDetails.push({
                                type: "Chuyển đi khoa khác",
                                fromKhoa: rowKhoa,
                                doiTuong: doiTuong,
                                soLuong: soLuong
                            });
                        }
                    }
                 });
               } catch(e) {}
            }
        }
      }

      result.prevC71Map = prevDataMap;
      
      // Lọc bỏ các cảnh báo trùng lặp (nếu báo cáo hiện tại đã có dữ liệu khớp)
      if (result.found && transferDetails.length > 0) {
          const norm = (s) => (s || "").trim().toUpperCase();
          transferDetails = transferDetails.filter(alert => {
              const dt = alert.doiTuong;
              const khoaLienQuan = norm(alert.fromKhoa);
              const reportRow = result.thongKe.find(r => r.doiTuong === dt);
              if (!reportRow) return true;
              
              const currentDetails = alert.type.includes("Nhận") ? reportRow.nhanKhoaChiTiet : reportRow.chuyenKhoaChiTiet;
              try {
                  const arr = JSON.parse(currentDetails);
                  // Không hiển thị nếu đã có dòng khớp cả Khoa và Số lượng
                  const matched = arr.some(item => 
                      norm(item.khoa) === khoaLienQuan && 
                      parseInt(item.soLuong) === parseInt(alert.soLuong)
                  );
                  return !matched;
              } catch(e) { return true; }
          });
      }

      if (transferDetails.length > 0) {
          result.transferAlerts = {
              details: transferDetails,
              expectedNhans: expectedNhans,
              expectedChuyens: expectedChuyens
          };
      }

      // 2. Nếu KHÔNG thấy báo cáo hôm nay -> Tạo sẵn
      if (!result.found) {
        const defaultDoiTuong = [
            "KHÁM CHỮA BỆNH CHUNG",
            "Tai nạn giao thông",
            "COVID-19",
            "Các đối tượng người bệnh khác (không gồm các đối tượng trên)"
        ];        
        
        if (Object.keys(prevDataMap).length > 0) {
             result.failButFoundPrev = true;
             result.thongKe = defaultDoiTuong.map(dt => ({
                doiTuong: dt,
                c1: prevDataMap[autoCheck(dt)] || 0,
                c21: 0, c22: 0, c3: 0, c32: 0,
                c41: 0, c42: 0, c43: 0,
                c51: 0, c52: 0, c61: 0, c62: 0, c71: 0, c72: 0
             }));
        } else if (Object.keys(expectedNhans).length > 0) {
             // Có nhận chuyển khoa nhưng chưa có data hôm qua
             result.failButFoundPrev = true;
             result.thongKe = defaultDoiTuong.map(dt => ({
                doiTuong: dt,
                c1: 0, c21: 0, c22: 0, c3: 0, c32: 0,
                c41: 0, c42: 0, c43: 0,
                c51: 0, c52: 0, c61: 0, c62: 0, c71: 0, c72: 0
             }));
        }
      }
    }

    // Đọc Data_DanhSachCa
    const sheetBN = ss.getSheetByName("Data_DanhSachCa");
    if (sheetBN && sheetBN.getLastRow() > 1) {
      const dataBN = sheetBN.getDataRange().getValues();
      for (let i = 1; i < dataBN.length; i++) {
        const trangThai = dataBN[i][40];
        if (trangThai === "Lịch sử") continue;

        const rowDate = Utilities.formatDate(new Date(dataBN[i][1]), "GMT+7", "yyyy-MM-dd");
        if (dataBN[i][2] === khoa && rowDate === ngay) {
          const formatDT = (val) => {
            if (!val) return '';
            try {
              const d = new Date(val);
              return Utilities.formatDate(d, 'GMT+7', "yyyy-MM-dd'T'HH:mm");
            } catch (e) { return String(val); }
          };
          result.benhNhan.push({
            maHSBA: dataBN[i][3] || '',
            maBN: dataBN[i][4] || '',
            ten: dataBN[i][5] || '',
            ngaySinh: dataBN[i][6] ? Utilities.formatDate(new Date(dataBN[i][6]), 'GMT+7', 'yyyy-MM-dd') : '',
            tuoi: dataBN[i][7] || '',
            gioi: dataBN[i][8] || '',
            danToc: dataBN[i][9] || '',
            quocTich: dataBN[i][10] || '',
            bhyt: dataBN[i][11] || '',
            tinh: dataBN[i][12] || '',
            huyen: dataBN[i][13] || '',
            xa: dataBN[i][14] || '',
            thon: dataBN[i][15] || '',
            lyDo: dataBN[i][16] || '',
            ngayTaiNan: formatDT(dataBN[i][17]),
            ngayVao: formatDT(dataBN[i][18]),
            ngayRa: formatDT(dataBN[i][19]),
            ketQua: dataBN[i][20] || '',
            tinhTrang: dataBN[i][21] || '',
            chuyenTu: dataBN[i][22] || '',
            noiPhatHien: dataBN[i][23] || '',
            chanDoan: dataBN[i][24] || '',
            icd: dataBN[i][25] || '',
            fileHSBAUrl: dataBN[i][26] || '',
            // New columns reading
            loaiGiayTo: dataBN[i][27] || '', 
            soGiayTo: dataBN[i][28] || '', 
            ngayCap: dataBN[i][29] ? Utilities.formatDate(new Date(dataBN[i][29]), 'GMT+7', 'yyyy-MM-dd') : '', 
            noiCap: dataBN[i][30] || '',
            htTinh: dataBN[i][31] || '', 
            htHuyen: dataBN[i][32] || '', 
            htXa: dataBN[i][33] || '', 
            htThon: dataBN[i][34] || '',
            phTinh: dataBN[i][35] || '', 
            phHuyen: dataBN[i][36] || '', 
            phXa: dataBN[i][37] || '', 
            phThon: dataBN[i][38] || '',
            chanDoanChiTiet: dataBN[i][39] || ''
          });
        }
      }
    }

    return result;
  } catch (e) {
    return { error: e.toString(), found: false };
  }
}

// ======================================================================
// PHẦN 3: TỔNG HỢP DỮ LIỆU (ĐÃ SỬA LỖI)
// ======================================================================

/**
 * Tổng hợp dữ liệu theo dải ngày — ĐÃ SỬA 3 LỖI LOGIC:
 * - Lỗi 2: Cột 7.2 giờ chỉ lấy ngày cuối (snapshot), không cộng dồn
 * - Lỗi 3: Cột 7.1 dùng ngày cuối thực tế (maxDateFound) thay vì endDate cứng
 * - Lỗi 4: Cột BN cũ (1) dùng ngày đầu thực tế (minDateFound) thay vì startDate cứng
 * (Lỗi 1 trùng lặp đã được sửa bằng cơ chế upsert ở saveReport)
 */
function getAggregatedReportRange(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data_ThongKe");
    if (!sheet) return { khoaDaNop: [], thongKe: [], chiSo: {} };

    let data = sheet.getDataRange().getValues();
    data.shift(); // Loại bỏ header

    // --- BẢO ĐẢM CHỈ LẤY DỮ LIỆU MỚI NHẤT CỦA MỖI (KHOA, NGÀY) ---
    let latestTsTK = {};
    data.forEach((row) => {
      const rowDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
      const khoa = row[2];
      const ts = new Date(row[0]).getTime();
      const key = khoa + "_" + rowDate;
      if (!latestTsTK[key] || ts > latestTsTK[key]) latestTsTK[key] = ts;
    });

    data = data.filter((row) => {
      const trangThai = row[22];
      if (trangThai === "Lịch sử") return false;
      const rowDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
      const khoa = row[2];
      const ts = new Date(row[0]).getTime();
      const key = khoa + "_" + rowDate;
      return ts === latestTsTK[key];
    });

    const ROW_NAMES = [
      "KHÁM CHỮA BỆNH CHUNG",
      "Tai nạn giao thông",
      "COVID-19",
      "Các đối tượng người bệnh khác (không gồm các đối tượng trên)"
    ];
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

    // === Bước 2: Duyệt và tổng hợp ===
    data.forEach((row) => {
      const rowDate = Utilities.formatDate(
        new Date(row[1]), "GMT+7", "yyyy-MM-dd"
      );

      if (rowDate >= startDate && rowDate <= endDate) {
        const khoa = row[2];
        const doiTuong = row[3];
        const idx = ROW_NAMES.indexOf(doiTuong);

        if (!report.khoaDaNop.includes(khoa)) report.khoaDaNop.push(khoa);

        if (idx !== -1) {
          // Các cột CỘNG DỒN: 2.1, 2.2, 3, 4.1, 4.2, 5.1, 5.2, 6.1, 6.2
          // Index trong row: 5 (2.1), 6 (2.2), 7 (3.1), 8 (4.1), 9 (4.2), 10 (5.1), 11 (5.2), 12 (6.1), 13 (6.2)
          const dynamicColIndices = [5, 6, 7, 8, 9, 10, 11, 12, 13];
          const dynamicValsIndices = [1, 2, 3, 4, 5, 6, 7, 8, 9];

          dynamicColIndices.forEach((colIdx, i) => {
            report.thongKe[idx].vals[dynamicValsIndices[i]] += parseFloat(
              row[colIdx] || 0,
            );
          });

          // BN cũ (1) — chỉ lấy từ ngày ĐẦU TIÊN thực tế có data
          if (rowDate === minDateFound) {
            report.thongKe[idx].vals[0] += parseFloat(row[4] || 0);
          }
          // BN hiện có (7.1) + Ca nặng (7.2) — chỉ lấy từ ngày CUỐI CÙNG thực tế
          if (rowDate === maxDateFound) {
            report.thongKe[idx].vals[10] += parseFloat(row[14] || 0);
            report.thongKe[idx].vals[11] += parseFloat(row[15] || 0);
          }
        }

        // Chỉ số chuyên biệt — chống trùng bằng key khoa_ngày
        const key = khoa + "_" + rowDate;
        if (!departmentDates[key]) {
          departmentDates[key] = {
            pt: parseFloat(row[16] || 0),
            ptcc: parseFloat(row[17] || 0),
            mau: parseFloat(row[18] || 0),
            tre: parseFloat(row[19] || 0),
            mo: parseFloat(row[20] || 0),
            sdt: row[21] || ''
          };
        }
      }
    });

    // === Bước 3: Đảm bảo hàng Tổng cộng (index 0) là tổng của các hàng chi tiết (1, 2, 3) ===
    // Điều này quan trọng vì báo cáo tổng hợp toàn viện không nên lấy trực tiếp hàng tổng từ từng khoa
    for (let c = 0; c < 12; c++) {
      report.thongKe[0].vals[c] = (report.thongKe[1].vals[c] || 0) + 
                                  (report.thongKe[2].vals[c] || 0) + 
                                  (report.thongKe[3].vals[c] || 0);
    }

    // === Bước 4: Lấy dữ liệu danh sách bệnh nhân ===
    const sheetBN = ss.getSheetByName("Data_DanhSachCa");
    report.benhNhan = [];

    if (sheetBN && sheetBN.getLastRow() > 1) {
      let dataBN = sheetBN.getDataRange().getValues();
      dataBN.shift(); // Bỏ header

      let latestTsBN = {};
      dataBN.forEach((row) => {
        const bnDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
        const khoa = row[2];
        const ts = new Date(row[0]).getTime();
        const key = khoa + "_" + bnDate;
        if (!latestTsBN[key] || ts > latestTsBN[key]) latestTsBN[key] = ts;
      });

      const formatDT = (val) => {
        if (!val) return '';
        try { return Utilities.formatDate(new Date(val), 'GMT+7', "yyyy-MM-dd HH:mm"); }
        catch (e) { return String(val); }
      };
      dataBN.forEach(row => {
        const bnDate = Utilities.formatDate(new Date(row[1]), "GMT+7", "yyyy-MM-dd");
        const khoa = row[2];
        const ts = new Date(row[0]).getTime();
        const key = khoa + "_" + bnDate;

        if (ts === latestTsBN[key] && bnDate >= startDate && bnDate <= endDate) {
          report.benhNhan.push({
            ngay: bnDate,
            khoa: row[2],
            maHSBA: row[3] || '',
            maBN: row[4] || '',
            ten: row[5] || '',
            ngaySinh: row[6] ? Utilities.formatDate(new Date(row[6]), 'GMT+7', 'yyyy-MM-dd') : '',
            tuoi: row[7] || '',
            gioi: row[8] || '',
            danToc: row[9] || '',
            quocTich: row[10] || '',
            bhyt: row[11] || '',
            tinh: row[12] || '',
            huyen: row[13] || '',
            xa: row[14] || '',
            thon: row[15] || '',
            lyDo: row[16] || '',
            ngayTaiNan: formatDT(row[17]),
            ngayVao: formatDT(row[18]),
            ngayRa: formatDT(row[19]),
            ketQua: row[20] || '',
            tinhTrang: row[21] || '',
            chuyenTu: row[22] || '',
            noiPhatHien: row[23] || '',
            chanDoan: row[24] || '',
            icd: row[25] || '',
            fileHSBAUrl: row[26] || '',
            // New columns reading
            loaiGiayTo: row[27] || '', 
            soGiayTo: row[28] || '', 
            ngayCap: row[29] ? Utilities.formatDate(new Date(row[29]), 'GMT+7', 'yyyy-MM-dd') : '', 
            noiCap: row[30] || '',
            htTinh: row[31] || '', 
            htHuyen: row[32] || '', 
            htXa: row[33] || '', 
            htThon: row[34] || '',
            phTinh: row[35] || '', 
            phHuyen: row[36] || '', 
            phXa: row[37] || '', 
            phThon: row[38] || '',
            chanDoanChiTiet: row[39] || ''
          });
        }
      });
    }

    // === Bước 4: Tổng hợp chỉ số chuyên biệt ===
    report.sdts = [];
    Object.keys(departmentDates).forEach((key) => {
      const parts = key.split("_");
      const d = parts[1];
      const deptData = departmentDates[key];

      report.chiSo.pt += deptData.pt;
      report.chiSo.ptcc += deptData.ptcc;
      report.chiSo.tre += deptData.tre;
      report.chiSo.mo += deptData.mo;

      if (deptData.sdt && !report.sdts.includes(deptData.sdt)) {
        report.sdts.push(deptData.sdt);
      }

      // Máu chỉ cộng của các khoa trong ngày cuối cùng có dữ liệu thực tế
      if (d === maxDateFound) {
        report.chiSo.mau += deptData.mau;
      }
    });

    return report;
  } catch (e) {
    return { error: e.toString() };
  }
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data_ThongKe");

    // Danh sách TẤT CẢ các khoa cần theo dõi
    const ALL_DEPARTMENTS = [
      "Nội tổng hợp", "Ngoại tổng hợp",
      "Phụ Sản", "Nhi", "Liên chuyên khoa",
      "Khám bệnh", "Phòng cấp cứu", "Hồi sức Cấp cứu"
    ];

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
        if (rowDate >= startDate && rowDate <= endDate) {
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

    // Lấy trạng thái duyệt
    let approvalMap = {};
    const sheetDuyet = ss.getSheetByName("Data_DuyetNgay");
    if (sheetDuyet && sheetDuyet.getLastRow() > 1) {
      const dataDuyet = sheetDuyet.getDataRange().getValues();
      for (let i = 1; i < dataDuyet.length; i++) {
        let rowDate = '';
        if (dataDuyet[i][0]) {
           try { rowDate = Utilities.formatDate(new Date(dataDuyet[i][0]), "GMT+7", "yyyy-MM-dd"); } 
           catch(e) { rowDate = String(dataDuyet[i][0]); }
        }
        if (rowDate >= startDate && rowDate <= endDate) {
          approvalMap[rowDate] = (dataDuyet[i][1] === "Đã duyệt");
        }
      }
    }

    return { dates: dates, departments: departments, approvalMap: approvalMap };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ======================================================================
// PHẦN 5: TẠO SHEET BÁO CÁO VẬT LÝ
// ======================================================================

/**
 * Hàm làm mới sheet vật lý "Báo cáo tổng" theo khoảng thời gian
 */
function refreshPhysicalSummarySheetRange(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetSummary =
    ss.getSheetByName("Báo cáo tổng") || ss.insertSheet("Báo cáo tổng");
  sheetSummary.clear();

  const data = getAggregatedReportRange(startDate, endDate);
  if (!data || !data.khoaDaNop || data.khoaDaNop.length === 0) return;

  const header1 = [
    "TT", "Khám, cấp cứu", "BN cũ (1)", "Khám bệnh (2)", "",
    "Vào viện (3.1)", "Chuyển viện (4)", "",
    "Ra viện (5)", "", "Tử vong (6)", "",
    "*BN hiện có tại thời điểm gửi báo cáo (7)", "",
  ];
  const header2 = [
    "", "", "(1)", "Tổng số (2.1)", "Khám BHYT (2.2)",
    "Tổng số (3.1)", "Ngoại trú (4.1)", "Nội trú (4.2)",
    "Tổng số (5.1)", "Tiên lượng TV xin về (5.2)",
    "Tử vong nội viện (6.1)", "Tử vong ngoại viện (6.2)",
    "Tổng số (7.1)", "Ca nặng, nguy kịch (7.2)",
  ];

  sheetSummary
    .getRange(1, 1, 1, 14)
    .setValues([header1])
    .setBackground("#cfe2f3")
    .setFontWeight("bold");
  sheetSummary
    .getRange(2, 1, 1, 14)
    .setValues([header2])
    .setBackground("#cfe2f3")
    .setFontWeight("bold");

  const merges = [
    [1, 1, 2, 1], [1, 2, 2, 1], [1, 4, 1, 2],
    [1, 7, 1, 2], [1, 9, 1, 2], [1, 11, 1, 2], [1, 13, 1, 2],
  ];
  merges.forEach((m) => sheetSummary.getRange(m[0], m[1], m[2], m[3]).merge());

  let rows = [];
  // data.thongKe[0] chính là hàng KHÁM CHỮA BỆNH CHUNG (tổng cộng)
  const totalItem = data.thongKe[0];
  
  // Chỉ lặp qua các hàng chi tiết (bỏ qua hàng đầu tiên)
  for (let i = 1; i < data.thongKe.length; i++) {
    const item = data.thongKe[i];
    let r = [i, item.doiTuong];
    item.vals.forEach((v) => r.push(v));
    rows.push(r);
  }

  // Thêm hàng TỔNG CỘNG ở cuối cùng lấy từ data.thongKe[0]
  if (totalItem) {
    rows.push(["Σ", "TỔNG CỘNG (KHÁM CHỮA BỆNH CHUNG)", ...totalItem.vals]);
  }

  sheetSummary
    .getRange(3, 1, rows.length, 14)
    .setValues(rows)
    .setBorder(true, true, true, true, true, true);
  
  // Định dạng hàng tổng cộng cuối cùng
  sheetSummary
    .getRange(3 + rows.length - 1, 1, 1, 14)
    .setFontWeight("bold")
    .setBackground("#d9ead3") // Màu xanh lá nhạt phân biệt hàng tổng
    .setFontColor("#274e13");

  const startRow = 3 + rows.length + 2;
  const subData = [
    ["Chỉ tiêu", "Số lượng", "Ghi chú"],
    ["Số ca phẫu thuật (loại 3 trở lên)", data.chiSo.pt, "Tổng trong khoảng"],
    ["- Trong đó: Cấp cứu do tai nạn", data.chiSo.ptcc, ""],
    ["Tổng trẻ sinh tại CSKCB", data.chiSo.tre, ""],
    ["- Trong đó: Sinh mổ", data.chiSo.mo, ""],
    ["Tổng lượng máu dự trữ tại BV (ml)", data.chiSo.mau, "Dữ liệu ngày cuối"],
    ["Khoảng thời gian", startDate + " đến " + endDate, ""],
    ["Khoa đã nộp", data.khoaDaNop.join(", "), ""],
  ];
  sheetSummary
    .getRange(startRow, 1, subData.length, 3)
    .setValues(subData)
    .setBorder(true, true, true, true, true, true);
  sheetSummary.setColumnWidth(2, 250);
  sheetSummary
    .getRange("A1:N50")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

/**
 * HÀM CHECK QUYỀN DRIVE (Chạy thủ công để cấp quyền)
 */
function checkDrivePermissions() {
  const folderName = 'Báo cáo tết';
  try {
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      Logger.log('OK: Tìm thấy folder ' + folderName);
    } else {
      DriveApp.createFolder(folderName);
      Logger.log('OK: Đã tạo folder ' + folderName);
    }
    return "Kết nối Drive thành công! Vui lòng thử lại báo cáo.";
  } catch (e) {
    Logger.log('Lỗi: ' + e.toString());
    return "Lỗi quyền Drive: " + e.toString();
  }
}
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
        case 'toggleApproval':
          result = toggleApproval(payload.ngay, payload.isApproved);
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

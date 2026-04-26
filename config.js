/**
 * AUTO-GENERATED CONFIG — Không chỉnh sửa trực tiếp
 * Nguồn: Google Sheet Config → GAS publishConfig → GitHub API
 * Cập nhật lúc: 2026-04-26T10:00:00+07:00 (bản khởi tạo)
 */
window.APP_CONFIG = {
  "version": "2026-04-26T10:00:00",

  "khoaPhong": [
    { "ma": "NOI_TH", "ten": "Nội tổng hợp", "nhom": "LAM_SANG", "sections": ["sectionNoiTru"], "disableCols": [21,22,41,62] },
    { "ma": "NGOAI_TH", "ten": "Ngoại tổng hợp", "nhom": "LAM_SANG", "sections": ["sectionNoiTru"], "disableCols": [21,22,41,62] },
    { "ma": "PHU_SAN", "ten": "Phụ Sản", "nhom": "LAM_SANG", "sections": ["sectionNoiTru","sectionSanKhoa"], "disableCols": [21,22,41,62] },
    { "ma": "NHI", "ten": "Nhi", "nhom": "LAM_SANG", "sections": ["sectionNoiTru"], "disableCols": [21,22,41,62] },
    { "ma": "LIEN_CK", "ten": "Liên chuyên khoa", "nhom": "LAM_SANG", "sections": ["sectionNoiTru"], "disableCols": [21,22,41,62] },
    { "ma": "KHAM_BENH", "ten": "Khám bệnh", "nhom": "NGOAI_TRU", "sections": ["sectionNoiTru"], "disableCols": [1,3,42,51,52,61,62,72] },
    { "ma": "PHONG_CC", "ten": "Phòng cấp cứu", "nhom": "NGOAI_TRU", "sections": ["sectionNoiTru"], "disableCols": [1,3,42,51,52,72] },
    { "ma": "HSCC", "ten": "Hồi sức Cấp cứu", "nhom": "LAM_SANG", "sections": ["sectionNoiTru"], "disableCols": [21,22,41,62] }
  ],

  "doiTuong": [
    { "stt": 0, "ten": "KHÁM CHỮA BỆNH CHUNG", "laDongTong": true },
    { "stt": 1, "ten": "Tai nạn giao thông", "laDongTong": false },
    { "stt": 2, "ten": "COVID-19", "laDongTong": false },
    { "stt": 3, "ten": "Các đối tượng người bệnh khác (không gồm các đối tượng trên)", "laDongTong": false }
  ],

  "bangThongKe": {
    "colCodes": [1, 21, 22, 3, 41, 42, 51, 52, 61, 62],
    "computedCols": {
      "c71": "c1+c3-c42-c51-c61-c62"
    }
  },

  "chiSoKhac": [
    { "ma": "phauThuat", "ten": "Số ca phẫu thuật (loại 3 trở lên)", "inputId": "phauThuat", "section": "sectionPhauThuat", "icon": "fa-scalpel", "mauNen": "orange", "donVi": "ca" },
    { "ma": "phauThuatCC", "ten": "Trong đó, PT cấp cứu do tai nạn", "inputId": "phauThuatCC", "section": "sectionPhauThuat", "icon": "fa-scalpel", "mauNen": "orange", "donVi": "ca" },
    { "ma": "treSinh", "ten": "TS trẻ sinh tại CSKCB", "inputId": "treSinh", "section": "sectionSanKhoa", "icon": "fa-baby", "mauNen": "pink", "donVi": "ca" },
    { "ma": "treSinhMo", "ten": "Trong đó, số trẻ sinh mổ đẻ", "inputId": "treSinhMo", "section": "sectionSanKhoa", "icon": "fa-baby", "mauNen": "pink", "donVi": "ca" },
    { "ma": "mauDuTru", "ten": "Tổng số lượng máu dự trữ tại CSKCB (ml)", "inputId": "mauDuTru", "section": "sectionMau", "icon": "fa-tint", "mauNen": "red", "donVi": "ml" }
  ],

  "danhMuc": {
    "LY_DO": {
      "1": { "label": "Tai nạn giao thông", "short": "TNGT" },
      "2": { "label": "Tai nạn do pháo nổ", "short": "Pháo nổ" },
      "3": { "label": "Vũ khí, vật liệu nổ tự chế", "short": "VK, VLN" },
      "4": { "label": "Ngộ độc thực phẩm", "short": "NĐTP" },
      "5": { "label": "Đối tượng khác", "short": "Khác" }
    },
    "KET_QUA": {
      "1": { "label": "Đang điều trị", "short": "Đang ĐT" },
      "2": { "label": "Khỏi ra viện", "short": "Khỏi" },
      "3": { "label": "Chuyển viện", "short": "Chuyển viện" },
      "4": { "label": "Tử vong tại CSKCB", "short": "TV tại CSKCB" },
      "5": { "label": "Tử vong trước khi đến CSKCB", "short": "TV trước CSKCB" },
      "6": { "label": "Tiên lượng nặng xin về", "short": "Nặng xin về" }
    },
    "TINH_TRANG": {
      "Ra viện": { "label": "Ra viện", "short": "" },
      "Xin ra viện": { "label": "Xin ra viện", "short": "" },
      "Trốn viện": { "label": "Trốn viện", "short": "" },
      "Chuyển viện": { "label": "Chuyển viện", "short": "" },
      "Khác": { "label": "Khác", "short": "" },
      "TV có khám nghiệm": { "label": "Tử vong có khám nghiệm tử thi", "short": "" },
      "TV không khám nghiệm": { "label": "Tử vong không khám nghiệm tử thi", "short": "" }
    },
    "LOAI_GIAY_TO": {
      "SDCN": { "label": "Số định danh cá nhân", "short": "" },
      "CCCD": { "label": "Căn cước công dân", "short": "" },
      "CMND": { "label": "Chứng minh nhân dân", "short": "" },
      "HC": { "label": "Hộ chiếu", "short": "" },
      "GCS": { "label": "Giấy chứng sinh", "short": "" },
      "BLX": { "label": "Bằng lái xe", "short": "" },
      "KHAC": { "label": "Khác", "short": "" },
      "KHONG": { "label": "Không có", "short": "" }
    },
    "LOAI_CHAN_DOAN": {
      "BC": { "label": "Bệnh chính", "short": "" },
      "BienChung": { "label": "Biến chứng", "short": "" },
      "BenhKem": { "label": "Bệnh kèm", "short": "" },
      "NguyenNhan": { "label": "Nguyên nhân", "short": "" },
      "TacNhan": { "label": "Tác nhân", "short": "" },
      "Khac": { "label": "Khác", "short": "" }
    }
  },

  "app": {
    "APP_TITLE": "Báo cáo trực lễ 30/4-1/5",
    "REPORT_TITLE": "BÁO CÁO NHANH TRỰC TUYẾN",
    "REPORT_SUBTITLE": "TÌNH HÌNH KHÁM CHỮA BỆNH, CẤP CỨU, TAI NẠN, NGỘ ĐỘC",
    "DATE_LOCK_PASSWORD": "cotec@123",
    "FOOTER_TEXT": "Phòng Kế hoạch Tổng hợp - Bệnh viện Bình Định @ 2026"
  }
};

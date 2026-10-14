// ĐIỀN ID FILE GOOGLE SHEETS CỦA BẠN VÀO GIỮA CẶP DẤU NHÁY ĐƠN DƯỚI ĐÂY
// (Cách lấy: Copy chuỗi ký tự dài nằm giữa chữ /d/ và /edit trên thanh địa chỉ của file Excel Google)
const SPREADSHEET_ID = '1g3Lc4RCiE23EzjLT1yshjUHX5GsYOJWgmJJn7ultesA';

// 1. XỬ LÝ TRA CỨU THÔNG TIN VĐV KHI NHẬP CCCD
function doGet(e) {
  var cccdTraCuu = e.parameter.cccd;
  
  // Mở chính xác file Google Sheets bằng ID cấu hình phía trên
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetVDV = ss.getSheetByName("DanhSachVDV");
  var data = sheetVDV.getDataRange().getValues();
  
  // Tìm kiếm CCCD trong bảng dữ liệu (Bỏ qua dòng tiêu đề hàng 1)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === cccdTraCuu.trim()) {
      var result = {
        found: true,
        name: data[i][1],
        dept: data[i][2]
      };
      return ContentService.createTextOutput(JSON.stringify(result))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // Nếu không tìm thấy số CCCD trùng khớp
  return ContentService.createTextOutput(JSON.stringify({ found: false }))
                       .setMimeType(ContentService.MimeType.JSON);
}

// 2. XỬ LÝ LƯU THÔNG TIN NGƯỜI THÂN ĐĂNG KÝ
function doPost(e) {
  // Mở chính xác file Google Sheets bằng ID cấu hình phía trên
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetDangKy = ss.getSheetByName("DangKyNguoiThan");
  
  try {
    var data = JSON.parse(e.postData.contents);
    var timestamp = new Date().toLocaleString("vi-VN");

    data.relatives.forEach(function(relative) {
      sheetDangKy.appendRow([
        timestamp,
        "'" + data.cccd, // Dấu nháy đơn giữ định dạng chuỗi số không bị mất số 0 đầu
        data.athleteName,
        data.athleteDept,
        relative.name,
        relative.relation
      ]);
    });

    return ContentService.createTextOutput(JSON.stringify({ "status": "success" }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

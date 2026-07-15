// Hàm khởi tạo Web App (Chạy đầu tiên khi load link)
// Đã tối ưu UI/UX: Loại bỏ các lệnh gọi Spreadsheet không cần thiết để tăng tốc độ render HTML < 1s
function doGet(e) {
  try {
    // Lưu ý: KHÔNG gọi initDatabase() ở đây để tránh làm chậm tốc độ khởi động App.
    // Việc khởi tạo Database chỉ thực hiện 1 lần duy nhất bằng cách chạy thủ công hàm initDatabase() trong trình soạn thảo GAS.
    
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('HKD ĐÀO THANH YẾN')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>Đã xảy ra lỗi khởi tạo hệ thống: ' + error.message + '</h1>');
  }
}

// Hàm nhúng các file HTML phụ vào file Index tổng
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
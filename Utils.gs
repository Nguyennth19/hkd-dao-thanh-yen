/* File: Utils.gs */

// Hàm tiện ích: Đọc giá trị cấu hình từ Sheet ThietLap
function getSetting(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ThietLap');
  if(!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for(let i = 0; i < data.length; i++){
    if(data[i][0] === key) return data[i][1];
  }
  return null;
}

// API Tiện ích: Kiểm tra Mã PIN đăng nhập
function verifyPinCode(pin) {
  const realPin = getSetting('PinCode');
  // So sánh chuỗi bảo đảm an toàn
  if (realPin && realPin.toString() === pin.toString()) {
    return { success: true };
  }
  return { success: false, message: 'Mã PIN không chính xác hoặc chưa được thiết lập!' };
}

// API Tiện ích: Xử lý Upload file từ base64 lên Google Drive
function uploadFileToDrive(base64Data, fileName, type) {
  try {
    // Phân luồng thư mục tương ứng theo loại Form lấy từ Cấu hình động
    let folderId = getSetting(type === 'Thu' ? 'DriveFolderThu' : 'DriveFolderChi');
    
    // Fallback: Đề phòng lỡ tay xóa cấu hình trong Sheet
    if (!folderId) {
      folderId = (type === 'Thu') ? '1jXP7THvA74s1HRlw7T68BZMG80sck8az' : '1Dz1m9S50__cJ5bTmyCkmyZVVE0dAcGVH';
    }
    
    let folder = DriveApp.getFolderById(folderId);
    
    // Tách phần Header của Base64
    let splitData = base64Data.split(',');
    let contentType = splitData[0].match(/:(.*?);/)[1];
    let bytes = Utilities.base64Decode(splitData[1]);
    let blob = Utilities.newBlob(bytes, contentType, fileName);
    
    // Tạo file và set quyền View
    let file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
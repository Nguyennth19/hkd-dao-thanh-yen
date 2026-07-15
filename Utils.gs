/* File: Utils.gs */

// API Tiện ích: Xử lý Upload file từ base64 lên Google Drive
function uploadFileToDrive(base64Data, fileName, type) {
  try {
    // Phân luồng thư mục tương ứng theo loại Form
    let folderId = (type === 'Thu') ? '1jXP7THvA74s1HRlw7T68BZMG80sck8az' : '1Dz1m9S50__cJ5bTmyCkmyZVVE0dAcGVH';
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
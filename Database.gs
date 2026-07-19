// Hàm khởi tạo CSDL: Kiểm tra và thiết lập các Sheet
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Khởi tạo Sheet Data_Thu
  let thuSheet = ss.getSheetByName('Data_Thu');
  if (!thuSheet) {
    thuSheet = ss.insertSheet('Data_Thu');
    thuSheet.appendRow(['BẢNG THEO DÕI HÓA ĐƠN BÁN RA / CÔNG NỢ KHÁCH HÀNG', '', '', '', '', '', '', '', '', '', '', '']);
    thuSheet.getRange("A1:L1").merge().setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14).setFontColor("#1e3a8a");
    thuSheet.appendRow(['Mã Hệ Thống', 'Ngày Giao Dịch', 'Khoản Thu Từ Đâu', 'Mã số thuế', 'Số Hóa Đơn', 'Nội dung hàng hóa/dịch vụ', 'Số Tiền Ban Đầu', 'Tiền Thuế (Nếu có)', 'Tiền Thực Nhận', 'Ghi Chú Chi Tiết', 'Hình Ảnh / Chứng Từ', 'Thời Gian Ghi Sổ']);
    thuSheet.getRange("A2:L2").setFontWeight("bold").setBackground("#3f6212").setFontColor("white");
    thuSheet.setFrozenRows(2);
  }
  
  // 2. Khởi tạo Sheet Data_Chi
  let chiSheet = ss.getSheetByName('Data_Chi');
  if (!chiSheet) {
    chiSheet = ss.insertSheet('Data_Chi');
    chiSheet.appendRow(['BẢNG THEO DÕI HÓA ĐƠN MUA VÀO (CHI PHÍ LƯU NỘI BỘ)', '', '', '', '', '', '', '', '', '', '', '']);
    chiSheet.getRange("A1:L1").merge().setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14).setFontColor("#1e3a8a");
    chiSheet.appendRow(['Mã Hệ Thống', 'Ngày Giao Dịch', 'Người Nhận / Đối Tác', 'Mã số thuế', 'Số HĐ/Biên lai', 'Nội dung hàng hóa/dịch vụ', 'Số Tiền Cần Chi', 'Tiền Thuế (Nếu có)', 'Tiền Thực Chi', 'Ghi Chú Chi Tiết', 'Hình Ảnh / Chứng Từ', 'Thời Gian Ghi Sổ']);
    chiSheet.getRange("A2:L2").setFontWeight("bold").setBackground("#3f6212").setFontColor("white");
    chiSheet.setFrozenRows(2);
  }
  
  // 3. Khởi tạo Sheet Danh_Muc (Cấu hình động)
  let dmSheet = ss.getSheetByName('Danh_Muc');
  if (!dmSheet) {
    dmSheet = ss.insertSheet('Danh_Muc');
    dmSheet.appendRow(['Nhóm', 'Tên Danh Mục']);
    dmSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#3f6212").setFontColor("white");
    
    const macDinh = [
      ['Thu', 'Tiktok shop'],
      ['Thu', 'Shopee'],
      ['Thu', 'Facebook'],
      ['Chi', 'Lương'],
      ['Chi', 'Thuê nhà'],
      ['Chi', 'Affiliate'],
      ['Chi', 'Thuế'],
      ['Chi', 'Vật dụng'],
      ['Chi', 'Khác']
    ];
    dmSheet.getRange(2, 1, macDinh.length, 2).setValues(macDinh);
    dmSheet.setFrozenRows(1);
  }
  
  // 4. Khởi tạo Sheet Thiet_Lap
  let settingsSheet = ss.getSheetByName('Thiet_Lap');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Thiet_Lap');
    settingsSheet.appendRow(['Tên Cấu Hình', 'Nội Dung Cài Đặt']);
    settingsSheet.appendRow(['TienTe', 'VND']);
    settingsSheet.appendRow(['MucTieuLoiNhuan', '235000000']);
    settingsSheet.appendRow(['PinCode', '123456']);
    settingsSheet.appendRow(['DriveFolderThu', '1jXP7THvA74s1HRlw7T68BZMG80sck8az']);
    settingsSheet.appendRow(['DriveFolderChi', '1Dz1m9S50__cJ5bTmyCkmyZVVE0dAcGVH']);
    settingsSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#3f6212").setFontColor("white");
    settingsSheet.setFrozenRows(1);
  }
  
  // 5. Khởi tạo Sheet Log_HeThong
  let logSheet = ss.getSheetByName('Log_HeThong');
  if (!logSheet) {
    logSheet = ss.insertSheet('Log_HeThong');
    logSheet.appendRow(['Thời Gian', 'Hành Động', 'Loại Giao Dịch', 'Nội Dung Chi Tiết', 'Trạng Thái']);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#3f6212").setFontColor("white");
    logSheet.setFrozenRows(1);
  }
}

// API: Lấy danh mục động cho Dropdown Frontend
function getDropdownCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Danh_Muc');
    if (!sheet) return { success: false, data: { thu: [], chi: [] } };
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Bỏ dòng tiêu đề (Dòng 1)
    
    let thu = [];
    let chi = [];
    
    data.forEach(row => {
      if (row[0] === 'Thu' && row[1] !== 'Khác') thu.push(row[1]);
      if (row[0] === 'Chi' && row[1] !== 'Khác') chi.push(row[1]);
    });
    
    thu.push('Khác');
    chi.push('Khác');
    
    return { success: true, data: { thu: thu, chi: chi } };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
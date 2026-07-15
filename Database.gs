// Hàm khởi tạo CSDL: Kiểm tra và thiết lập các Sheet
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Khởi tạo/Cập nhật Sheet GiaoDich
  let transSheet = ss.getSheetByName('GiaoDich');
  if (!transSheet) {
    transSheet = ss.insertSheet('GiaoDich');
    transSheet.appendRow(['ID', 'Ngày', 'Loại', 'Danh Mục', 'Số Tiền', 'Thuế', 'Thực Tế', 'Ghi Chú', 'Tài Liệu', 'Thời Gian Tạo']);
    transSheet.getRange("A1:J1").setFontWeight("bold").setBackground("#e5e7eb");
    transSheet.setFrozenRows(1);
    transSheet.setColumnWidth(9, 180);
  } else {
    let headers = transSheet.getRange(1, 1, 1, transSheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Tài Liệu')) {
      transSheet.insertColumnBefore(transSheet.getLastColumn());
      transSheet.getRange(1, transSheet.getLastColumn() - 1).setValue('Tài Liệu').setFontWeight("bold");
    }
  }
  
  // 2. Khởi tạo Sheet DanhMuc (Cấu hình động)
  let dmSheet = ss.getSheetByName('DanhMuc');
  if (!dmSheet) {
    dmSheet = ss.insertSheet('DanhMuc');
    dmSheet.appendRow(['Loại', 'Tên Danh Mục']);
    dmSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#e5e7eb");
    
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
  
  // 3. Khởi tạo Sheet ThietLap
  let settingsSheet = ss.getSheetByName('ThietLap');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('ThietLap');
    settingsSheet.appendRow(['Khóa', 'Giá trị']);
    settingsSheet.appendRow(['TienTe', 'VND']);
    settingsSheet.appendRow(['MucTieuLoiNhuan', '235000000']);
    settingsSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#e5e7eb");
    settingsSheet.setFrozenRows(1);
  }
}

// API: Lấy danh mục động cho Dropdown Frontend
function getDropdownCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DanhMuc');
    if (!sheet) return { success: false, data: { thu: [], chi: [] } };
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Bỏ dòng tiêu đề
    
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
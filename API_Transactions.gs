function saveTransaction(record) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('GiaoDich');
    
    const id = Utilities.getUuid().split('-')[0];
    const timestamp = new Date();
    
    let amount = parseFloat(record.amount) || 0;
    let tax = parseFloat(record.tax) || 0;
    let net = amount - tax; 
    
    let finalCategory = record.category;
    
    if (record.category === 'Khác' && record.customCategory && record.customCategory.trim() !== '') {
      finalCategory = record.customCategory.trim();
      const dmSheet = ss.getSheetByName('DanhMuc');
      
      let currentDmData = dmSheet.getDataRange().getValues();
      let daTonTai = currentDmData.some(row => row[0] === record.type && row[1].toLowerCase() === finalCategory.toLowerCase());
      
      if (!daTonTai) {
        dmSheet.appendRow([record.type, finalCategory]);
      }
    }
    
    sheet.appendRow([
      id,
      record.date,
      record.type,
      finalCategory,
      amount,
      tax,
      net,
      record.note,
      record.fileUrl || '', 
      timestamp
    ]);
    
    return { success: true, message: 'Đã lưu trữ giao dịch thành công!' };
  } catch (error) {
    return { success: false, message: 'Lỗi ghi nhận CSDL: ' + error.toString() };
  }
}

function getTransactions(typeFilter, monthFilter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('GiaoDich');
    if (!sheet) return { success: false, data: [] };

    const data = sheet.getDataRange().getValues();
    data.shift(); 
    
    let results = [];
    data.forEach(row => {
      // Logic kiểm tra và lọc theo tháng
      if (monthFilter && monthFilter !== 'all') {
        let dateObj = row[1];
        if (typeof dateObj === 'string') {
          if (!dateObj.startsWith(monthFilter)) return;
        } else if (dateObj instanceof Date) {
          let mYear = dateObj.getFullYear();
          let mMonth = ('0' + (dateObj.getMonth() + 1)).slice(-2);
          if (`${mYear}-${mMonth}` !== monthFilter) return;
        }
      }

      if(row[2] === typeFilter || typeFilter === 'All') {
        results.push({
          id: row[0],
          date: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "dd/MM/yyyy") : row[1],
          type: row[2],
          category: row[3],
          amount: row[4],
          tax: row[5],
          net: row[6],
          note: row[7],
          fileUrl: row[8] || ''
        });
      }
    });
    
    return { success: true, data: results.reverse() };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
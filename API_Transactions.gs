// GIAI ĐOẠN 1: HÀM TIỆN ÍCH GHI LOG HỆ THỐNG
function writeSystemLog(action, type, details, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Log_HeThong');
    if (sheet) {
      sheet.appendRow([new Date(), action, type, details, status]);
    }
  } catch (e) {}
}

function saveTransaction(record) {
  const lock = LockService.getScriptLock();
  try {
    // Khóa luồng tối đa 30 giây để chờ nếu có người khác đang ghi dữ liệu
    lock.waitLock(30000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheetName = record.type === 'Thu' ? 'Data_Thu' : 'Data_Chi';
    const sheet = ss.getSheetByName(targetSheetName);
    
    if (!sheet) throw new Error("Không tìm thấy Sheet " + targetSheetName);
    
    const id = Utilities.getUuid().split('-')[0];
    const timestamp = new Date();
    
    let amount = parseFloat(record.amount) || 0;
    let tax = parseFloat(record.tax) || 0;
    let net = amount - tax; 
    
    let finalCategory = record.category;
    
    if (record.category === 'Khác' && record.customCategory && record.customCategory.trim() !== '') {
      finalCategory = record.customCategory.trim();
      const dmSheet = ss.getSheetByName('Danh_Muc');
      if (dmSheet) {
        let currentDmData = dmSheet.getDataRange().getValues();
        let daTonTai = currentDmData.some(row => row[0] === record.type && row[1].toLowerCase() === finalCategory.toLowerCase());
        
        if (!daTonTai) {
          dmSheet.appendRow([record.type, finalCategory]);
        }
      }
    }
    
    // Tạo cấu trúc Hyperlink nếu có URL
    let fileLink = record.fileUrl ? `=HYPERLINK("${record.fileUrl}", "Tài liệu")` : "";
    
    // Ghi 12 cột theo chuẩn Thuế mới (Bắt đầu từ A đến L)
    sheet.appendRow([
      id,                                // Cột A: Mã Hệ Thống
      record.date,                       // Cột B: Ngày Giao Dịch
      finalCategory,                     // Cột C: Khoản Thu / Mục Đích Chi
      record.taxCode || '',              // Cột D: Mã số thuế (Sẽ thêm ở Frontend GĐ 3)
      record.invoiceNumber || '',        // Cột E: Số HĐ/Biên lai (Sẽ thêm ở Frontend GĐ 3)
      record.description || '',          // Cột F: Nội dung hàng hóa (Sẽ thêm ở Frontend GĐ 3)
      amount,                            // Cột G: Số Tiền Ban Đầu / Cần Chi
      tax,                               // Cột H: Tiền Thuế
      net,                               // Cột I: Tiền Thực Nhận / Thực Chi
      record.note || '',                 // Cột J: Ghi Chú Chi Tiết
      fileLink,                          // Cột K: Hình Ảnh / Chứng Từ
      timestamp                          // Cột L: Thời Gian Ghi Sổ
    ]);
    
    SpreadsheetApp.flush();
    
    // Ghi Log hệ thống
    writeSystemLog("Thêm mới", record.type, `Ghi nhận ${record.type} (${finalCategory}): ${net} đ`, "Thành công");
    
    return { success: true, message: 'Đã lưu trữ giao dịch thành công!' };
  } catch (error) {
    writeSystemLog("Thêm mới", record.type || "Không rõ", `Lỗi: ${error.toString()}`, "Lỗi");
    return { success: false, message: 'Lỗi ghi nhận CSDL: ' + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Hàm hỗ trợ nội bộ đọc dữ liệu 1 Sheet tối ưu
function getOptimizedSheetData(sheet, monthFilter) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { data: [], formulas: [] }; // Dữ liệu bắt đầu từ dòng 3

  let startIdx = -1;
  let endIdx = -1;
  let numRows = 0;
  
  if (!monthFilter || monthFilter === 'all') {
    startIdx = 0;
    numRows = lastRow - 2;
  } else {
    // TỐI ƯU: Thuật toán dò tìm startRow và endRow của tháng để không kéo mảng lớn vào RAM
    const dateValues = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
    
    for(let i = 0; i < dateValues.length; i++) {
      let d = dateValues[i][0];
      let match = false;
      if (typeof d === 'string' && d.startsWith(monthFilter)) match = true;
      else if (d instanceof Date) {
        let mYear = d.getFullYear();
        let mMonth = ('0' + (d.getMonth() + 1)).slice(-2);
        if (`${mYear}-${mMonth}` === monthFilter) match = true;
      }
      
      if (match) {
        if (startIdx === -1) startIdx = i;
        endIdx = i;
      }
    }
    
    if (startIdx !== -1) {
      numRows = endIdx - startIdx + 1;
    }
  }
  
  if (numRows > 0) {
    const range = sheet.getRange(startIdx + 3, 1, numRows, sheet.getLastColumn());
    return { 
      data: range.getValues(), 
      formulas: range.getFormulas() 
    };
  }
  
  return { data: [], formulas: [] };
}

function getTransactions(typeFilter, monthFilter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let results = [];
    
    const processSheetData = (sheetData, typeStr) => {
      sheetData.data.forEach((row, index) => {
        // Lọc lại lần nữa đề phòng dữ liệu nhập lộn xộn ngày tháng
        if (monthFilter && monthFilter !== 'all') {
          let dateObj = row[1];
          let match = false;
          if (typeof dateObj === 'string') {
            if (dateObj.startsWith(monthFilter)) match = true;
          } else if (dateObj instanceof Date) {
            let mYear = dateObj.getFullYear();
            let mMonth = ('0' + (dateObj.getMonth() + 1)).slice(-2);
            if (`${mYear}-${mMonth}` === monthFilter) match = true;
          }
          if (!match) return;
        }

        // Bóc tách URL từ hàm Hyperlink ở Cột K (Index 10)
        let fileUrl = '';
        let formulaK = sheetData.formulas[index][10];
        if (formulaK && formulaK.toString().includes('HYPERLINK')) {
          let match = formulaK.match(/HYPERLINK\("(.*?)"/);
          if (match) fileUrl = match[1];
        } else {
          fileUrl = row[10] || '';
        }

        results.push({
          id: row[0],
          date: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "dd/MM/yyyy") : row[1],
          type: typeStr,
          category: row[2], // Khoản thu / Đối tác
          amount: row[6],
          tax: row[7],
          net: row[8],
          note: row[9],
          fileUrl: fileUrl,
          timestamp: row[11] // Dùng để sort
        });
      });
    };

    if (typeFilter === 'Thu' || typeFilter === 'All') {
      const sheetThu = ss.getSheetByName('Data_Thu');
      if (sheetThu) processSheetData(getOptimizedSheetData(sheetThu, monthFilter), 'Thu');
    }
    
    if (typeFilter === 'Chi' || typeFilter === 'All') {
      const sheetChi = ss.getSheetByName('Data_Chi');
      if (sheetChi) processSheetData(getOptimizedSheetData(sheetChi, monthFilter), 'Chi');
    }
    
    // Sắp xếp giảm dần theo thời gian tạo (Mới nhất lên đầu)
    results.sort((a, b) => {
      let timeA = a.timestamp instanceof Date ? a.timestamp.getTime() : 0;
      let timeB = b.timestamp instanceof Date ? b.timestamp.getTime() : 0;
      return timeB - timeA;
    });
    
    return { success: true, data: results };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getRecentLogs(limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Log_HeThong');
    if (!sheet) return { success: false, data: [] };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] }; // Log bắt đầu từ dòng 2
    
    const numRows = Math.min(limit, lastRow - 1);
    const startRow = lastRow - numRows + 1;
    
    const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
    
    let results = [];
    data.forEach(row => {
      results.push({
        date: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : row[0],
        type: row[2],
        category: row[1], // Mượn trường category để chứa Hành động
        net: row[4],      // Mượn trường net để chứa Trạng thái
        note: row[3]      // Chi tiết log
      });
    });
    
    return { success: true, data: results.reverse() };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function deleteTransaction(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const deleteFromSheet = (sheetName, typeStr) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return false;
      const data = sheet.getDataRange().getValues();
      for (let i = 2; i < data.length; i++) { // Bắt đầu từ dòng 3 (index 2)
        if (data[i][0] === id) {
          let amountStr = data[i][8];
          sheet.deleteRow(i + 1);
          writeSystemLog("Xóa dữ liệu", typeStr, `Đã xóa giao dịch ${id} (${amountStr} đ)`, "Thành công");
          return true;
        }
      }
      return false;
    };
    
    if (deleteFromSheet('Data_Thu', 'Thu')) {
      SpreadsheetApp.flush();
      return { success: true, message: 'Đã xóa khoản Thu thành công!' };
    }
    
    if (deleteFromSheet('Data_Chi', 'Chi')) {
      SpreadsheetApp.flush();
      return { success: true, message: 'Đã xóa khoản Chi thành công!' };
    }
    
    return { success: false, message: 'Không tìm thấy giao dịch để xóa.' };
  } catch(e) {
    writeSystemLog("Xóa dữ liệu", "Lỗi", `Lỗi khi xóa ${id}: ${e.toString()}`, "Lỗi");
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateTransactionAmount(id, newAmount) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const updateInSheet = (sheetName, typeStr) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return false;
      const data = sheet.getDataRange().getValues();
      for (let i = 2; i < data.length; i++) { // Bắt đầu từ dòng 3 (index 2)
        if (data[i][0] === id) {
          const rowNum = i + 1;
          const tax = parseFloat(sheet.getRange(rowNum, 8).getValue()) || 0; // Tiền thuế Cột H
          const net = parseFloat(newAmount) - tax;
          
          sheet.getRange(rowNum, 7).setValue(parseFloat(newAmount)); // Tiền ban đầu Cột G
          sheet.getRange(rowNum, 9).setValue(net);                   // Tiền thực Cột I
          
          writeSystemLog("Sửa số tiền", typeStr, `Cập nhật giao dịch ${id} thành ${net} đ`, "Thành công");
          return true;
        }
      }
      return false;
    };
    
    if (updateInSheet('Data_Thu', 'Thu')) {
      SpreadsheetApp.flush();
      return { success: true, message: 'Đã cập nhật số tiền khoản Thu!' };
    }
    
    if (updateInSheet('Data_Chi', 'Chi')) {
      SpreadsheetApp.flush();
      return { success: true, message: 'Đã cập nhật số tiền khoản Chi!' };
    }
    
    return { success: false, message: 'Không tìm thấy giao dịch để cập nhật.' };
  } catch(e) {
    writeSystemLog("Sửa số tiền", "Lỗi", `Lỗi khi sửa ${id}: ${e.toString()}`, "Lỗi");
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
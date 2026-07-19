function getDashboardData(monthFilter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let tongThu = 0;
    let tongChi = 0;
    let phanBoChiPhi = {};
    
    // Đọc dữ liệu khoản THU (Từ sheet Data_Thu, dòng 3 trở đi)
    const thuSheet = ss.getSheetByName('Data_Thu');
    if (thuSheet) {
      const lastRowThu = thuSheet.getLastRow();
      if (lastRowThu >= 3) {
        let dataThu = [];
        if (!monthFilter || monthFilter === 'all') {
          dataThu = thuSheet.getRange(3, 1, lastRowThu - 2, thuSheet.getLastColumn()).getValues();
        } else {
          const dateValues = thuSheet.getRange(3, 2, lastRowThu - 2, 1).getValues();
          let startIdx = -1, endIdx = -1;
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
            dataThu = thuSheet.getRange(startIdx + 3, 1, endIdx - startIdx + 1, thuSheet.getLastColumn()).getValues();
          }
        }
        
        dataThu.forEach(row => {
          if (monthFilter && monthFilter !== 'all') {
            let dateObj = row[1];
            let match = false;
            if (typeof dateObj === 'string' && dateObj.startsWith(monthFilter)) match = true;
            else if (dateObj instanceof Date) {
              let mYear = dateObj.getFullYear();
              let mMonth = ('0' + (dateObj.getMonth() + 1)).slice(-2);
              if (`${mYear}-${mMonth}` === monthFilter) match = true;
            }
            if (!match) return;
          }
          let soTienThuc = parseFloat(row[8]) || 0; // Cột I: Tiền Thực Nhận
          tongThu += soTienThuc;
        });
      }
    }
    
    // Đọc dữ liệu khoản CHI (Từ sheet Data_Chi, dòng 3 trở đi)
    const chiSheet = ss.getSheetByName('Data_Chi');
    if (chiSheet) {
      const lastRowChi = chiSheet.getLastRow();
      if (lastRowChi >= 3) {
        let dataChi = [];
        if (!monthFilter || monthFilter === 'all') {
          dataChi = chiSheet.getRange(3, 1, lastRowChi - 2, chiSheet.getLastColumn()).getValues();
        } else {
          const dateValues = chiSheet.getRange(3, 2, lastRowChi - 2, 1).getValues();
          let startIdx = -1, endIdx = -1;
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
            dataChi = chiSheet.getRange(startIdx + 3, 1, endIdx - startIdx + 1, chiSheet.getLastColumn()).getValues();
          }
        }
        
        dataChi.forEach(row => {
          if (monthFilter && monthFilter !== 'all') {
            let dateObj = row[1];
            let match = false;
            if (typeof dateObj === 'string' && dateObj.startsWith(monthFilter)) match = true;
            else if (dateObj instanceof Date) {
              let mYear = dateObj.getFullYear();
              let mMonth = ('0' + (dateObj.getMonth() + 1)).slice(-2);
              if (`${mYear}-${mMonth}` === monthFilter) match = true;
            }
            if (!match) return;
          }
          let danhMuc = row[2]; // Cột C: Mục đích chi
          let soTienThuc = parseFloat(row[8]) || 0; // Cột I: Tiền Thực Chi
          
          tongChi += soTienThuc;
          if(phanBoChiPhi[danhMuc]) {
            phanBoChiPhi[danhMuc] += soTienThuc;
          } else {
            phanBoChiPhi[danhMuc] = soTienThuc;
          }
        });
      }
    }

    const settingsSheet = ss.getSheetByName('Thiet_Lap');
    let mucTieu = 0;
    if (settingsSheet) {
      const settingsData = settingsSheet.getDataRange().getValues();
      settingsData.forEach(row => {
        if(row[0] === 'MucTieuLoiNhuan') mucTieu = parseFloat(row[1]) || 0;
      });
    }

    let loiNhuan = tongThu - tongChi;
    let bienLoiNhuan = tongThu > 0 ? Math.round((loiNhuan / tongThu) * 100) : 0;

    return {
      success: true,
      data: {
        tongThu: tongThu,
        tongChi: tongChi,
        loiNhuan: loiNhuan,
        mucTieu: mucTieu,
        bienLoiNhuan: bienLoiNhuan,
        phanBoChiPhi: phanBoChiPhi
      }
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
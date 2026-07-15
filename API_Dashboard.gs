function getDashboardData(monthFilter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transSheet = ss.getSheetByName('GiaoDich');
    if (!transSheet) throw new Error("Chưa có dữ liệu giao dịch.");

    const data = transSheet.getDataRange().getValues();
    data.shift();

    let tongThu = 0;
    let tongChi = 0;
    let phanBoChiPhi = {};

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

      let loai = row[2];
      let danhMuc = row[3];
      let soTienThuc = parseFloat(row[6]) || 0;

      if (loai === 'Thu') {
        tongThu += soTienThuc;
      } else if (loai === 'Chi') {
        tongChi += soTienThuc;
        if(phanBoChiPhi[danhMuc]) {
          phanBoChiPhi[danhMuc] += soTienThuc;
        } else {
          phanBoChiPhi[danhMuc] = soTienThuc;
        }
      }
    });

    const settingsSheet = ss.getSheetByName('ThietLap');
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
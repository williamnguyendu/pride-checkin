const SPREADSHEET_ID = '1xlGdKp9WZY2J6WcUvkfYrx1GTJHvrdrLvzeFmQPY86o';

function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getEvents') {
    return response(getEvents());
  }
  if (action === 'checkIn') {
    return response(processCheckIn(e.parameter.eventId, e.parameter.qrCode));
  }
  
  return response({ status: 'error', message: 'Action không hợp lệ.' });
}

function response(res) {
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

function getEvents() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("ThongTinSuKien");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const colMa = headers.indexOf("MaDangKySuKien");
    const colTen = headers.indexOf("TenSuKien");
    const colNgay = headers.indexOf("NgayToChuc");
    const colGio = headers.indexOf("GioToChuc");
    const colDiaDiem = headers.indexOf("DiaDiemToChuc");
    const colTrangThai = headers.indexOf("TrangThaiCheckIn");
    
    const events = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][colTrangThai] === "Cho phép") {
        events.push({
          id: String(data[i][colMa]),
          name: data[i][colTen],
          date: formatDate(data[i][colNgay]),
          time: data[i][colGio] ? formatTime(data[i][colGio]) : "Cả ngày",
          location: data[i][colDiaDiem]
        });
      }
    }
    return { status: 'success', data: events };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function processCheckIn(eventId, qrCode) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("DSDKSuKien");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const colHoTen = headers.indexOf("HoTen");
    const colSDT = headers.indexOf("SoDienThoai");
    const colMaBiMat = headers.indexOf("MaBiMat");
    const colCheckIn = headers.indexOf("CheckIn");
    const colMaSuKien = headers.indexOf("MaSuKienWorkshop");
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colMaSuKien]) === String(eventId) && String(data[i][colMaBiMat]) === String(qrCode)) {
        
        // Kiểm tra nếu đã check-in
        if (data[i][colCheckIn] !== "") {
          return { status: 'error', message: 'Vé này đã được check-in trước đó.' };
        }
        
        // Cập nhật thời gian check-in
        const now = new Date();
        const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        sheet.getRange(i + 1, colCheckIn + 1).setValue(timestamp);
        
        return {
          status: 'success',
          data: {
            hoTen: data[i][colHoTen],
            sdt: data[i][colSDT]
          }
        };
      }
    }
    return { status: 'error', message: 'Vé không hợp lệ, vui lòng kiểm tra lại hoặc liên hệ Ban tổ chức.' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function formatDate(date) {
  if (!date) return "";
  try {
    const d = new Date(date);
    if (isNaN(d.getTime())) return date;
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch(e) { return date; }
}

function formatTime(time) {
  if (!time) return "";
  try {
    const d = new Date(time);
    if (isNaN(d.getTime())) return time;
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  } catch(e) { return time; }
}
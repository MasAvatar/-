/**
 * Google Apps Script for Klong Si Crane Booking System (เวอร์ชันแก้ไข Date Object Bug ใน Apps Script)
 */

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var bookings = [];
  var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
  // เริ่มจากแถวที่ 2 (ดัชนี 1) ข้ามหัวตาราง
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) { // ตรวจสอบว่าวันที่ไม่ว่าง
      bookings.push({
        date: normalizeDate(row[0], timezone),
        crane: String(row[1] || '').trim(),
        address: String(row[2] || ''),
        phone: String(row[3] || '')
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(bookings))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var response = {};
  var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
  try {
    var jsonPayload = e.postData.contents;
    var data = JSON.parse(jsonPayload);
    
    if (data.action === 'delete') {
      var dateToDelete = data.date;
      var craneToDelete = data.crane;
      
      var values = sheet.getDataRange().getValues();
      var deleted = false;
      var debugInfo = [];
      
      // ค้นหาแถวเพื่อลบจากล่างขึ้นบน
      for (var i = values.length - 1; i >= 1; i--) {
        var row = values[i];
        if (!row[0]) continue;
        
        var rowDate = normalizeDate(row[0], timezone);
        var rowCrane = String(row[1] || '').trim();
        
        var isDateMatch = (rowDate === dateToDelete);
        var isCraneMatch = (rowCrane === craneToDelete);
        
        debugInfo.push({
          rowIndex: i + 1,
          rawDateValue: String(row[0]),
          normalizedDate: rowDate,
          rowCrane: rowCrane,
          dateMatch: isDateMatch,
          craneMatch: isCraneMatch
        });
        
        if (isDateMatch && isCraneMatch) {
          sheet.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }
      
      if (deleted) {
        response = { success: true, message: "ลบคิวงานวันที่ " + dateToDelete + " ขนาด " + craneToDelete + " สำเร็จ" };
      } else {
        response = { 
          success: false, 
          message: "ไม่พบข้อมูลคิวงานที่ตรงกันในระบบ",
          sentDate: dateToDelete,
          sentCrane: craneToDelete,
          debug: debugInfo
        };
      }
    } else {
      // เพิ่มข้อมูลคิวงาน
      var dateVal = data.date;
      var craneVal = data.crane;
      var addressVal = data.address || '';
      var phoneVal = data.phone || '';
      
      var values = sheet.getDataRange().getValues();
      var conflict = false;
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var rowDate = normalizeDate(row[0], timezone);
        
        if (rowDate === dateVal && String(row[1]).trim() === craneVal) {
          conflict = true;
          break;
        }
      }
      
      if (conflict) {
        response = { success: false, message: "ขออภัย! วันที่และขนาดรถเครนนี้มีการจองไว้ก่อนแล้วในระบบ" };
      } else {
        sheet.appendRow([dateVal, craneVal, addressVal, phoneVal]);
        response = { success: true, message: "บันทึกข้อมูลการจองรถเครนในวันที่ " + dateVal + " เรียบร้อยแล้ว!" };
      }
    }
  } catch (error) {
    response = { success: false, message: "เกิดข้อผิดพลาดบนระบบเซิร์ฟเวอร์: " + error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * แปลงค่าวันที่ในช่องเซลล์ของ Google Sheet ให้กลายเป็นฟอร์แมตมาตรฐาน yyyy-MM-DD เสมอ
 */
function normalizeDate(dateVal, timezone) {
  if (!dateVal) return "";
  
  // 1. ตรวจสอบว่าเป็นวัตถุ Date หรือไม่ (ใช้การแปลงชื่อคลาสเพื่อป้องกันการบั๊กของ Apps Script Engine)
  if (Object.prototype.toString.call(dateVal) === '[object Date]') {
    return Utilities.formatDate(dateVal, timezone, "yyyy-MM-dd");
  }
  
  var dateStr = String(dateVal).trim();
  
  // 2. ถ้าเป็น String แบบ ISO หรือ YYYY-MM-DD
  var match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return match[1] + "-" + match[2] + "-" + match[3];
  }
  
  // 3. ถ้าเป็น String รูปแบบอื่น ลองตัดเอาข้อความภาษาไทยในวงเล็บของ Timezone ออกก่อนแปลงเป็น Date
  // เช่น "Thu Jun 04 2026 00:00:00 GMT+0700 (เวลาอินโดจีน)" -> "Thu Jun 04 2026 00:00:00 GMT+0700"
  var cleanDateStr = dateStr.replace(/\s*\([^)]*\)/g, "");
  var parsed = new Date(cleanDateStr);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, timezone, "yyyy-MM-dd");
  }
  
  return dateStr;
}

/**
 * Occ-Health Data Hub - Backend Script (Full Version)
 * ฉบับสมบูรณ์: รวมระบบ KPI แยก Sheet และผู้รับผิดชอบ 3 คน
 */

var ss = SpreadsheetApp.getActiveSpreadsheet();
// ==========================================
// [12] ระบบรักษาความปลอดภัย (Security)
// ==========================================

/*var APP_PASSWORD = "10827"; // 🔑 [ตั้งรหัสผ่านตรงนี้ครับ]*/
var ADMIN_PASSWORD = "9999"; // 🔐 [เพิ่ม] รหัสสำหรับแก้ไขข้อมูล (Admin เท่านั้น)

function checkLoginPass(input) {
  return input.toString() == APP_PASSWORD.toString();
}

function checkAdminPass(input) {
  return input.toString() == ADMIN_PASSWORD.toString();
}
// ==========================================
// [1] ตั้งค่า Folder และ Calendar
// ==========================================

// 📂 ID โฟลเดอร์เก็บไฟล์ (แยกตามกลุ่มงาน)
var FOLDER_IDS = {
  "งานคลินิก": "15zzMm4HQCYXRVPRfIoHIIHwXEf1yuJ_s", 
  "งานมลพิษ": "1H6tuPM-_mvWZqE6OY5TCwL6J4BlWxlMY",
  "งานอาชีวฯป้องกัน": "12FJwmiXPBU3XVWBAffWFtlLphD7eNDJX",
  "งาน Check Up": "1HeCW_vJRx44my2iInx5zzvF0cKi7XGmo",
  "งานอาชีวฯในโรงพยาบาล": "12opS7Azs7ahwbMhZV39LpCy5RRfyXUiX",
  "ศูนย์เชี่ยวชาญฯ": "197W_P0Oyz79clmEiqRYKIYgdu_8yQgBs",
  "งานโครงการ": "" // ใส่ ID โฟลเดอร์สำหรับงานโครงการ/KPI เพิ่มได้ที่นี่
};

// 🗓️ ID ปฏิทิน Google Calendar
var CALENDAR_IDS = {
  "งานศูนย์เชี่ยวชาญฯ": "occ.hrh@gmail.com",
  "งาน Check up": "faceb90ae4f71e253e66122dcf532b254c1f4163dbc630cc5b8c75801b77f0ab@group.calendar.google.com",
  "งานคลินิก": "9f90b848303156d77b3aac262d07b3e33c8dc86bb8da6313a809e7fe9efe7ff4@group.calendar.google.com",
  "งานปรับเปลี่ยน": "5a012a720d26bef7cea3911d980feb44442213df4ca4d2a91455016ce45fe89f@group.calendar.google.com",
  "งานมลพิษ": "4913e0e1b441d120a4ce37ff142678fc74e562c718658a4fdc2556bdeb6ffeb9@group.calendar.google.com",
  "งานอาชีวฯ ใน รพ.": "5cd3f5b4c4e22c6c3ea051682de7ed506daec303017a3bcdeba5689e6f6d12ce@group.calendar.google.com",
  "ตารางเวรสอบสวน": "d241b33f34e9cbde41026fa8e4528fb8c04549b2f71d84409539dcdc143258bd@group.calendar.google.com",
  "อบรม/ดูงาน/งานอื่นๆ": "800978574c6b4b18b5228f253185eec809a787bc39e1ca37aa114dd2fcd56f4c@group.calendar.google.com"
};

// ฟังก์ชันเปิดเว็บ
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('Occ-Health Data Hub');
}

// ==========================================
// [2] ระบบจัดการปฏิทิน (Calendar Functions)
// ==========================================

function addEventToCalendar(data) {
  var targetCalId = CALENDAR_IDS[data.calName];
  if (!targetCalId) return "Error: ไม่พบ ID ปฏิทิน";

  try {
    var cal = CalendarApp.getCalendarById(targetCalId);
    if (!cal) return "Error: เข้าถึงปฏิทินไม่ได้";
    var date = new Date(data.date);
    cal.createAllDayEvent(data.title, date, {description: data.desc});
    return "Success";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function getTodayShifts() {
  var calId = CALENDAR_IDS["ตารางเวรสอบสวน"];
  try {
    var cal = CalendarApp.getCalendarById(calId);
    if (!cal) return "ไม่พบปฏิทิน";
    var today = new Date();
    var events = cal.getEventsForDay(today); 
    if (events.length === 0) return "วันนี้ไม่มีเวร";
    var details = events.map(function(e) { return e.getDescription(); })
      .filter(function(desc) { return desc !== ""; })
      .join(" / ");
    return details || "มีเวร (แต่ไม่ระบุชื่อ)";
  } catch (e) { return "Error"; }
}

// ==========================================
// [3] ระบบจัดการข้อมูล (Data Handling)
// ==========================================

function getAllData() {
  var taskData = getRawData('Tasks');
  var projectData = getRawData('Projects'); 
  var kpiData = getRawData('KPI'); // ดึงข้อมูลจาก Sheet KPI
  var contactData = getRawData('Contacts');
  
  return JSON.stringify({
    tasks: taskData.filter(function(t) { return t.status !== 'Archived'; }),
    projects: projectData.filter(function(p) { return p.status !== 'Archived'; }),
    kpis: kpiData.filter(function(k) { return k.status !== 'Archived'; }),
    contacts: contactData
  });
}

function getRawData(sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data.shift();
  return data.map(function(row) {
    var obj = {};
    headers.forEach(function(header, i) { obj[header] = row[i]; });
    return obj;
  });
}

// ==========================================
// [4] ฟังก์ชันบันทึกข้อมูล (Add/Edit/Update)
// ==========================================

function saveItemToSheet(data) {
  // เลือก Sheet ตามโหมดที่ส่งมา
  var sheetName = 'Tasks';
  if (data.mode === 'project') sheetName = 'Projects';
  if (data.mode === 'kpi') sheetName = 'KPI'; 
  
  var sheet = ss.getSheetByName(sheetName);
  var timestamp = new Date(); // เก็บเวลาปัจจุบัน
  
  // --- กรณีเพิ่มงานใหม่ (Add) ---
  if (data.action == 'add') {
    var newId = new Date().getTime().toString();
    var initialProgress = '0';
    if (data.type == 'checklist') {
      var items = data.target.split(',');
      var jsonArr = items.map(function(item) { return { item: item.trim(), status: false, file: "" }; });
      initialProgress = JSON.stringify(jsonArr);
    }
    // บันทึกลง Sheet (คอลัมน์ที่ 8 คือ task_class เก็บชื่อผู้รับผิดชอบ 3 คน)
    sheet.appendRow([newId, data.title, data.deadline, data.type, data.target, initialProgress, data.category, data.task_class, data.kpi_source, 'Active', '', data.is_daily, '', timestamp ]);
    return "Success";
  } 

  // --- กรณีแก้ไข หรืออัปเดต (Edit/Update) ---
  var rangeData = sheet.getDataRange().getValues();
  for (var i = 1; i < rangeData.length; i++) {
    if (rangeData[i][0].toString() == data.id.toString()) {
      var row = i + 1;
      
      // อัปเดต Timestamp เสมอเมื่อมีการแก้ไข
      if(sheet.getLastColumn() >= 14) sheet.getRange(row, 14).setValue(timestamp);

      if (data.action == 'แก้ไข') {
        sheet.getRange(row, 2).setValue(data.title);
        sheet.getRange(row, 3).setValue(data.deadline);
        sheet.getRange(row, 4).setValue(data.type);
        sheet.getRange(row, 5).setValue(data.target);
        sheet.getRange(row, 7).setValue(data.category);
        sheet.getRange(row, 8).setValue(data.task_class); // อัปเดตรายชื่อ 3 คน
        sheet.getRange(row, 9).setValue(data.kpi_source);
        sheet.getRange(row, 12).setValue(data.is_daily);
        
        // เช็คว่าเปลี่ยนประเภทการวัดผลหรือไม่
        var oldType = rangeData[i][3]; var oldTarget = rangeData[i][4];
        if (oldType != data.type || (data.type == 'checklist' && oldTarget != data.target)) {
           var newProg = '0';
           if (data.type == 'checklist') {
             var items = data.target.split(',');
             var jsonArr = items.map(function(item) { return { item: item.trim(), status: false, file: "" }; });
             newProg = JSON.stringify(jsonArr);
           }
           sheet.getRange(row, 6).setValue(newProg); 
           sheet.getRange(row, 13).setValue('');
        }
        return "Edited";
      }
      
      if (data.action == 'update_progress') {
        // บันทึก Log การเปลี่ยนแปลงตัวเลข
        if(data.progress !== undefined && !String(data.progress).includes('[')) {
           var oldVal = parseInt(rangeData[i][5] || 0);
           var newVal = parseInt(data.progress);
           if (!isNaN(oldVal) && !isNaN(newVal)) { 
             var diff = newVal - oldVal;
             var logStr = diff > 0 ? "+" + diff : diff.toString(); 
             if(diff !== 0) sheet.getRange(row, 13).setValue(logStr);

             // ✅✅✅ [ขั้นตอนที่ 3] แทรกตรงนี้ครับ ✅✅✅
             // ถ้ามีการบวกเพิ่ม (diff > 0) ให้บันทึกลง Sheet: Work_Log ด้วย
             if (diff > 0) {
                recordTransaction(data.id, rangeData[i][1], rangeData[i][6], diff);
             }
             // ✅✅✅ จบส่วนที่แทรก ✅✅✅

           }
        }
        
        if(data.progress !== undefined) sheet.getRange(row, 6).setValue(data.progress);
        if(data.status) sheet.getRange(row, 10).setValue(data.status);
        
        // กรณีแนบไฟล์หลัก (Evidence)
        if (data.fileData && data.fileName && !data.isChecklistItem) {
          var category = rangeData[i][6];
          var url = uploadToDrive(data.fileData, data.fileName, category);
          sheet.getRange(row, 11).setValue(url); 
          return "FileUploaded";
        }
        
        // กรณี Reset งาน
        if(data.reset) { 
          sheet.getRange(row, 6).setValue(data.new_progress); 
          sheet.getRange(row, 10).setValue('Active');
          sheet.getRange(row, 13).setValue(''); 
        }
        return "Updated";
      }
      
      if (data.action == 'upload_checklist_item') {
        var category = rangeData[i][6]; 
        var url = uploadToDrive(data.fileData, data.fileName, category);
        var checklistArr = []; try { checklistArr = JSON.parse(rangeData[i][5].toString()); } catch(e) {}
        if (checklistArr[data.itemIndex]) { checklistArr[data.itemIndex].file = url; }
        sheet.getRange(row, 6).setValue(JSON.stringify(checklistArr)); 
        return "ItemFileUploaded";
      }

      if (data.action == 'delete_file') {
        if (data.itemIndex != -1) {
           // ลบไฟล์ใน Checklist
           var checklistArr = []; try { checklistArr = JSON.parse(rangeData[i][5].toString()); } catch(e) {}
           if (checklistArr[data.itemIndex]) { 
             deleteFileFromDrive(checklistArr[data.itemIndex].file);
             checklistArr[data.itemIndex].file = "";
           }
           sheet.getRange(row, 6).setValue(JSON.stringify(checklistArr));
        } else {
           // ลบไฟล์หลัก
           deleteFileFromDrive(rangeData[i][10]);
           sheet.getRange(row, 11).setValue("");
        }
        return "FileDeleted";
      }
    }
  }
}

// ==========================================
// [5] ระบบจัดการไฟล์ (Drive)
// ==========================================

function uploadToDrive(base64Data, fileName, category) {
  try {
    var folder;
    if (FOLDER_IDS[category] && FOLDER_IDS[category] !== "") { 
      try { folder = DriveApp.getFolderById(FOLDER_IDS[category]); } catch(e) { folder = getCentralFolder(); } 
    } else { folder = getCentralFolder(); }
    
    var contentType = base64Data.substring(5, base64Data.indexOf(';')); 
    var bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    var blob = Utilities.newBlob(bytes, contentType, fileName); 
    var file = folder.createFile(blob); 
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
    return file.getUrl();
  } catch (e) { return "Error: " + e.toString(); }
}

function getCentralFolder() { 
  var folderName = "Task_Evidence"; 
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName); 
}

function deleteFileFromDrive(fileUrl) {
  if (!fileUrl || fileUrl == "") return;
  try {
    var id = fileUrl.match(/[-\w]{25,}/);
    if (id) DriveApp.getFileById(id[0]).setTrashed(true);
  } catch (e) { Logger.log("Error deleting file: " + e.toString()); }
}

// ==========================================
// [6] ระบบจัดการสมุดโทรศัพท์ (Contacts)
// ==========================================

// อย่าลืมประกาศ ADMIN_PASSWORD ด้านบนสุดนะครับ
// var ADMIN_PASSWORD = "9999"; 

function saveContactToSheet(data) {
  var sheet = ss.getSheetByName('Contacts');
  
  // --- กรณีเพิ่มรายชื่อ (เหมือนเดิม) ---
  if (data.action == 'add_contact') { 
    if (data.authPass != ADMIN_PASSWORD) return "WrongPass";
    
    var newId = new Date().getTime().toString();
    // ตั้งค่าเริ่มต้น: ถ้าไม่ระบุอะไร ให้ใช้เลข 1234
    var defaultPin = "1234"; 
    
    // เพิ่ม Column ที่ 4 เก็บ PIN
    sheet.appendRow([newId, data.name, data.phone, defaultPin]); 
    return "Success";
  } 
  
  // --- กรณีแก้ไข (ปรับปรุงใหม่) ---
  else if (data.action == 'edit_contact') { 
    var rangeData = sheet.getDataRange().getValues();
    for (var i = 1; i < rangeData.length; i++) { 
      if (rangeData[i][0].toString() == data.id.toString()) { 
        
        // Col D คือ index 3 (เริ่มนับจาก 0: A=0, B=1, C=2, D=3)
        var storedUserPin = rangeData[i][3] ? rangeData[i][3].toString() : "1234"; 
        
        // เช็คว่ารหัสที่กรอกมา (authPass) ตรงกับ Admin หรือ รหัสเดิม หรือไม่
        if (data.authPass != ADMIN_PASSWORD && data.authPass != storedUserPin) {
          return "WrongPass";
        }

        // 1. อัปเดตชื่อ-เบอร์
        sheet.getRange(i+1, 2).setValue(data.name);
        sheet.getRange(i+1, 3).setValue(data.phone);
        
        // 2. 🔥 ถ้ามีการกรอกรหัสใหม่มา (newUserPin ไม่ว่าง) ให้อัปเดต PIN ด้วย
        if (data.newUserPin && data.newUserPin.toString().trim() !== "") {
           sheet.getRange(i+1, 4).setValue(data.newUserPin.toString().trim());
        }
        
        return "Updated"; 
      } 
    } 
  }
  
  // --- กรณีลบ (เหมือนเดิม) ---
  else if (data.action == 'delete_contact') {
     // (ใช้ code เดิมจากคำตอบก่อนหน้าได้เลย)
     if (data.authPass != ADMIN_PASSWORD) return "WrongPass";
     var rangeData = sheet.getDataRange().getValues();
    for (var i = 1; i < rangeData.length; i++) { 
      if (rangeData[i][0].toString() == data.id.toString()) { 
        sheet.deleteRow(i+1);
        return "Deleted";
      } 
    }
  }
}
// ==========================================
// [7] ระบบจัดการเลขรันเอกสาร (DocRunning)
// ==========================================

function getDocRunningNumber(type) {
  var sheet = getDocRunningSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == type) return data[i][1];
  }
  return 0;
}

function incrementDocRunningNumber(type) {
  var sheet = getDocRunningSheet();
  var data = sheet.getDataRange().getValues();
  var found = false;
  var newNum = 1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == type) {
      var currentNum = parseInt(data[i][1]);
      newNum = currentNum + 1;
      sheet.getRange(i + 1, 2).setValue(newNum);
      found = true;
      break;
    }
  }
  if (!found) { sheet.appendRow([type, 1]); newNum = 1; }
  return newNum;
}

function setDocRunningNumber(type, newNum) {
  var sheet = getDocRunningSheet();
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == type) {
      sheet.getRange(i + 1, 2).setValue(parseInt(newNum));
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([type, parseInt(newNum)]);
  return "Saved";
}

function getDocRunningSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DocRunning");
  if (!sheet) {
    sheet = ss.insertSheet("DocRunning");
    sheet.appendRow(["DocType", "LastNumber"]); 
  }
  return sheet;
}
// ==========================================
// [8] ระบบปิดยอดรายเดือน (Monthly Snapshot)
// ==========================================

function saveMonthlySnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Tasks");
  var targetSheet = ss.getSheetByName("History_Log");
  
  // 1. ถ้ายังไม่มี Sheet History ให้สร้างใหม่
  if (!targetSheet) {
    targetSheet = ss.insertSheet("History_Log");
    targetSheet.appendRow(["Month_Year", "Task_ID", "Task_Title", "Category", "Total_Count", "Timestamp"]);
  }
  
  var data = sourceSheet.getDataRange().getValues();
  var timestamp = new Date();
  var monthYear = Utilities.formatDate(timestamp, "Asia/Bangkok", "yyyy-MM"); // เช่น 2025-01
  var savedCount = 0;

  // 2. วนลูปเก็บข้อมูลเฉพาะงานแบบ "นับจำนวน" (type = number)
  for (var i = 1; i < data.length; i++) {
    var row = i + 1;
    var type = data[i][3]; // Column D
    var progress = parseInt(data[i][5] || 0); // Column F
    var status = data[i][9]; // Column J

    if (type == 'number' && status != 'Archived') {
      // บันทึกลง History
      targetSheet.appendRow([
        monthYear,
        data[i][0], // ID
        data[i][1], // Title
        data[i][6], // Category
        progress,   // ยอดที่ทำได้
        timestamp
      ]);
      
      // 3. รีเซ็ตยอดเป็น 0 (ถ้าต้องการ)
      // ถ้าไม่อยากให้รีเซ็ตอัตโนมัติ ให้ลบบรรทัดนี้ออกครับ
      sourceSheet.getRange(row, 6).setValue(0);
      sourceSheet.getRange(row, 13).setValue("Reset " + monthYear); // Clear Log
      
      savedCount++;
    }
  }
  
  return "บันทึกเรียบร้อย " + savedCount + " รายการ ประจำเดือน " + monthYear;
}
// ==========================================
// [9] ระบบดึงข้อมูลรายงานย้อนหลัง
// ==========================================

function getHistoryMonths() {
  var sheet = ss.getSheetByName("History_Log");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var months = [];
  // วนลูปเก็บชื่อเดือน (ไม่เอาซ้ำ)
  for (var i = 1; i < data.length; i++) {
    var m = data[i][0]; // Column A: Month_Year
    if (m && months.indexOf(m) === -1) months.push(m);
  }
  return months.sort().reverse(); // เรียงเดือนล่าสุดขึ้นก่อน
}

function getHistoryReport(month) {
  var sheet = ss.getSheetByName("History_Log");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    // เช็คเดือน และต้องไม่ใช่แถวที่ถูก Archive
    if (data[i][0] == month) {
      result.push({
        title: data[i][2],     // Task_Title
        category: data[i][3],  // Category
        progress: data[i][4]   // Total_Count
      });
    }
  }
  return result;
}
// ==========================================
// [10] ระบบบันทึก Transaction รายวัน (เพื่อให้ดูย้อนหลังได้แม่นยำ)
// ==========================================

// 1. ฟังก์ชันช่วยบันทึก Log (จะถูกเรียกตอนกด Save/Update)
function recordTransaction(taskId, title, category, addedAmount) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Work_Log");
  
  if (!logSheet) { // กันเหนียว ถ้าลืมสร้าง Sheet
    logSheet = ss.insertSheet("Work_Log");
    logSheet.appendRow(["Timestamp", "Task_ID", "Task_Title", "Category", "Amount_Added"]);
  }
  
  // บันทึกเวลาปัจจุบัน และ ยอดที่บวกเพิ่ม
  logSheet.appendRow([new Date(), taskId, title, category, addedAmount]);
}
// 2. ฟังก์ชันดึงรายงานตามช่วงเวลา (สำหรับหน้า Report) - แก้ไขเรื่อง Checklist
function getReportByDateRange(startDateStr, endDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // A. เตรียมข้อมูล Task ปัจจุบัน
  var taskSheet = ss.getSheetByName("Tasks");
  var taskData = taskSheet.getDataRange().getValues();
  var taskMap = {}; 
  
  for (var i = 1; i < taskData.length; i++) {
    var tid = taskData[i][0];
    taskMap[tid] = {
      title: taskData[i][1],
      category: taskData[i][6],
      target: taskData[i][4],
      progress: taskData[i][5], // เก็บ progress ปัจจุบัน (JSON หรือ ตัวเลข)
      type: taskData[i][3],     // เก็บประเภทงาน (checklist / number)
      range_total: 0 
    };
  }

  // B. ดึงข้อมูลจาก Log มาคำนวณ (เฉพาะงานตัวเลข)
  var logSheet = ss.getSheetByName("Work_Log");
  if (logSheet) {
    var logData = logSheet.getDataRange().getValues();
    var start = new Date(startDateStr); start.setHours(0,0,0,0);
    var end = new Date(endDateStr); end.setHours(23,59,59,999);
    
    for (var j = 1; j < logData.length; j++) {
      var logDate = new Date(logData[j][0]);
      var logId = logData[j][1];
      var amount = parseInt(logData[j][4] || 0);
      
      if (logDate >= start && logDate <= end) {
        if (taskMap[logId]) {
          taskMap[logId].range_total += amount;
        }
      }
    }
  }
  
  // C. แปลงกลับเป็น Array
  var reportList = [];
  for (var key in taskMap) {
    var t = taskMap[key];
    reportList.push({
      id: key,
      title: t.title,
      category: t.category,
      target: t.target,
      progress: t.progress,      // ส่ง progress ดิบๆ ไปให้หน้าเว็บแกะเอง
      range_total: t.range_total,
      type: t.type               // ส่ง Type จริงๆ ไป
    });
  }
  
  return reportList;
}
// ==========================================
// [11] ระบบดึงยอดรายวัน (Today's Counter)
// ==========================================
function getTodayLogStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Work_Log");
  if (!logSheet) return {}; // ถ้าไม่มี Sheet Log ให้ส่งค่าว่างกลับไป
  
  var data = logSheet.getDataRange().getValues();
  var todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
  var stats = {};

  // วนลูปดู Log ทั้งหมด
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0]; // Column A: Timestamp
    // แปลงวันที่ใน Log เป็นรูปแบบ yyyy-MM-dd เพื่อเทียบกับวันนี้
    var logDateStr = Utilities.formatDate(new Date(rowDate), "Asia/Bangkok", "yyyy-MM-dd");
    
    if (logDateStr === todayStr) {
      var taskId = data[i][1];
      var amount = parseInt(data[i][4] || 0);
      
      if (!stats[taskId]) stats[taskId] = 0;
      stats[taskId] += amount;
    }
  }
  return stats; // ส่งกลับเป็นก้อน เช่น { "ID_123": 5, "ID_456": 2 }
}
// ==========================================
// [12] ระบบแจ้งเตือนผ่าน Telegram (Auto Alert)
// ==========================================

// 🔑 ใส่ข้อมูลจาก Telegram ตรงนี้ครับ
var TELEGRAM_TOKEN = "8349554549:AAE9reU225Nod4z_ONWZ_Ea6wQFaifbxOb4"; 
var TELEGRAM_CHAT_ID = "-1002490816700";
var WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz-oSRMZxiQnxEdF3T1AihAYNqjYSkoCayebnooeQ2fVj0c2G3Jj67uFq40LE544BPFMg/exec"; 

function autoSendDailyReport() {
  var today = new Date();
  var day = today.getDay(); // 0=อาทิตย์, 6=เสาร์
  
  if (day === 0 || day === 6) {
    console.log("วันนี้วันหยุด ไม่ส่งรายงานครับ");
    return;
  }

  // 1. คำนวณยอดที่เพิ่มขึ้นตั้งแต่ 15.30 ของเมื่อวาน
  var changesMap = getChangesSinceYesterday();
  
  var msgBody = "";
  
  // 2. ดึงงานประจำ (ส่ง changesMap เข้าไปเทียบ)
  msgBody += getRoutineReportWithGroups(changesMap);
  
  // 3. ดึง Project และ KPI (ส่ง changesMap เข้าไปเทียบ)
  msgBody += getSectionReport("Projects", "🚀 โครงการ (Projects)", changesMap);
  msgBody += getSectionReport("KPI", "📈 ตัวชี้วัด (KPIs)", changesMap);

  if (msgBody === "") {
    msgBody = "▫️ (ไม่มีงาน Active ในระบบ)\n";
  }

  var dateStr = Utilities.formatDate(today, "Asia/Bangkok", "dd/MM/yyyy");
  
  var message = "📊 *รายงานสถานะภาพรวม* (" + dateStr + ")\n" +
                "กลุ่มงานอาชีวเวชกรรมฯ\n" +
                "========================\n" +
                msgBody +
                "========================\n" +
                "ℹ️ *ตัดยอด 15.30 น. ของเมื่อวาน - ปัจจุบัน*\n" +
                "🔗 [เข้าสู่ระบบ Occ-Health Data Hub](" + WEB_APP_URL + ")";

  sendTelegramMsg(message);
}

// --- ฟังก์ชันคำนวณยอดที่เพิ่มขึ้น (New) ---
function getChangesSinceYesterday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Work_Log");
  if (!logSheet) return {};

  var data = logSheet.getDataRange().getValues();
  var changes = {};
  
  // ตั้งเวลาตัดยอด: เมื่อวาน เวลา 15:30:00
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 1); // ย้อนไป 1 วัน
  cutoff.setHours(15, 30, 0, 0);       // เวลา 15:30

  // วนลูปดู Log (เริ่มแถว 2)
  for (var i = 1; i < data.length; i++) {
    var timestamp = new Date(data[i][0]); // Col A: Timestamp
    var taskId = data[i][1];              // Col B: ID
    var amount = parseInt(data[i][4]) || 0; // Col E: Amount Added

    // ถ้าเวลาใน Log เกิดขึ้น "หลัง" เวลาตัดยอด ให้เอามาบวก
    if (timestamp > cutoff) {
      if (!changes[taskId]) changes[taskId] = 0;
      changes[taskId] += amount;
    }
  }
  return changes;
}

// --- ฟังก์ชันดึงงานประจำ แบบจัดกลุ่ม (Updated) ---
function getRoutineReportWithGroups(changesMap) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Tasks");
  if (!sheet) return "";
  
  var data = sheet.getDataRange().getValues();
  var groups = {};
  var hasData = false;

  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];       // ต้องใช้ ID เพื่อไปเทียบกับ Log
    var title = data[i][1];
    var type = data[i][3];
    var target = data[i][4];
    var progress = data[i][5];
    var category = data[i][6];
    var taskClass = data[i][7];
    var status = data[i][9];

    if (status === "Active" && taskClass === "งานประจำ") {
      var displayValue = "";
      
      if (type === 'number') {
        var num = parseInt(progress) || 0;
        var tar = parseInt(target) || 0;
        var targetStr = (tar > 0) ? " / " + tar.toLocaleString() : "";
        
        // เช็คว่าวันนี้มีบวกเพิ่มไหม
        var added = changesMap[id] || 0;
        var addedStr = (added > 0) ? " (+" + added + ")" : ""; // ถ้ามีให้โชว์ (+xx)
        
        displayValue = num.toLocaleString() + targetStr + addedStr; // รวมร่าง
        
      } else {
        // Checklist
        try {
          var items = JSON.parse(progress);
          var done = items.filter(function(x){return x.status}).length;
          displayValue = done + "/" + items.length + " ขั้นตอน"; 
        } catch(e) { displayValue = "N/A"; }
      }

      if (!groups[category]) groups[category] = [];
      groups[category].push("▫️ " + title + ": *" + displayValue + "*");
      hasData = true;
    }
  }

  if (!hasData) return "";
  
  var output = "*📋 งานประจำ (Routine)*\n";
  for (var catName in groups) {
    if (groups[catName].length > 0) {
      output += "📂 *" + catName + "*\n" + groups[catName].join("\n") + "\n\n";
    }
  }
  return output;
}

// --- ฟังก์ชันดึง Project/KPI (Updated) ---
function getSectionReport(sheetName, headerTitle, changesMap) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "";

  var data = sheet.getDataRange().getValues();
  var sectionContent = "";
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var id = data[i][0]; // ID
    var title = data[i][1];
    var type = data[i][3];
    var target = data[i][4];
    var progress = data[i][5];
    var status = data[i][9];

    if (status === "Active") {
      var displayValue = "";
      
      if (type === 'number') {
        var num = parseInt(progress) || 0;
        var tar = parseInt(target) || 0;
        var targetStr = (tar > 0) ? " / " + tar.toLocaleString() : "";
        
        // เช็คว่าวันนี้มีบวกเพิ่มไหม
        var added = changesMap[id] || 0;
        var addedStr = (added > 0) ? " (+" + added + ")" : "";
        
        displayValue = num.toLocaleString() + targetStr + addedStr;

      } else {
        try {
          var items = JSON.parse(progress);
          var done = items.filter(function(x){return x.status}).length;
          displayValue = done + "/" + items.length + " ขั้นตอน"; 
        } catch(e) { displayValue = "N/A"; }
      }
      sectionContent += "▫️ " + title + ":  *" + displayValue + "*\n";
      count++;
    }
  }

  if (count > 0) return "*" + headerTitle + "*\n" + sectionContent + "\n";
  return "";
}
// --- ส่ง Telegram (เหมือนเดิม) ---
function sendTelegramMsg(msg) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  var payload = { "chat_id": TELEGRAM_CHAT_ID, "text": msg, "parse_mode": "Markdown" };
  var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload) };
  try { UrlFetchApp.fetch(url, options); } catch(e) { console.log(e); }
}
// ==========================================
// [13] ระบบ War Room (Update V.6: Dynamic Tasks)
// ==========================================

// 1. ประกาศ/ยกเลิก (Reset Custom Tasks ด้วย)
function setEmergencyState(password, isActive, message) {
  if (password != ADMIN_PASSWORD) return "WrongPass";
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty('EMERGENCY_ACTIVE', isActive);
  props.setProperty('EMERGENCY_MSG', message || "เกิดเหตุฉุกเฉิน!");
  
  if (isActive) {
    // 1. Reset Main Checklist (30 ช่อง SOP)
    var defaultChecklist = [];
    for(var i=0; i<30; i++) defaultChecklist.push({status: false, file: null});
    props.setProperty('EMERGENCY_CHECKLIST', JSON.stringify(defaultChecklist));
    
    // 2. ✅ Reset Custom Tasks (ภารกิจเพิ่มเติม) -> เริ่มต้นเป็นค่าว่าง
    props.setProperty('EMERGENCY_CUSTOM_TASKS', "[]");

    // 3. Reset Log & Attendance
    var startLog = [{time: getTimeNow(), msg: "เริ่มเปิดศูนย์ War Room: " + message}];
    props.setProperty('EMERGENCY_LOGS', JSON.stringify(startLog));
    props.setProperty('EMERGENCY_ATTENDANCE', "[]");

    // 4. แจ้งเตือน Telegram (ปรับปรุงข้อความ)
    var alertMsg = "🚨 *EMERGENCY ALERT!* 🚨\n\n" + 
                   "⚠️ *เหตุการณ์:* " + message + "\n\n" +
                   "🔴 *รายงานตัว และปฏิบัติตามแผน*\n" +
                   "🔗 [👉 กดเพื่อเข้าสู่ War Room](" + WEB_APP_URL + ")";
                   
    try { sendTelegramMsg(alertMsg); } catch(e) {}
  } else {
    // ปิด
    var cancelMsg = "✅ *ยกเลิกภาวะฉุกเฉิน*";
    try { sendTelegramMsg(cancelMsg); } catch(e) {}
  }
  return "Success";
}

// 2. ดึงสถานะ (ส่ง Custom Tasks กลับไปด้วย)
function getEmergencyState() {
  var props = PropertiesService.getScriptProperties();
  return {
    isActive: props.getProperty('EMERGENCY_ACTIVE') === 'true',
    message: props.getProperty('EMERGENCY_MSG'),
    checklist: JSON.parse(props.getProperty('EMERGENCY_CHECKLIST') || "[]"),
    customTasks: JSON.parse(props.getProperty('EMERGENCY_CUSTOM_TASKS') || "[]"), // ✅ ส่งกลับ
    logs: JSON.parse(props.getProperty('EMERGENCY_LOGS') || "[]"),
    attendance: JSON.parse(props.getProperty('EMERGENCY_ATTENDANCE') || "[]")
  };
}

// 3. (Main SOP) อัปเดต Checklist หลัก (เหมือนเดิม)
function updateChecklist(password, index, isChecked) {
  var props = PropertiesService.getScriptProperties();
  var checklist = JSON.parse(props.getProperty('EMERGENCY_CHECKLIST') || "[]");
  if (!checklist[index] || typeof checklist[index] !== 'object') {
    checklist[index] = { status: isChecked, file: null };
  } else {
    checklist[index].status = isChecked;
  }
  props.setProperty('EMERGENCY_CHECKLIST', JSON.stringify(checklist));
  return checklist;
}

// --- ✨ ส่วนใหม่: จัดการภารกิจเพิ่มเติม (Custom Tasks) ---

// 4. เพิ่มภารกิจใหม่
function addCustomTask(taskName) {
  var props = PropertiesService.getScriptProperties();
  var tasks = JSON.parse(props.getProperty('EMERGENCY_CUSTOM_TASKS') || "[]");
  
  tasks.push({
    id: new Date().getTime(), // ใช้เวลาเป็น ID
    name: taskName,
    status: false,
    file: null
  });
  
  props.setProperty('EMERGENCY_CUSTOM_TASKS', JSON.stringify(tasks));
  return tasks;
}

// 5. อัปเดตสถานะภารกิจเพิ่มเติม (ติ๊กถูก)
function updateCustomTask(index, isChecked) {
  var props = PropertiesService.getScriptProperties();
  var tasks = JSON.parse(props.getProperty('EMERGENCY_CUSTOM_TASKS') || "[]");
  
  if (tasks[index]) {
    tasks[index].status = isChecked;
    props.setProperty('EMERGENCY_CUSTOM_TASKS', JSON.stringify(tasks));
  }
  return tasks;
}

// 6. ลบภารกิจเพิ่มเติม
function deleteCustomTask(index) {
  var props = PropertiesService.getScriptProperties();
  var tasks = JSON.parse(props.getProperty('EMERGENCY_CUSTOM_TASKS') || "[]");
  
  if (index >= 0 && index < tasks.length) {
    tasks.splice(index, 1); // ลบทิ้ง
    props.setProperty('EMERGENCY_CUSTOM_TASKS', JSON.stringify(tasks));
  }
  return tasks;
}

// 3. (แก้ไข) ฟังก์ชันกดรายงานตัว (ต้องมี PIN)
function submitEmergencyAttendance(name, inputPin) {
  // 1. ตรวจสอบรหัสผ่านก่อน
  if (!verifyUserPin(name, inputPin)) {
    return "WrongPIN";
  }

  var props = PropertiesService.getScriptProperties();
  var list = JSON.parse(props.getProperty('EMERGENCY_ATTENDANCE') || "[]");
  
  // เช็คว่าเคยรายงานตัวไปหรือยัง (กันซ้ำ)
  var existing = list.find(x => x.name == name);
  if (!existing) {
    list.unshift({
      name: name,
      role: "เจ้าหน้าที่", // (อนาคตอาจดึงตำแหน่งจริงมาใส่)
      time: getTimeNow()
    });
    props.setProperty('EMERGENCY_ATTENDANCE', JSON.stringify(list));
  }
  
  return list; // ส่งรายการล่าสุดกลับไป
}

// 4. (เพิ่มใหม่) ฟังก์ชันตรวจสอบรหัสผ่าน (PIN หรือ 4 ตัวท้ายเบอร์โทร)
function verifyUserPin(name, inputPin) {
  // ⚠️ ตั้งค่า: ชื่อ Sheet ที่เก็บรายชื่อ (แก้ให้ตรงกับของคุณป๊อป)
  var sheetName = "Contacts"; // หรือ "Phonebook" หรือ "รายชื่อ"
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return true; // ถ้าหา Sheet ไม่เจอ ให้ผ่านไปก่อน (กันระบบล่ม)

  var data = sheet.getDataRange().getValues();
  // สมมติ: Col A=ID, B=ชื่อ, C=เบอร์, D=PIN (ปรับตามจริง)
  // ให้ Loop หาชื่อ
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == name) { // เจอชื่อแล้ว (Col B)
      var phone = String(data[i][2]).replace(/-/g, "").trim(); // เบอร์โทร (Col C)
      var storedPin = String(data[i][3]).trim(); // PIN (Col D) สมมติว่าอยู่คอลัมน์ 4
      
      // 1. ถ้ามี PIN ให้เช็ค PIN
      if (storedPin !== "" && storedPin !== "undefined") {
        return storedPin == inputPin;
      } 
      // 2. ถ้าไม่มี PIN ให้เช็ค 4 ตัวท้ายเบอร์โทร
      else if (phone.length >= 4) {
        var last4 = phone.substr(phone.length - 4);
        return last4 == inputPin;
      }
      // 3. ถ้าไม่มีทั้งคู่ ให้ผ่านเลย (หรือจะให้ใส่ 0000 ก็ได้)
      return true;
    }
  }
  return false; // หาชื่อไม่เจอ
}
// 3. อัปเดต Checklist (ติ๊กถูก/ผิด)
function updateChecklist(password, index, isChecked) {
  var props = PropertiesService.getScriptProperties();
  var checklist = JSON.parse(props.getProperty('EMERGENCY_CHECKLIST') || "[]");
  
  // ตรวจสอบและแปลงโครงสร้างข้อมูลให้เป็น Object (กัน Error จากข้อมูลเก่า)
  if (!checklist[index] || typeof checklist[index] !== 'object') {
    checklist[index] = { status: isChecked, file: null };
  } else {
    checklist[index].status = isChecked; // อัปเดตเฉพาะสถานะ ไฟล์ยังคงเดิม
  }
  
  props.setProperty('EMERGENCY_CHECKLIST', JSON.stringify(checklist));
  return checklist;
}

// 4. (ใหม่) อัปโหลดหลักฐานลง Checklist
function uploadEmergencyEvidence(data) {
  // data = { index, fileData, fileName, mimeType }
  var props = PropertiesService.getScriptProperties();
  var checklist = JSON.parse(props.getProperty('EMERGENCY_CHECKLIST') || "[]");
  
  // 4.1 สร้างโฟลเดอร์เก็บไฟล์ (ถ้ายังไม่มี)
  var folderName = "WarRoom_Evidence";
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  
  // 4.2 สร้างไฟล์จาก Base64
  var blob = Utilities.newBlob(Utilities.base64Decode(data.fileData), data.mimeType, data.fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // เปิดแชร์ให้ดูได้
  
  // 4.3 บันทึกลง Checklist
  if (!checklist[data.index] || typeof checklist[data.index] !== 'object') {
    checklist[data.index] = { status: true, file: null };
  }
  
  // เก็บข้อมูลไฟล์
  checklist[data.index].file = {
    name: data.fileName,
    url: file.getDownloadUrl(),
    id: file.getId()
  };
  checklist[data.index].status = true; // อัปโหลดเสร็จถือว่าทำข้อนั้นแล้ว (Auto Check)
  
  props.setProperty('EMERGENCY_CHECKLIST', JSON.stringify(checklist));
  return checklist;
}

// 5. (ใหม่) ลบหลักฐานออกจาก Checklist
function deleteEmergencyEvidence(index) {
  var props = PropertiesService.getScriptProperties();
  var checklist = JSON.parse(props.getProperty('EMERGENCY_CHECKLIST') || "[]");
  
  if (checklist[index] && checklist[index].file) {
    checklist[index].file = null; // ลบ Link ออก (ไฟล์ใน Drive ยังอยู่กันพลาด)
    props.setProperty('EMERGENCY_CHECKLIST', JSON.stringify(checklist));
  }
  return checklist;
}

// --- ส่วนจัดการ Timeline Log (คงเดิม) ---

function addCommanderLog(msg) {
  var props = PropertiesService.getScriptProperties();
  var logs = JSON.parse(props.getProperty('EMERGENCY_LOGS') || "[]");
  logs.unshift({ time: getTimeNow(), msg: msg });
  if (logs.length > 50) logs.pop(); // เพิ่มลิมิตเป็น 50 ข้อความ
  props.setProperty('EMERGENCY_LOGS', JSON.stringify(logs));
  return logs;
}

function editCommanderLog(password, index, newMsg) {
  if (password != ADMIN_PASSWORD) return "WrongPass";
  var props = PropertiesService.getScriptProperties();
  var logs = JSON.parse(props.getProperty('EMERGENCY_LOGS') || "[]");
  if (index >= 0 && index < logs.length) {
    logs[index].msg = newMsg;
    props.setProperty('EMERGENCY_LOGS', JSON.stringify(logs));
  }
  return logs;
}

function deleteCommanderLog(password, index) {
  if (password != ADMIN_PASSWORD) return "WrongPass";
  var props = PropertiesService.getScriptProperties();
  var logs = JSON.parse(props.getProperty('EMERGENCY_LOGS') || "[]");
  if (index >= 0 && index < logs.length) {
    logs.splice(index, 1);
    props.setProperty('EMERGENCY_LOGS', JSON.stringify(logs));
  }
  return logs;
}

// Helper
function getTimeNow() {
  var d = new Date();
  return Utilities.formatDate(d, "Asia/Bangkok", "HH:mm");
}
function testTelegram() {
  var msg = "🔔 *ทดสอบระบบ:* บอท War Room ใช้งานได้แล้วครับ!";
  sendTelegramMsg(msg);
}
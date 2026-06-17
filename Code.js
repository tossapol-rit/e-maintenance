function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบแจ้งซ่อมออนไลน์ (E-Maintenance)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log("Error in doGet: " + error.toString());
    return HtmlService.createHtmlOutput("เกิดข้อผิดพลาดในการโหลดหน้าเว็บ: " + error.toString());
  }
}

// ฟังก์ชันดึง Spreadsheet และ Sheet ต่างๆ เพื่อให้เรียกใช้ง่ายและจัดการ Error ได้
function getSpreadsheet() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!sheetId) throw new Error("ไม่พบ SPREADSHEET_ID ใน Properties");
  return SpreadsheetApp.openById(sheetId);
}

// ฟังก์ชันสำหรับตรวจสอบสิทธิ์ผู้ใช้จาก LINE User ID
function checkUserRole(lineUserId) {
  try {
    var ss = getSpreadsheet();
    
    // ตรวจสอบในแท็บ Executives
    var execSheet = ss.getSheetByName('Executives');
    if (execSheet) {
      var execData = execSheet.getDataRange().getValues();
      for (var i = 1; i < execData.length; i++) {
        if (execData[i][3] === lineUserId) { // คอลัมน์ D (index 3) คือ LINE_User_ID
          return { role: 'Executive', info: { id: execData[i][0], name: execData[i][1] } };
        }
      }
    }
    
    // ตรวจสอบในแท็บ Technicians
    var techSheet = ss.getSheetByName('Technicians');
    if (techSheet) {
      var techData = techSheet.getDataRange().getValues();
      for (var i = 1; i < techData.length; i++) {
        if (techData[i][4] === lineUserId) { // คอลัมน์ E (index 4) คือ LINE_User_ID
          return { role: 'Technician', info: { id: techData[i][0], name: techData[i][1] } };
        }
      }
    }
    
    // ตรวจสอบในแท็บ Admins
    var adminSheet = ss.getSheetByName('Admins');
    if (adminSheet) {
      var adminData = adminSheet.getDataRange().getValues();
      for (var i = 1; i < adminData.length; i++) {
        if (adminData[i][3] === lineUserId) { // คอลัมน์ D (index 3) คือ LINE_User_ID
          return { role: 'Admin', info: { id: adminData[i][0], name: adminData[i][1] } };
        }
      }
    }
    
    return { role: 'Unknown', info: null };
  } catch (error) {
    Logger.log("Error in checkUserRole: " + error.toString());
    throw error;
  }
}

// ฟังก์ชันสร้าง Job ID อัตโนมัติ (JOB-YYYYMMDD-HHMMSS)
function generateJobId() {
  var now = new Date();
  var year = now.getFullYear();
  var month = ('0' + (now.getMonth() + 1)).slice(-2);
  var day = ('0' + now.getDate()).slice(-2);
  var hours = ('0' + now.getHours()).slice(-2);
  var minutes = ('0' + now.getMinutes()).slice(-2);
  var seconds = ('0' + now.getSeconds()).slice(-2);
  return 'JOB-' + year + month + day + '-' + hours + minutes + seconds;
}

// ฟังก์ชันดึงข้อมูลช่างทั้งหมดสำหรับให้ Admin เลือกมอบหมายงาน
function getAllTechnicians() {
  try {
    var ss = getSpreadsheet();
    var techSheet = ss.getSheetByName('Technicians');
    if (!techSheet) return [];
    var data = techSheet.getDataRange().getValues();
    var techs = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // ถ้ามี Tech_ID
        techs.push({ id: data[i][0], name: data[i][1] });
      }
    }
    return techs;
  } catch (error) {
    Logger.log("Error in getAllTechnicians: " + error.toString());
    return [];
  }
}

// ฟังก์ชันดึงงานทั้งหมด (สำหรับ Admin)
function getAllJobs() {
  try {
    var ss = getSpreadsheet();
    var jobSheet = ss.getSheetByName('Jobs');
    if (!jobSheet) return [];
    var data = jobSheet.getDataRange().getDisplayValues(); // ใช้ getDisplayValues เพื่อให้ได้ format วันที่ที่อ่านง่าย
    var jobs = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        jobs.push({
          jobId: data[i][0],
          customerName: data[i][1],
          phone: data[i][2],
          desc: data[i][3],
          status: data[i][4],
          assignedTech: data[i][5],
          reportedAt: data[i][6],
          approvalStatus: data[i][12],
          sparePartCost: data[i][13]
        });
      }
    }
    // เรียงงานล่าสุดขึ้นก่อน
    return jobs.reverse();
  } catch (error) {
    Logger.log("Error in getAllJobs: " + error.toString());
    return [];
  }
}

// ฟังก์ชันดึงงานที่รออนุมัติ (สำหรับ Executive)
function getPendingApprovalJobs() {
  try {
    var jobs = getAllJobs();
    return jobs.filter(function(job) {
      return job.approvalStatus === 'Pending Approval';
    });
  } catch (error) {
    Logger.log("Error in getPendingApprovalJobs: " + error.toString());
    return [];
  }
}

// ฟังก์ชันดึงงานของช่างแต่ละคน
function getJobsByTech(techId) {
  try {
    var jobs = getAllJobs();
    return jobs.filter(function(job) {
      return job.assignedTech === techId && job.status !== 'Closed';
    });
  } catch (error) {
    Logger.log("Error in getJobsByTech: " + error.toString());
    return [];
  }
}

// ฟังก์ชันสำหรับแอดมินสร้างงานใหม่
function createNewJob(customerName, phone, desc, techId) {
  try {
    var ss = getSpreadsheet();
    var jobSheet = ss.getSheetByName('Jobs');
    if (!jobSheet) throw new Error("ไม่พบแท็บ Jobs");
    
    var jobId = generateJobId();
    var reportedAt = new Date();
    
    // โครงสร้างคอลัมน์: 
    // [Job_ID, Customer_Name, Customer_Phone, Issue_Description, Job_Status, Assigned_Tech, Reported_At, Tech_Photo_Before, Tech_Photo_After, Tech_Location, Tech_Note, Closed_At, Approval_Status, Spare_Part_Cost, Executive_Comment]
    var newRow = [
      jobId, 
      customerName, 
      phone, 
      desc, 
      'Open', // Job_Status
      techId, // Assigned_Tech
      reportedAt, // Reported_At
      '', // Tech_Photo_Before
      '', // Tech_Photo_After
      '', // Tech_Location
      '', // Tech_Note
      '', // Closed_At
      'No Request', // Approval_Status
      0, // Spare_Part_Cost
      '' // Executive_Comment
    ];
    
    jobSheet.appendRow(newRow);
    
    // [เพิ่ม Logic PUSH/PULL Notification] ตรวจสอบโหมดก่อนส่ง LINE
    var notifMode = getNotificationMode();
    if (notifMode === 'PUSH') {
      // TODO: เพิ่มโค้ดสำหรับส่ง LINE Messaging API ไปหาช่าง (techId) ตรงนี้
      Logger.log("โหมด PUSH: ส่ง LINE แจ้งเตือนงานใหม่ให้ช่าง " + techId);
    } else {
      // โหมด PULL - ช่างต้องเข้ามากดดูงานเอง ไม่ต้องยิงข้อความ
      Logger.log("โหมด PULL: บันทึกข้อมูลลงชีตเรียบร้อย (ข้ามการส่ง LINE)");
    }
    
    return { success: true, message: "สร้างงาน " + jobId + " สำเร็จ" };
  } catch (error) {
    Logger.log("Error in createNewJob: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ฟังก์ชันบันทึกรูปภาพลง Google Drive และคืนค่า URL
function saveImageToDrive(base64Data, filename) {
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
    if (!folderId) throw new Error("ไม่พบ DRIVE_FOLDER_ID ใน Properties");
    
    var folder = DriveApp.getFolderById(folderId);
    
    // ลบส่วนหัวของ base64 (เช่น data:image/jpeg;base64,) ออกก่อน
    var splitBase = base64Data.split(',');
    var base64String = splitBase.length > 1 ? splitBase[1] : splitBase[0];
    
    var blob = Utilities.newBlob(Utilities.base64Decode(base64String), 'image/jpeg', filename);
    var file = folder.createFile(blob);
    return file.getUrl(); // หรือใช้ getDownloadUrl() ขึ้นอยู่กับความต้องการ
  } catch (error) {
    Logger.log("Error in saveImageToDrive: " + error.toString());
    return "";
  }
}

// ฟังก์ชันสำหรับช่างอัปเดตงาน
function updateJobByTech(jobId, status, photoBase64, lat, lng, note, sparePartCost) {
  try {
    var ss = getSpreadsheet();
    var jobSheet = ss.getSheetByName('Jobs');
    var data = jobSheet.getDataRange().getValues();
    
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === jobId) {
        rowIndex = i + 1; // บวก 1 เพราะ Google Sheets เริ่มที่แถว 1 และ i เริ่มที่ 0 (รวม header เป็นแถวแรก)
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("ไม่พบ Job ID: " + jobId);
    
    var imageUrl = "";
    if (photoBase64 && photoBase64.length > 0) {
      imageUrl = saveImageToDrive(photoBase64, jobId + "_" + status + ".jpg");
    }
    
    var locationStr = (lat && lng) ? (lat + "," + lng) : "";
    var cost = parseFloat(sparePartCost) || 0;
    
    // อัปเดตข้อมูลตามสถานะ
    jobSheet.getRange(rowIndex, 5).setValue(status); // Job_Status
    
    if (status === 'In Progress') {
      if (imageUrl) jobSheet.getRange(rowIndex, 8).setValue(imageUrl); // Tech_Photo_Before
    } else if (status === 'Closed') {
      if (imageUrl) jobSheet.getRange(rowIndex, 9).setValue(imageUrl); // Tech_Photo_After
      jobSheet.getRange(rowIndex, 12).setValue(new Date()); // Closed_At
    }
    
    if (locationStr) jobSheet.getRange(rowIndex, 10).setValue(locationStr); // Tech_Location
    if (note) jobSheet.getRange(rowIndex, 11).setValue(note); // Tech_Note
    
    if (cost > 0) {
      jobSheet.getRange(rowIndex, 14).setValue(cost); // Spare_Part_Cost
      // ตรวจสอบว่ายังไม่อนุมัติ จึงจะเปลี่ยนสถานะเป็น Pending
      var currentApproval = data[rowIndex-1][12];
      if (currentApproval !== 'Approved') {
        jobSheet.getRange(rowIndex, 13).setValue('Pending Approval'); // Approval_Status
      }
    }
    
    return { success: true, message: "อัปเดตงาน " + jobId + " สำเร็จ" };
  } catch (error) {
    Logger.log("Error in updateJobByTech: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ฟังก์ชันสำหรับผู้บริหารอนุมัติงบ
function approveBudget(jobId, isApproved, comment) {
  try {
    var ss = getSpreadsheet();
    var jobSheet = ss.getSheetByName('Jobs');
    var data = jobSheet.getDataRange().getValues();
    
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === jobId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("ไม่พบ Job ID: " + jobId);
    
    var approvalStatus = isApproved ? 'Approved' : 'Rejected';
    
    jobSheet.getRange(rowIndex, 13).setValue(approvalStatus); // Approval_Status
    jobSheet.getRange(rowIndex, 15).setValue(comment || ''); // Executive_Comment
    
    return { success: true, message: "บันทึกการพิจารณาอนุมัติเรียบร้อยแล้ว" };
  } catch (error) {
    Logger.log("Error in approveBudget: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// ส่วนเพิ่มเติม: จัดการโหมดการแจ้งเตือน (PUSH / PULL)
// ==========================================

// ฟังก์ชันดึงโหมดการแจ้งเตือนปัจจุบันจากแท็บ Settings
function getNotificationMode() {
  try {
    var ss = getSpreadsheet();
    var settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) return "PUSH"; // ถ้ายังไม่มีแท็บให้ default เป็น PUSH
    
    var data = settingsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'Notification_Mode') {
        return data[i][1]; // คืนค่า "PUSH" หรือ "PULL"
      }
    }
    return "PUSH"; 
  } catch (error) {
    Logger.log("Error in getNotificationMode: " + error.toString());
    return "PUSH"; 
  }
}

// ฟังก์ชันอัปเดตโหมดการแจ้งเตือน
function updateNotificationMode(mode) {
  try {
    var ss = getSpreadsheet();
    var settingsSheet = ss.getSheetByName('Settings');
    
    // ถ้ายังไม่มีแท็บ Settings ให้สร้างใหม่
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.appendRow(['Setting_Name', 'Setting_Value', 'Description']);
    }
    
    var data = settingsSheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'Notification_Mode') {
        rowIndex = i + 1; // ตำแหน่งแถวใน Sheet (1-index basis)
        break;
      }
    }
    
    if (rowIndex !== -1) {
      settingsSheet.getRange(rowIndex, 2).setValue(mode); // อัปเดตค่าเดิม
    } else {
      settingsSheet.appendRow(['Notification_Mode', mode, 'โหมดการแจ้งเตือน (PUSH หรือ PULL)']); // เพิ่มแถวใหม่
    }
    
    return { success: true, mode: mode };
  } catch (error) {
    Logger.log("Error in updateNotificationMode: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

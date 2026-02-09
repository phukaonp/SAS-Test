// Code.gs

// --- ใส่ ID ของ Folder รูปภาพใน Google Drive ตรงนี้ ---
var IMAGE_FOLDER_ID = "1hx4oS5XilzsjZjLOVPZPSw2tu-SUPndy"; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Neon Batch Work Order & OnSite')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 1. ฟังก์ชันสร้างใบงาน (Batch Create)
function submitBatchData(jobList) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('WorkOrders');
    
    var today = new Date();
    var dateStr = Utilities.formatDate(today, "GMT+7", "yyyyMMdd");
    var informDate = Utilities.formatDate(today, "GMT+7", "yyyy-MM-dd HH:mm");
    
    var lastRow = sheet.getLastRow();
    var startId = lastRow; // รันเลขต่อจากแถวล่าสุด (ถ้ามี Header แถว 1 จะเริ่มที่ 2-1=1)
    if (lastRow === 1 && sheet.getRange(1,1).getValue() === "") startId = 0; // กรณีชีตว่าง
    
    var outputRows = [];
    var resultList = []; 

    for (var i = 0; i < jobList.length; i++) {
      var item = jobList[i];
      var currentNum = startId + i; // Logic การรันเลขอาจต้องปรับตาม Data จริงที่มี
      var idSuffix = ("000" + currentNum).slice(-3);
      var jobId = "JO-" + dateStr + "-" + idSuffix;

      // เรียง Column A-R (1-18) และจอง S-V (19-22) ไว้ว่างๆ
      var row = [
        jobId,              // A: Job ID
        "Pending",          // B: Status
        informDate,         // C: Inform Date
        item.onsiteDate,    // D: Onsite
        item.dueDate,       // E: Due Date
        "",                 // F: Closed Date
        item.site,          // G: Site
        item.building,      // H: Building
        item.floor,         // I: Floor
        item.unit,          // J: Unit
        item.endUser,       // K: EndUser
        item.owner,         // L: Owner
        item.responsible,   // M: Responsible
        item.supplier,      // N: Supplier
        item.scope,         // O: Scope
        item.defectDetail,  // P: Defect Detail
        item.defectCategory,// Q: Defect Category
        item.remark,        // R: Remark
        "", "", "", ""      // S, T, U, V (รูปภาพ)
      ];
      
      outputRows.push(row);
      
      // เก็บข้อมูลส่งกลับไปหน้าเว็บเพื่อทำ On-Site ต่อ
      resultList.push({
        id: jobId,
        defect: item.defectDetail,
        unit: item.unit,
        site: item.site
      });
    }

    if (outputRows.length > 0) {
      sheet.getRange(lastRow + 1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
    }
    
    return { success: true, count: outputRows.length, jobs: resultList };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 2. ฟังก์ชันอัปโหลดรูปภาพ
// ใน Code.gs (แก้ไขเฉพาะฟังก์ชัน uploadImage)

function uploadImage(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('WorkOrders');
    
    // 1. แปลง Base64 เป็นไฟล์
    var contentType = data.mimeType || "image/jpeg";
    var blob = Utilities.newBlob(Utilities.base64Decode(data.base64), contentType, data.fileName);
    
    // 2. บันทึกลง Drive
    var folder = DriveApp.getFolderById(IMAGE_FOLDER_ID); // *อย่าลืมใส่ ID Folder ด้านบนสุดของไฟล์ด้วยนะครับ*
    var file = folder.createFile(blob);
    var fileUrl = file.getUrl();
    
    // 3. หาแถวของ Job ID นั้นเพื่อบันทึก Link
    var textFinder = sheet.getRange("A:A").createTextFinder(data.jobId).matchEntireCell(true);
    var foundRange = textFinder.findNext();
    
    if (foundRange) {
      var row = foundRange.getRow();
      var colIndex = 0;
      
      // --- แก้ไขลำดับ Column ตรงนี้ครับ ---
      switch (data.imgType) {
        case 'unit':   colIndex = 19; break; // Col S: Unit (ลำดับใหม่)
        case 'before': colIndex = 20; break; // Col T: Before
        case 'during': colIndex = 21; break; // Col U: During
        case 'after':  colIndex = 22; break; // Col V: After
      }
      
      if (colIndex > 0) {
        sheet.getRange(row, colIndex).setValue(fileUrl);
      }
      return { success: true, url: fileUrl };
    } else {
      return { success: false, error: "Job ID not found" };
    }

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function forcePermission() {
  DriveApp.getRootFolder();
}

function forceAuth() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
}

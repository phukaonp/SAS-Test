const SPREADSHEET_ID = '1BkhC_02odW8OINve6c3Ec4QI4cr_DEQvFGCVWrgebfg';
const IMAGE_FOLDER_ID = '1pD5dfsyjrtoy7k3IUGaCGPMo6-SiCJPO';

const USER_SHEET_COLUMNS = {
  USER_ID: 1,
  PASSWORD: 2,
  FULL_NAME: 3,
  // Col D(4) = Name (unused)
  ROLE: 5,
  POSITION: 6,
  APPROVED: 7,
  // Col H(8) = SiteResponsible (unused)
  EMAIL: 9,
  LINE: 10,
  PHONE: 11,
  // Col L(12) = reserved
  TEAM: 13,
  TIMESTAMP: 14
};

const USER_ROLE_OPTIONS = ['SAS Staff', 'Supplier'];
const USER_POSITION_OPTIONS = ['Admin', 'Staff', 'Supplier'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAS Defect Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUserSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('User');
  if (!sheet) {
    initSheets();
    sheet = ss.getSheetByName('User');
  }
  return sheet;
}

function normalizeBoolean_(value) {
  if (value === true) return true;
  const text = String(value || '').trim().toLowerCase();
  return text === 'true' || text === '1' || text === 'yes';
}

function normalizeAllowedValue_(value, allowedValues, fieldLabel, allowBlank) {
  const text = String(value == null ? '' : value).trim();
  if (!text) {
    if (allowBlank) return '';
    throw new Error(fieldLabel + ' ไม่สามารถเว้นว่างได้');
  }
  if (allowedValues.indexOf(text) === -1) {
    throw new Error(fieldLabel + ' ไม่ถูกต้อง');
  }
  return text;
}

function getUserTeamOptions_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const teamSheet = ss.getSheetByName('Team');
  if (!teamSheet) return [];

  const lastRow = teamSheet.getLastRow();
  if (lastRow < 2) return [];

  const values = teamSheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
  const teams = values
    .map(function(row) { return String(row[0] || '').trim(); })
    .filter(function(team) { return team !== ''; });

  return teams.filter(function(team, index) {
    return teams.indexOf(team) === index;
  });
}

function getPositionFromRole_(roleValue) {
  const normalizedRole = String(roleValue || '').trim();
  if (normalizedRole === 'Supplier') return 'Supplier';
  if (normalizedRole === 'SAS Staff') return 'Staff';
  return '';
}

function normalizeSerializableValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return value == null ? '' : value;
}

function ensureAdminAccount_() {
  const sheet = getUserSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'phukao') {
      sheet.getRange(i + 1, 2).setValue("'555559");
      sheet.getRange(i + 1, 7).setValue(true);
      sheet.getRange(i + 1, 5).setValue('SAS Staff');
      sheet.getRange(i + 1, 6).setValue('Admin');
      if (!String(data[i][2] || '').trim()) sheet.getRange(i + 1, 3).setValue('phukao');
      return;
    }
  }
  const newRow = new Array(14).fill('');
  newRow[0] = 'phukao';
  newRow[1] = "'555559";
  newRow[2] = 'phukao';
  newRow[4] = 'SAS Staff';
  newRow[5] = 'Admin';
  newRow[6] = true;
  newRow[13] = new Date();
  sheet.appendRow(newRow);
}

function buildUserObject_(row, rowIndex) {
  // การแมปข้อมูลตามคำขอ: 
  // Col E (Index 4) = Role, Col F (Index 5) = Position, Col G (Index 6) = Approved, Col M (Index 12) = Team
  const roleValue = String(row[4] || '').trim();
  const position = String(row[5] || '').trim();
  const approved = normalizeBoolean_(row[6]);
  
  // แปลงให้เป็นพิมพ์เล็กทั้งหมด เพื่อป้องกันปัญหา Case Sensitive เวลาล็อกอิน
  const posLower = position.toLowerCase();
  const isAdmin = posLower === 'admin';
  const isStaff = posLower === 'staff' || posLower === 'sas staff';
  const isSupplier = posLower === 'supplier';
  
  const fullName = String(row[2] || '').trim() || String(row[0] || '').trim();
  let role = 'PENDING';
  if (isAdmin) role = 'ADMIN';
  else if (isStaff) role = 'STAFF';
  else if (isSupplier) role = 'SUPPLIER';

  return {
    rowIndex: rowIndex,
    userId: String(row[0] || '').trim(),
    password: String(row[1] || '').trim().replace(/^'/, ''),
    fullName: fullName,
    roleValue: roleValue,
    position: position,
    isAdmin: isAdmin,
    isStaff: isStaff,
    isSupplier: isSupplier,
    role: role,
    approved: approved,
    email: String(row[8] || '').trim(),
    line: String(row[9] || '').trim(),
    phone: String(row[10] || '').trim(),
    team: String(row[12] || '').trim(),
    timestamp: normalizeSerializableValue_(row[13] || row[14] || ''),
    sheetValues: {
      role: roleValue,
      position: position,
      approved: approved,
      email: String(row[8] || '').trim(),
      line: String(row[9] || '').trim(),
      phone: String(row[10] || '').trim(),
      team: String(row[12] || '').trim()
    }
  };
}

function getUserRowByUserId_(userId) {
  const sheet = getUserSheet_();
  const data = sheet.getDataRange().getValues();
  const inputUserId = String(userId || '').trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === inputUserId) {
      return {
        sheet: sheet,
        rowIndex: i + 1,
        row: data[i],
        user: buildUserObject_(data[i], i + 1)
      };
    }
  }
  return null;
}

function assertAdmin_(actingUserId) {
  ensureAdminAccount_();
  if (String(actingUserId || '').trim().toLowerCase() === 'phukao') return;
  const found = getUserRowByUserId_(actingUserId);
  if (!found || !found.user.isAdmin) {
    throw new Error('ไม่มีสิทธิ์ใช้งานส่วนนี้');
  }
}

function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsInfo = {
    'JOB': ['JobID', 'Site', 'Owner', 'OwnerCompany', 'Staff', 'ReplyDueDate', 'Remark', 'Timestamp', 'Status'],
    'TASK': ['TaskID', 'JobID', 'Scope', 'Building', 'Unit', 'Status', 'CustomerName', 'TargetFixDate', 'ActualStartDate', 'ActualEndDate', 'Duration', 'Remark', 'Timestamp'],
    'DEFECT': [
      'DefectID', 'TaskID', 'TargetStartDate', 'TargetEndDate', 'Status', 'MainCategory', 
      'SubCategory', 'Description', 'Major', 'Team', 'ImgUnit', 'ImgBefore', 'ImgDuring', 'ImgAfter', 'Timestamp', 
      'VOSteps', 'ActualStartDate', 'ActualEndDate', 'Remark' 
    ],
    // อัปเดต Header ของ Sheet User ให้ตรงกับโครงสร้างใหม่
    'User': ['UserID', 'Password', 'FullName', '', 'Role', 'Position', 'Approved', '', 'Email', 'Line', 'Phone', '', 'Team', 'Timestamp'],
    'MainDefect': ['ID', 'MainCategory_Name'],
    'SecondaryDefect': ['ID', 'MainCategory_Ref', 'SubCategory_Name'] // แก้ไขตัวสะกด
  };

  Object.keys(sheetsInfo).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet && name === 'SecondaryDefect') sheet = ss.getSheetByName('SeconadaryDefect');
    
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheetsInfo[name]);
      sheet.getRange(1, 1, 1, sheetsInfo[name].length).setFontWeight("bold").setBackground("#f3f4f6");
    }
  });

  ensureAdminAccount_();
}

function getAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const getSheetData = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    return data.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        if (header) { 
          obj[header] = row[index];
        }
      });
      obj['_raw'] = row; 
      return obj;
    });
  };

  const jobs = getSheetData('JOB');
  const tasks = getSheetData('TASK');
  const defects = getSheetData('DEFECT');

  const structuredJobs = jobs.map(job => {
    const jobTasks = tasks.filter(t => t.JobID === job.JobID).map(task => {
      const taskDefects = defects.filter(d => d.TaskID === task.TaskID).map(def => ({
          id: def.DefectID || def['DefectID'],
          mainCategory: def.MainCategory || def['ลักษณะงานหลัก'],
          subCategory: def.SubCategory || def['ลักษณะงานรอง'],
          description: def.Description || def['รายละเอียด'],
          major: def.Major || def['Major'], 
          team: def.Team || def['ทีมเข้าแก้ไข'],
          imgUnit: def.ImgUnit || def['รูปภาพเลขยูนิต'], 
          imgBefore: def.ImgBefore || def['รูปภาพก่อนแก้ไข'],
          imgDuring: def.ImgDuring || def['รูปภาพระหว่างแก้ไข'],
          imgAfter: def.ImgAfter || def['รูปภาพหลังแก้ไข'],
          status: def.Status || def['DefectStatus'] || def['สถานะ defect'],
          targetStartDate: def.TargetStartDate || def['วันเข้าแก้ไข'] || def['TargetStartDate'] || '',
          targetEndDate: def.TargetEndDate || def['วันแก้ไขเสร็จสิ้น'] || def['TargetEndDate'] || '',
          voSteps: def.VOSteps || def['ขั้นตอนการแก้ไข'] || def['VOSteps'] || '',
          remark: def.Remark || def['หมายเหตุ'] || ''
      }));

      return {
        id: task.TaskID,
        scope: task.Scope,
        building: task.Building,
        unit: task.Unit,
        status: task.Status || task['TaskStatus'] || task['สถานะ'] || task['สถานะใบงาน'] || 'รอดำเนินการ',
        customerName: task._raw[6] || task.CustomerName || task['ชื่อลูกค้า'] || '', 
        targetFixDate: task.TargetFixDate,
        actualStartDate: task.ActualStartDate,
        actualEndDate: task.ActualEndDate,
        duration: task.Duration,
        remark: task.Remark,
        defects: taskDefects
      };
    });
      
    return {
      id: job.JobID,
      site: job.Site,
      owner: job.Owner,
      ownerCompany: job.OwnerCompany,
      staff: job.Staff,
      replyDueDate: job.ReplyDueDate,
      remark: job.Remark,
      status: job.Status || job['JobStatus'] || job['สถานะ'] || job['สถานะใบงานหลัก'] || 'รอดำเนินการ',
      tasks: jobTasks
    };
  });

  return JSON.stringify(structuredJobs);
}

function getAllDataForUser(userId) {
  ensureAdminAccount_();
  const found = getUserRowByUserId_(userId);
  const allJobs = JSON.parse(getAllData());

  if (!found) {
    return JSON.stringify(allJobs);
  }

  if (found.user.isAdmin || found.user.isStaff) {
    return JSON.stringify(allJobs);
  }

  if (!found.user.isSupplier) {
    return JSON.stringify([]);
  }

  const team = String(found.user.team || '').trim();
  if (!team) {
    return JSON.stringify([]);
  }

  const filteredJobs = allJobs
    .map(job => ({
      ...job,
      tasks: (job.tasks || [])
        .map(task => ({
          ...task,
          defects: (task.defects || []).filter(defect => String(defect.team || '').trim() === team)
        }))
        .filter(task => (task.defects || []).length > 0)
    }))
    .filter(job => (job.tasks || []).length > 0);

  return JSON.stringify(filteredJobs);
}

function addJob(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const data = sheet.getDataRange().getValues();
  
  const siteStr = formData.site || 'UNKNOWN';
  let maxNum = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === siteStr) { 
      const id = data[i][0];
      const parts = id.split('-');
      if (parts.length >= 3) {
        const num = parseInt(parts[parts.length - 1], 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }
  
  const newNumStr = String(maxNum + 1).padStart(4, '0');
  const newId = `JOB-${siteStr}-${newNumStr}`;
  
  sheet.appendRow([
    newId,                        
    formData.site || '',          
    formData.owner || '',         
    formData.ownerCompany || '',  
    formData.staff || '',         
    formData.replyDueDate || '',  
    formData.remark || '',        
    new Date(),                   
    'รอดำเนินการ'                   
  ]);
  return newId;
}

function addTask(jobId, formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('TASK');
  const data = sheet.getDataRange().getValues();
  
  let maxNum = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === jobId) { 
      const id = data[i][0];
      const parts = id.split('-T');
      if (parts.length >= 2) {
        const num = parseInt(parts[parts.length - 1], 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }

  const newNumStr = String(maxNum + 1).padStart(3, '0');
  const newId = `${jobId}-T${newNumStr}`;

  const historyLog = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy, HH:mm:ss');
  
  let durationDate = '';
  if (formData.targetFixDate) {
    const parts = formData.targetFixDate.split('-'); 
    if (parts.length === 3) {
      const dateObj = new Date(parts[0], parts[1] - 1, parts[2]);
      dateObj.setDate(dateObj.getDate() + 14);
      durationDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd'); 
    }
  }

  sheet.appendRow([
    newId,                        
    jobId,                        
    formData.scope || 'SAS',      
    formData.building || '',      
    formData.unit || '',          
    'รอดำเนินการ',                  
    formData.customerName || '',  
    formData.targetFixDate || '', 
    durationDate,                 
    formData.remark || '',        
    historyLog                    
  ]);
  
  return newId;
}

function addDefect(taskId, defectData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();

  let maxNum = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === taskId) { 
      const id = data[i][0];
      const parts = id.split('-DF');
      if (parts.length >= 2) {
        const num = parseInt(parts[parts.length - 1], 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }

  const newNumStr = String(maxNum + 1).padStart(3, '0');
  const newId = `${taskId}-DF${newNumStr}`;

  function uploadBase64(base64Str, filename) {
    if (!base64Str) return '';
    try {
      const splitBase = base64Str.split(',');
      const contentType = splitBase[0].split(';')[0].replace('data:', '');
      const byteCharacters = Utilities.base64Decode(splitBase[1]);
      const blob = Utilities.newBlob(byteCharacters, contentType, filename);
      const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return "https://drive.google.com/uc?export=view&id=" + file.getId();
    } catch (e) {
      return '';
    }
  }

  let imgBeforeUrl = '';
  if (defectData.imgBefore) {
    const ts = new Date().getTime();
    imgBeforeUrl = uploadBase64(defectData.imgBefore, `Before_${newId}_${ts}`);
  }

  const rowData = new Array(19).fill(''); 
  
  rowData[0] = newId;                        
  rowData[1] = taskId;                       
  rowData[2] = defectData.targetStartDate || ''; 
  rowData[3] = defectData.targetEndDate || '';   
  rowData[4] = 'ยังไม่แก้ไข';                  
  rowData[5] = defectData.mainCategory;      
  rowData[6] = defectData.subCategory;       
  rowData[7] = defectData.description;       
  rowData[8] = defectData.major;             
  rowData[9] = defectData.team;              
  rowData[10] = '';                          
  rowData[11] = imgBeforeUrl;                
  rowData[12] = '';                          
  rowData[13] = '';                          
  rowData[14] = new Date();                  
  rowData[15] = defectData.voSteps || '';    

  sheet.appendRow(rowData);
  return newId;
}

function updateTask(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('TASK');
  const data = sheet.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === formData.id) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('ไม่พบ TaskID ที่ต้องการแก้ไขในฐานข้อมูล');
  }

  const currentStatus = String(data[rowIndex - 1][5] || '').trim();
  if (currentStatus === 'Closed') {
    throw new Error('ไม่สามารถแก้ไข Task ที่สถานะ Closed ได้');
  }

  sheet.getRange(rowIndex, 3).setValue(formData.scope || 'SAS');
  sheet.getRange(rowIndex, 4).setValue(formData.building || '');
  sheet.getRange(rowIndex, 5).setValue(formData.unit || '');
  sheet.getRange(rowIndex, 7).setValue(formData.customerName || '');
  sheet.getRange(rowIndex, 8).setValue(formData.targetFixDate || '');

  let durationDate = '';
  if (formData.targetFixDate) {
    const parts = String(formData.targetFixDate).split('-');
    if (parts.length === 3) {
      const dateObj = new Date(parts[0], parts[1] - 1, parts[2]);
      dateObj.setDate(dateObj.getDate() + 14);
      durationDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  sheet.getRange(rowIndex, 9).setValue(durationDate);
  sheet.getRange(rowIndex, 10).setValue(formData.remark || '');
  sheet.getRange(rowIndex, 11).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy, HH:mm:ss'));

  SpreadsheetApp.flush();
  return 'Update Success';
}

function updateDefect(defectData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectData.id) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('ไม่พบ DefectID ที่ต้องการแก้ไขในฐานข้อมูล');
  }

  const currentStatus = String(data[rowIndex - 1][4] || '').trim();
  if (currentStatus === 'แก้ไขแล้ว') {
    throw new Error('ไม่สามารถแก้ไข Defect ที่สถานะ แก้ไขแล้ว ได้');
  }

  const teamValue = normalizeAllowedValue_(defectData.team, getUserTeamOptions_(), 'Team', false);

  sheet.getRange(rowIndex, 3).setValue(defectData.targetStartDate || '');
  sheet.getRange(rowIndex, 4).setValue(defectData.targetEndDate || '');
  sheet.getRange(rowIndex, 6).setValue(defectData.mainCategory || '');
  sheet.getRange(rowIndex, 7).setValue(defectData.subCategory || '');
  sheet.getRange(rowIndex, 8).setValue(defectData.description || '');
  sheet.getRange(rowIndex, 9).setValue(defectData.major || 'ไม่ใช่');
  sheet.getRange(rowIndex, 10).setValue(teamValue);
  sheet.getRange(rowIndex, 16).setValue(defectData.voSteps || '');

  SpreadsheetApp.flush();
  return 'Update Success';
}

function updateJob(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === formData.id) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error("ไม่พบ JobID ที่ต้องการแก้ไขในฐานข้อมูล");
  }

  sheet.getRange(rowIndex, 2).setValue(formData.site || '');
  sheet.getRange(rowIndex, 3).setValue(formData.owner || '');
  sheet.getRange(rowIndex, 4).setValue(formData.ownerCompany || '');
  sheet.getRange(rowIndex, 5).setValue(formData.staff || '');
  sheet.getRange(rowIndex, 6).setValue(formData.replyDueDate || '');
  sheet.getRange(rowIndex, 7).setValue(formData.remark || '');

  return "Update Success";
}

function getMasterData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let result = { 
    sites: [], 
    owners: [],
    mainCategories: [],
    subCategories: {}
  };

  const projectSheet = ss.getSheetByName('Project');
  if (projectSheet) {
    const pLastRow = projectSheet.getLastRow();
    if (pLastRow >= 2) {
      const pData = projectSheet.getRange(2, 2, pLastRow - 1, 1).getDisplayValues();
      const sites = pData.map(r => r[0]).filter(s => s !== '');
      result.sites = [...new Set(sites)];
    }
  }

  const ownerSheet = ss.getSheetByName('Owner');
  if (ownerSheet) {
    const oLastRow = ownerSheet.getLastRow();
    if (oLastRow >= 2) {
      const oData = ownerSheet.getRange(2, 2, oLastRow - 1, 3).getDisplayValues();
      result.owners = oData
        .filter(row => row[2] !== '') 
        .map(row => ({
          ownerCompany: row[0], 
          site: row[1],         
          owner: row[2]         
        }));
    }
  }

  const mainDefectSheet = ss.getSheetByName('MainDefect');
  if (mainDefectSheet) {
    const mLastRow = mainDefectSheet.getLastRow();
    if (mLastRow >= 2) {
      const mData = mainDefectSheet.getRange(2, 2, mLastRow - 1, 1).getDisplayValues();
      const mCategories = mData.map(r => r[0].toString().trim()).filter(c => c !== '');
      result.mainCategories = [...new Set(mCategories)];
    }
  }

  const teamSheet = ss.getSheetByName('Team');
  if (teamSheet) {
    const tLastRow = teamSheet.getLastRow();
    if (tLastRow >= 2) {
      // ดึงข้อมูลจาก Column C (Index 3) ของ Sheet Team
      const tData = teamSheet.getRange(2, 3, tLastRow - 1, 1).getDisplayValues();
      const teamsList = tData.map(r => r[0].toString().trim()).filter(t => t !== '');
      result.teams = [...new Set(teamsList)];
    }
  }

  const subDefectSheet = ss.getSheetByName('SecondaryDefect') || ss.getSheetByName('SeconadaryDefect'); 
  if (subDefectSheet) {
    const sLastRow = subDefectSheet.getLastRow();
    if (sLastRow >= 2) {
      const sData = subDefectSheet.getRange(2, 1, sLastRow - 1, 4).getDisplayValues();
      sData.forEach(row => {
        const mainCat = row[1] ? row[1].toString().trim() : ''; 
        const subCat = row[3] ? row[3].toString().trim() : '';  
        
        if (mainCat && subCat) {
          if (!result.subCategories[mainCat]) {
            result.subCategories[mainCat] = [];
          }
          if (!result.subCategories[mainCat].includes(subCat)) {
            result.subCategories[mainCat].push(subCat);
          }
        }
      });
    }
  }

  return JSON.stringify(result);
}

function deleteJob(jobId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('JOB');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  throw new Error("ไม่พบ JobID ที่ต้องการลบ");
}

function deleteTask(taskId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('TASK');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  throw new Error("ไม่พบ TaskID ที่ต้องการลบ");
}

function deleteDefect(defectId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  throw new Error("ไม่พบ DefectID ที่ต้องการลบ");
}

function updateTaskStatusAndJob(taskId, newStatus) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('TASK');
  const taskData = taskSheet.getDataRange().getValues();

  let jobId = '';
  let taskRowIndex = -1;

  for (let i = 1; i < taskData.length; i++) {
    if (taskData[i][0] === taskId) {
      taskRowIndex = i + 1;
      jobId = taskData[i][1]; 
      taskData[i][5] = newStatus; 
      break;
    }
  }
  
  if (taskRowIndex !== -1) {
    taskSheet.getRange(taskRowIndex, 6).setValue(newStatus);

    if (jobId) {
       const jobSheet = ss.getSheetByName('JOB');
       const jobData = jobSheet.getDataRange().getValues();
       
       let allTasksFinished = true;
       for (let i = 1; i < taskData.length; i++) {
         if (taskData[i][1] === jobId) {
           const status = taskData[i][5];
           if (status === 'รอดำเนินการ' || status === 'Active' || status === '') {
             allTasksFinished = false;
             break;
           }
         }
       }

       let jobRowIndex = -1;
       for (let j = 1; j < jobData.length; j++) {
         if (jobData[j][0] === jobId) {
           jobRowIndex = j + 1;
           break;
         }
       }

       if (jobRowIndex !== -1) {
         if (allTasksFinished) {
           jobSheet.getRange(jobRowIndex, 9).setValue('Closed');
         } else if (newStatus === 'Active') {
           if (jobData[jobRowIndex - 1][8] !== 'Active') { 
             jobSheet.getRange(jobRowIndex, 9).setValue('Active');
           }
         }
       }
    }
    
    SpreadsheetApp.flush();
    return "Success";
  }
  throw new Error("ไม่พบข้อมูลใบงานย่อยที่ต้องการเปลี่ยนสถานะ");
}

function uploadSingleDefectImage(defectId, field, base64Str) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Defect not found");

  if (!base64Str) return '';
  if (base64Str.startsWith('http')) return base64Str; 

  try {
    const splitBase = base64Str.split(',');
    const contentType = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    
    const ts = new Date().getTime();
    const filename = `${field}_${defectId}_${ts}`;
    const blob = Utilities.newBlob(byteCharacters, contentType, filename);
    
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    const file = folder.createFile(blob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
    
    const colMap = { 'imgUnit': 11, 'imgBefore': 12, 'imgDuring': 13, 'imgAfter': 14 };
    if (colMap[field]) {
      sheet.getRange(rowIndex, colMap[field]).setValue(url);
    }
    return url;
  } catch (e) {
    throw new Error('Upload failed: ' + e.toString());
  }
}

function getDefectImagePreviewDataUrl(imageUrl) {
  if (!imageUrl) return '';
  if (String(imageUrl).indexOf('data:image') === 0) return imageUrl;

  const match = String(imageUrl).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(imageUrl).match(/id=([a-zA-Z0-9_-]+)/);
  if (!match || !match[1]) return imageUrl;

  try {
    const file = DriveApp.getFileById(match[1]);
    const blob = file.getBlob();
    const contentType = blob.getContentType() || 'image/jpeg';
    const base64 = Utilities.base64Encode(blob.getBytes());
    return `data:${contentType};base64,${base64}`;
  } catch (e) {
    throw new Error('Preview image load failed: ' + e.toString());
  }
}

function updateDefectStatus(defectId, status) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DEFECT');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defectId) {
      sheet.getRange(i + 1, 5).setValue(status);
      return "Success";
    }
  }
  throw new Error("ไม่พบข้อมูล DefectID ที่ต้องการเปลี่ยนสถานะ");
}

function getPrintableImgUrlForPdf_(url) {
  if (!url) return '';
  if (String(url).startsWith('data:image')) return url;
  const match = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(url).match(/id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    try {
      const file = DriveApp.getFileById(match[1]);
      const blob = file.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      const mimeType = blob.getContentType() || 'image/jpeg';
      return `data:${mimeType};base64,${base64}`;
    } catch (e) {
      return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1000`;
    }
  }
  return url;
}

function getPdfLogoDataUrl_() {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="72" height="72" viewBox="0 0 24 24" fill="none" stroke="%230D504C" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg>';
  return 'data:image/svg+xml;utf8,' + svg;
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildTaskPlanPdfHtml_(job, task) {
  const logoDataUrl = getPdfLogoDataUrl_();
  const defects = (task.defects || []).filter(def => (def.status || 'รอดำเนินการ') === 'รอดำเนินการ');
  const featuredDefect = defects[0] || null;
  const featuredImage = featuredDefect ? (getPrintableImgUrlForPdf_(featuredDefect.imgBefore) || getPrintableImgUrlForPdf_(featuredDefect.imgUnit)) : '';
  const featuredRepairDate = featuredDefect
    ? (featuredDefect.actualStartDate || featuredDefect.targetStartDate || task.actualStartDate || task.targetFixDate || '')
    : '';
  const ownerName = escapeHtml_(job.owner || '-');
  const staffName = escapeHtml_(job.staff || '-');

  let html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
      <style>
        * { box-sizing: border-box; }
        @page { size: A4; margin: 12mm; }
        body { font-family: 'Sarabun', sans-serif; color: #111111; line-height: 1.45; font-size: 11px; margin: 0; padding: 0; background: #ffffff; }
        .sheet { width: 100%; min-height: 272mm; padding: 12mm 12mm 10mm; display: flex; flex-direction: column; }
        .hero { padding: 0 0 8mm; color: #111111; }
        .hero-grid { width: 100%; border-collapse: collapse; }
        .hero-grid td { vertical-align: top; }
        .brand { width: 90px; }
        .brand-badge { width: 60px; height: 38px; text-align: left; }
        .brand-badge img { width: 40px; height: 40px; object-fit: contain; }
        .hero-main { text-align: center; padding-top: 2px; }
        .hero-title { font-size: 18px; font-weight: 500; letter-spacing: 0; margin: 0; }
        .hero-subtitle { font-size: 10px; margin-top: 2px; }
        .hero-ids { width: 115px; text-align: right; }
        .hero-label { font-size: 9px; color: #555555; }
        .hero-value { font-size: 11px; font-weight: 500; margin-top: 2px; min-height: 16px; }
        .body { flex: 1; display: flex; flex-direction: column; }
        .section-title { font-size: 12px; font-weight: 700; color: #111111; margin: 0 0 4mm; text-transform: uppercase; }
        .info-table { width: 100%; border-collapse: collapse; margin-bottom: 6mm; }
        .info-table td { padding: 2.2mm 0; vertical-align: top; font-size: 10px; }
        .info-label { width: 16%; font-weight: 400; white-space: nowrap; }
        .info-value { width: 34%; font-weight: 400; padding-left: 2mm; }
        .detail-card { page-break-inside: avoid; }
        .detail-shell { width: 100%; border-collapse: collapse; table-layout: fixed; }
        .detail-shell td { vertical-align: top; }
        .detail-visual { width: 40%; padding-right: 8mm; }
        .detail-meta { width: 60%; }
        .image-box { border: 1px solid #222222; height: 92mm; padding: 6mm; background: #ffffff; text-align: center; display: flex; align-items: center; justify-content: center; }
        .image-box img { max-width: 100%; max-height: 78mm; object-fit: contain; }
        .no-image { width: 100%; height: 78mm; color: #444444; font-size: 12px; display: flex; align-items: center; justify-content: center; }
        .detail-caption { margin-top: 4mm; text-align: center; font-size: 10px; }
        .meta-line { margin-bottom: 8mm; }
        .meta-inline { font-size: 10px; font-weight: 600; }
        .meta-title { font-size: 11px; font-weight: 700; margin: 0 0 2mm; }
        .meta-body { font-size: 10px; min-height: 18mm; white-space: pre-wrap; word-break: break-word; }
        .meta-small { font-size: 10px; font-weight: 600; min-height: 8mm; }
        .signature-wrap { margin-top: auto; padding-top: 18mm; page-break-inside: avoid; }
        .signature-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
        .signature-table td { width: 50%; vertical-align: bottom; }
        .signature-table td:last-child { padding-left: 18mm; }
        .signature-block { min-height: 34mm; }
        .signature-title { font-size: 11px; font-weight: 400; text-transform: uppercase; margin-bottom: 17mm; }
        .signature-line { border-bottom: 1px solid #111111; height: 1px; }
        .signature-name { font-size: 10px; margin-top: 4mm; }
        .signature-date { font-size: 10px; margin-top: 1.5mm; }
      </style>
    </head>
    <body>
      <div class="sheet">
        <div class="hero">
          <table class="hero-grid">
            <tr>
              <td class="brand">
                <div class="brand-badge"><img src="${logoDataUrl}" /></div>
              </td>
              <td class="hero-main">
                <div class="hero-title">Repair Planning Document</div>
                <div class="hero-subtitle">เอกสารแผนเข้าแก้ไข</div>
              </td>
              <td class="hero-ids">
                <div class="hero-label">Job ID:</div>
                <div class="hero-value">${escapeHtml_(job.id)}</div>
                <div class="hero-label" style="margin-top: 4px;">Task ID:</div>
                <div class="hero-value">${escapeHtml_(task.id)}</div>
              </td>
            </tr>
          </table>
        </div>
        <div class="body">
          <div class="section-title">Information</div>
          <table class="info-table">
            <tr>
              <td class="info-label">ชื่อลูกค้า :</td><td class="info-value">${escapeHtml_(task.customerName || '-')}</td>
              <td class="info-label">SITE :</td><td class="info-value">${escapeHtml_(job.site || '-')}</td>
            </tr>
            <tr>
              <td class="info-label">Owner / ผู้ดูแล :</td><td class="info-value">${ownerName}</td>
              <td class="info-label">Building - Unit :</td><td class="info-value">${escapeHtml_(`${task.building || '-'} - ${task.unit || '-'}`)}</td>
            </tr>
            <tr>
              <td class="info-label">Staff/ผู้รับผิดชอบ :</td><td class="info-value">${staffName}</td>
              <td class="info-label"></td><td class="info-value"></td>
            </tr>
          </table>
          <div class="section-title">Detail</div>
  `;

  if (featuredDefect) {
    html += `
          <div class="detail-card">
            <table class="detail-shell">
              <tr>
                <td class="detail-visual">
                  <div class="image-box">
                    ${featuredImage ? `<img src="${featuredImage}" />` : `<div class="no-image">รูปภาพก่อนแก้ไข</div>`}
                  </div>
                  <div class="detail-caption">รูปภาพก่อนแก้ไข</div>
                </td>
                <td class="detail-meta">
                  <div class="meta-line">
                    <div class="meta-inline">ลักษณะงานหลัก - ลักษณะงานรอง</div>
                    <div class="meta-body">${escapeHtml_(`${featuredDefect.mainCategory || '-'} - ${featuredDefect.subCategory || '-'}`)}</div>
                  </div>
                  <div class="meta-line">
                    <div class="meta-title">รายละเอียด Defect</div>
                    <div class="meta-body">${escapeHtml_(featuredDefect.description || '-')}</div>
                  </div>
                  <div class="meta-line">
                    <div class="meta-title">ทีมที่เข้าแก้</div>
                    <div class="meta-small">${escapeHtml_(featuredDefect.team || '-')}</div>
                  </div>
                  <div class="meta-line">
                    <div class="meta-title">กำหนดวันที่เข้าแก้ไข</div>
                    <div class="meta-small">${escapeHtml_(featuredRepairDate || 'ยังไม่ได้ระบุ')}</div>
                  </div>
                  <div class="meta-line" style="margin-bottom: 0;">
                    <div class="meta-title">Major</div>
                    <div class="meta-small">${escapeHtml_(featuredDefect.major || 'ไม่ใช่')}</div>
                  </div>
                </td>
              </tr>
            </table>
          </div>
    `;
  } else {
    html += `
          <div class="detail-card">
            <table class="detail-shell">
              <tr>
                <td class="detail-visual">
                  <div class="image-box"><div class="no-image">รูปภาพก่อนแก้ไข</div></div>
                  <div class="detail-caption">รูปภาพก่อนแก้ไข</div>
                </td>
                <td class="detail-meta">
                  <div class="meta-line">
                    <div class="meta-inline">ลักษณะงานหลัก - ลักษณะงานรอง</div>
                    <div class="meta-body">-</div>
                  </div>
                  <div class="meta-line">
                    <div class="meta-title">รายละเอียด Defect</div>
                    <div class="meta-body">ไม่มีรายการ Defect สถานะรอดำเนินการในใบงานย่อยนี้</div>
                  </div>
                  <div class="meta-line">
                    <div class="meta-title">ทีมที่เข้าแก้</div>
                    <div class="meta-small">-</div>
                  </div>
                  <div class="meta-line">
                    <div class="meta-title">กำหนดวันที่เข้าแก้ไข</div>
                    <div class="meta-small">-</div>
                  </div>
                  <div class="meta-line" style="margin-bottom: 0;">
                    <div class="meta-title">Major</div>
                    <div class="meta-small">-</div>
                  </div>
                </td>
              </tr>
            </table>
          </div>
    `;
  }

  html += `
          <div class="signature-wrap">
            <table class="signature-table">
              <tr>
                <td>
                  <div class="signature-block">
                    <div class="signature-title">Staff Signature</div>
                    <div class="signature-line"></div>
                    <div class="signature-name">( ${staffName} )</div>
                    <div class="signature-date">วันที่: ......../......../..............</div>
                  </div>
                </td>
                <td>
                  <div class="signature-block">
                    <div class="signature-title">Owner Signature</div>
                    <div class="signature-line"></div>
                    <div class="signature-name">( ${ownerName} )</div>
                    <div class="signature-date">วันที่: ......../......../..............</div>
                  </div>
                </td>
              </tr>
            </table>
          </div>
        </div>
      </div>
    </body>
  </html>`;

  return html;
}

function exportTaskPlansToPDF(jobId, taskId) {
  const allDataStr = getAllData();
  const allJobs = JSON.parse(allDataStr);
  const job = allJobs.find(j => j.id === jobId);

  if (!job) throw new Error('ไม่พบข้อมูลใบงานหลัก (Job)');
  if (!job.tasks || job.tasks.length === 0) throw new Error('ไม่มีใบงานย่อยให้ Export');

  const pendingTasks = job.tasks.filter(task => (task.status || 'รอดำเนินการ') === 'รอดำเนินการ');
  if (pendingTasks.length === 0) throw new Error('ไม่มีใบงานย่อยสถานะรอดำเนินการสำหรับ Export');

  const requestedTaskIds = Array.isArray(taskId)
    ? taskId.filter(Boolean)
    : (taskId ? [taskId] : []);

  const selectedTasks = requestedTaskIds.length > 0
    ? pendingTasks.filter(task => requestedTaskIds.indexOf(task.id) !== -1)
    : pendingTasks;

  if (selectedTasks.length === 0) throw new Error('ไม่พบใบงานย่อยที่เลือก หรือใบงานนั้นไม่ได้อยู่ในสถานะรอดำเนินการ');

  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  const exportedFiles = [];

  selectedTasks.forEach(task => {
    const html = buildTaskPlanPdfHtml_(job, task);
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(`RepairPlan_${task.id}.pdf`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    exportedFiles.push({ taskId: task.id, url: file.getUrl(), name: file.getName() });
  });

  return JSON.stringify(exportedFiles);
}

function exportDefectReportToPDF(taskId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const allDataStr = getAllData();
  
  const allJobs = JSON.parse(allDataStr);
  
  let targetTask = null;
  let targetJob = null;
  
  for (const job of allJobs) {
    const foundTask = job.tasks.find(t => t.id === taskId);
    if (foundTask) {
      targetTask = foundTask;
      targetJob = job;
      break;
    }
  }

  if (!targetTask) throw new Error("ไม่พบข้อมูลใบงานย่อย (Task)");

  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  
  const getPrintableImgUrl = (url) => {
    if (!url) return '';
    if (url.startsWith('data:image')) return url;
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      try {
        const file = DriveApp.getFileById(match[1]);
        const blob = file.getBlob();
        const base64 = Utilities.base64Encode(blob.getBytes());
        const mimeType = blob.getContentType();
        return `data:${mimeType};base64,${base64}`;
      } catch (e) {
        return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w500`;
      }
    }
    return url;
  };

  let html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Sarabun', sans-serif; color: #1e293b; line-height: 1.6; font-size: 14px; margin: 0; padding: 10px; }
          .header-title { text-align: center; color: #0f172a; margin-bottom: 25px; font-size: 24px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
          table.header-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
          table.header-table th, table.header-table td { border: 1px solid #cbd5e1; padding: 10px 12px; text-align: left; vertical-align: top; }
          table.header-table th { background-color: #f1f5f9; width: 18%; font-weight: 600; color: #334155; }
          table.header-table td { width: 32%; color: #0f172a; }
          .section-title { font-size: 18px; font-weight: 600; color: #1e40af; border-bottom: 2px solid #93c5fd; padding-bottom: 8px; margin-bottom: 15px; }
          
          .defect-card { border: 1px solid #e2e8f0; margin-bottom: 25px; padding: 15px; page-break-inside: avoid; border-radius: 8px; background-color: #ffffff; }
          .defect-info { margin-bottom: 15px; padding: 12px; background-color: #f8fafc; border-radius: 6px; border-left: 4px solid #3b82f6; }
          .defect-info strong { color: #0f172a; }
          
          .img-grid { display: table; width: 100%; table-layout: fixed; margin-top: 15px; }
          .img-cell { display: table-cell; width: 25%; padding: 0 5px; text-align: center; vertical-align: top; }
          .img-cell img { width: 100%; max-height: 180px; object-fit: contain; border: 1px solid #cbd5e1; border-radius: 6px; padding: 2px; }
          .img-label { font-size: 13px; font-weight: 700; margin-bottom: 8px; color: #1e40af; background-color: #eff6ff; padding: 4px 0; border-radius: 4px; }
          .no-img { height: 120px; background: #f8fafc; border: 2px dashed #cbd5e1; border-radius: 6px; display: flex; align-items: center; justify-content: center; color: #94a3b8; font-size: 12px; margin-top: 5px; }

          .signature-container {
              width: 100%;
              margin-top: 50px;
              page-break-inside: avoid;
          }
          .signature-table {
              width: 100%;
              border-collapse: collapse;
              text-align: center;
          }
          .signature-table td {
              width: 50%;
              padding: 10px 20px;
              vertical-align: bottom;
          }
          .sign-line {
              border-bottom: 1px dashed #94a3b8;
              width: 70%;
              margin: 40px auto 10px auto;
          }
          .sign-text {
              color: #334155;
              font-size: 14px;
              line-height: 1.5;
          }
          .sign-name {
              font-weight: 600;
              color: #0f172a;
          }
        </style>
      </head>
      <body>
        <div class="header-title">เอกสารแก้ไข Defect</div>
        <table class="header-table">
          <tr>
            <th>Job ID</th><td>${targetJob.id}</td>
            <th>Task ID</th><td>${targetTask.id}</td>
          </tr>
          <tr>
            <th>Site</th><td>${targetJob.site}</td>
            <th>Scope</th><td>${targetTask.scope}</td>
          </tr>
          <tr>
            <th>Owner / ผู้ดูแล</th><td>${targetJob.owner}</td>
            <th>Building / Unit</th><td>${targetTask.building} - ${targetTask.unit}</td>
          </tr>
          <tr>
            <th>Company</th><td>${targetJob.ownerCompany || '-'}</td>
            <th>ชื่อลูกค้า</th><td>${targetTask.customerName || '-'}</td>
          </tr>
          <tr>
            <th>Staff / ผู้จัดทำ</th><td colspan="3">${targetJob.staff || '-'}</td>
          </tr>
        </table>

        <div class="section-title">รายละเอียดผลการแก้ไข Defect</div>
  `;

  if (targetTask.defects && targetTask.defects.length > 0) {
    targetTask.defects.forEach((def) => {
      let imgUnit = getPrintableImgUrl(def.imgUnit);
      let imgBefore = getPrintableImgUrl(def.imgBefore);
      let imgDuring = getPrintableImgUrl(def.imgDuring);
      let imgAfter = getPrintableImgUrl(def.imgAfter);

      const renderImg = (src, label) => `
        <div class="img-cell">
          <div class="img-label">${label}</div>
          ${src ? `<img src="${src}" />` : `<div class="no-img">ไม่มีรูปภาพ</div>`}
        </div>
      `;

      html += `
      <div class="defect-card">
        <div class="defect-info">
          <div style="margin-bottom: 6px;">
            <strong>สถานะ:</strong> <span style="color: #047857; font-weight: 600;">${def.status}</span> &nbsp;|&nbsp; 
            <strong>ลักษณะงานหลัก:</strong> ${def.mainCategory} &nbsp;|&nbsp; 
            <strong>ลักษณะงานรอง:</strong> ${def.subCategory}
          </div>
          <div style="margin-bottom: 6px;">
            <strong>ทีมเข้าแก้ไข:</strong> ${def.team}
          </div>
          <div>
            <strong>รายละเอียด:</strong> ${def.description}
          </div>
        </div>
        
        <div class="img-grid">
          ${renderImg(imgUnit, '1. รูปภาพเลขยูนิต')}
          ${renderImg(imgBefore, '2. รูปภาพก่อนแก้ไข')}
          ${renderImg(imgDuring, '3. รูปภาพระหว่างแก้ไข')}
          ${renderImg(imgAfter, '4. รูปภาพหลังแก้ไข')}
        </div>
      </div>
      `;
    });
  } else {
    html += `<p style="text-align:center; color:#94a3b8; padding: 30px 0; font-style: italic;">- ไม่มีรายการ Defect ในใบงานย่อยนี้ -</p>`;
  }

  html += `
        <div class="signature-container">
            <table class="signature-table">
                <tr>
                    <td>
                        <div class="sign-line"></div>
                        <div class="sign-text sign-name">( ${targetJob.owner || '.........................................................'} )</div>
                        <div class="sign-text">ผู้อนุมัติ (Owner)</div>
                        <div class="sign-text" style="margin-top: 5px;">วันที่: ......../......../..............</div>
                    </td>
                    <td>
                        <div class="sign-line"></div>
                        <div class="sign-text sign-name">( ${targetTask.customerName || '.........................................................'} )</div>
                        <div class="sign-text">ลูกค้า (Customer)</div>
                        <div class="sign-text" style="margin-top: 5px;">วันที่: ......../......../..............</div>
                    </td>
                </tr>
            </table>
        </div>
      </body>
    </html>
  `;

  const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(`DefectReport_${targetTask.id}.pdf`);
  const file = folder.createFile(blob);
  return JSON.stringify({ taskId: targetTask.id, url: file.getUrl() });
}

function registerUser(formData) {
  ensureAdminAccount_();
  const sheet = getUserSheet_();
  
  const data = sheet.getDataRange().getValues();
  const inputUserId = String(formData.userId || '').trim().toLowerCase();
  const fullName = String(formData.fullName || '').trim();
  const password = String(formData.password || '').trim();
  const confirmPassword = String(formData.confirmPassword == null ? formData.password : formData.confirmPassword).trim();
  const email = String(formData.email || '').trim();
  const line = String(formData.line || '').trim();
  const phone = String(formData.phone || '').trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === inputUserId) {
      throw new Error('User ID นี้มีผู้ใช้งานแล้ว กรุณาใช้ชื่ออื่น');
    }
  }

  if (!inputUserId) throw new Error('กรุณากรอก User ID');
  if (!fullName) throw new Error('กรุณากรอกชื่อ - นามสกุล');
  if (!password) throw new Error('กรุณากรอก Password');
  if (!confirmPassword) throw new Error('กรุณากรอก Confirm Password');
  if (password !== confirmPassword) {
    throw new Error('Password และ Confirm Password ไม่ตรงกัน');
  }

  const roleValue = normalizeAllowedValue_(formData.roleValue, USER_ROLE_OPTIONS, 'Role', false);
  const positionValue = getPositionFromRole_(roleValue);

  const newRow = new Array(15).fill('');
  newRow[0] = String(formData.userId || '').trim();
  newRow[1] = "'" + confirmPassword;
  newRow[2] = fullName;
  newRow[4] = roleValue;
  newRow[5] = positionValue;
  newRow[6] = false;
  newRow[8] = email;
  newRow[9] = line;
  newRow[10] = phone;
  newRow[12] = '';
  newRow[13] = new Date();

  sheet.appendRow(newRow);
  SpreadsheetApp.flush(); // บังคับให้ระบบเขียนข้อมูลลง Sheet ทันที
  
  return 'Success';
}

function loginUser(userId, password) {
  ensureAdminAccount_();
  const sheet = getUserSheet_();
  if (!sheet) throw new Error('ไม่พบฐานข้อมูลผู้ใช้งาน กรุณาติดต่อผู้ดูแลระบบ');

  const data = sheet.getDataRange().getValues();
  
  const inputUserId = String(userId).trim().toLowerCase(); // เปลี่ยนเป็น toLowerCase ป้องกัน Case Sensitive
  const inputPassword = String(password).trim();

  for (let i = 1; i < data.length; i++) {
    const sheetUserId = String(data[i][0]).trim().toLowerCase(); // เปลี่ยนเป็น toLowerCase เพื่อเปรียบเทียบ
    let sheetPassword = String(data[i][1]).trim();

    if (sheetPassword.startsWith("'")) {
        sheetPassword = sheetPassword.substring(1);
    }

    if (sheetUserId === inputUserId && sheetPassword === inputPassword) {
      const user = buildUserObject_(data[i], i + 1);
      if (!user.approved) {
        throw new Error('บัญชีของคุณยังไม่ได้รับการอนุมัติจากผู้ดูแลระบบ');
      }
      return user;
    }
  }
  
  throw new Error('User ID หรือ รหัสผ่าน ไม่ถูกต้อง');
}

function getCurrentUserProfile(userId) {
  ensureAdminAccount_();
  const found = getUserRowByUserId_(userId);
  if (!found) throw new Error('ไม่พบข้อมูลผู้ใช้งาน');
  return found.user;
}

function updateUserProfile(userId, profileData) {
  ensureAdminAccount_();
  const found = getUserRowByUserId_(userId);
  if (!found) throw new Error('ไม่พบข้อมูลผู้ใช้งาน');

  const fullName = String(profileData.fullName || '').trim();
  const email = String(profileData.email || '').trim();
  const line = String(profileData.line || '').trim();
  const phone = String(profileData.phone || '').trim();
  const password = String(profileData.password || '').trim();
  const confirmPassword = String(profileData.confirmPassword || '').trim();

  if (!fullName) throw new Error('ชื่อ - นามสกุล ไม่สามารถเว้นว่างได้');
  if ((password || confirmPassword) && password !== confirmPassword) {
    throw new Error('Password และ Confirm Password ไม่ตรงกัน');
  }

  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.FULL_NAME).setValue(fullName);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.EMAIL).setValue(email);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.LINE).setValue(line);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.PHONE).setValue(phone);
  if (password) {
    found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.PASSWORD).setValue("'" + password);
  }

  SpreadsheetApp.flush();

  return getCurrentUserProfile(userId);
}

function addManagedUser(actingUserId, payload) {
  assertAdmin_(actingUserId);
  if (!payload) throw new Error('ไม่พบข้อมูลผู้ใช้งาน');

  const sheet = getUserSheet_();
  const data = sheet.getDataRange().getValues();
  const inputUserId = String(payload.userId || '').trim().toLowerCase();
  const fullName = String(payload.fullName || '').trim();
  const password = String(payload.password || '').trim();
  const confirmPassword = String(payload.confirmPassword == null ? payload.password : payload.confirmPassword).trim();
  const roleValue = normalizeAllowedValue_(payload.roleValue, USER_ROLE_OPTIONS, 'Role', false);
  const positionValue = normalizeAllowedValue_(payload.position || getPositionFromRole_(roleValue) || 'Staff', USER_POSITION_OPTIONS, 'Position', false);
  const approved = normalizeBoolean_(payload.approved);
  const teamValue = normalizeAllowedValue_(payload.team || '', getUserTeamOptions_(), 'Team', true);

  if (!inputUserId) throw new Error('กรุณากรอก User ID');
  if (!fullName) throw new Error('กรุณากรอกชื่อ - นามสกุล');
  if (!password) throw new Error('กรุณากรอก Password');
  if (!confirmPassword) throw new Error('กรุณากรอก Confirm Password');
  if (password !== confirmPassword) throw new Error('Password และ Confirm Password ไม่ตรงกัน');
  if (positionValue === 'Supplier' && !teamValue) throw new Error('กรุณาเลือกทีมเข้าแก้ไขสำหรับผู้ใช้ Supplier');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === inputUserId) {
      throw new Error('User ID นี้มีผู้ใช้งานแล้ว กรุณาใช้ชื่ออื่น');
    }
  }

  const newRow = new Array(15).fill('');
  newRow[0] = String(payload.userId).trim();
  newRow[1] = "'" + confirmPassword;
  newRow[2] = fullName;
  newRow[4] = roleValue;
  newRow[5] = positionValue;
  newRow[6] = approved;
  newRow[8] = String(payload.email || '').trim();
  newRow[9] = String(payload.line || '').trim();
  newRow[10] = String(payload.phone || '').trim();
  newRow[12] = teamValue;
  newRow[13] = new Date();

  sheet.appendRow(newRow);
  SpreadsheetApp.flush();
  return 'Success';
}

function getUsersForManagement(actingUserId) {
  assertAdmin_(actingUserId);
  const sheet = getUserSheet_();
  const data = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    users.push(buildUserObject_(data[i], i + 1));
  }

  return JSON.stringify(users);
}

function updateManagedUser(actingUserId, payload) {
  assertAdmin_(actingUserId);
  if (!payload || !payload.userId) throw new Error('ไม่พบ userId สำหรับการบันทึกข้อมูล');
  
  const found = getUserRowByUserId_(payload.userId);
  if (!found) throw new Error('ไม่พบข้อมูลผู้ใช้งาน');

  const fullName = String(payload.fullName == null ? found.user.fullName : payload.fullName).trim();
  const password = String(payload.password == null ? found.user.password : payload.password).trim();
  const confirmPassword = String(payload.confirmPassword == null ? payload.password : payload.confirmPassword).trim();
  const roleValue = normalizeAllowedValue_(payload.roleValue == null ? found.user.roleValue : payload.roleValue, USER_ROLE_OPTIONS, 'Role', false);
  const positionValue = normalizeAllowedValue_(payload.position == null ? found.user.position || getPositionFromRole_(roleValue) : payload.position, USER_POSITION_OPTIONS, 'Position', false);
  const approved = normalizeBoolean_(payload.approved == null ? found.user.approved : payload.approved);
  const teamValue = normalizeAllowedValue_(payload.team == null ? found.user.team : payload.team, getUserTeamOptions_(), 'ทีมเข้าแก้ไข', true);
  const email = String(payload.email == null ? found.user.email : payload.email).trim();
  const line = String(payload.line == null ? found.user.line : payload.line).trim();
  const phone = String(payload.phone == null ? found.user.phone : payload.phone).trim();

  if (!fullName) throw new Error('Full Name ไม่สามารถเว้นว่างได้');
  if (!password) throw new Error('Password ไม่สามารถเว้นว่างได้');
  if (password !== confirmPassword) {
    throw new Error('Password และ Confirm Password ไม่ตรงกัน');
  }
  if (positionValue === 'Supplier' && !teamValue) {
    throw new Error('กรุณาเลือกทีมเข้าแก้ไขสำหรับผู้ใช้ Supplier');
  }

  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.PASSWORD).setValue("'" + confirmPassword);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.FULL_NAME).setValue(fullName);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.ROLE).setValue(roleValue);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.POSITION).setValue(positionValue);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.APPROVED).setValue(approved);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.EMAIL).setValue(email);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.LINE).setValue(line);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.PHONE).setValue(phone);
  found.sheet.getRange(found.rowIndex, USER_SHEET_COLUMNS.TEAM).setValue(teamValue);

  SpreadsheetApp.flush(); // บังคับให้บันทึกลง Sheet ทันที เพื่อให้ตั้งค่า Approved มีผลกับการ Login ทันที

  return JSON.stringify(buildUserObject_(found.sheet.getRange(found.rowIndex, 1, 1, 15).getValues()[0], found.rowIndex));
}

function deleteManagedUser(actingUserId, targetUserId) {
  assertAdmin_(actingUserId);
  const found = getUserRowByUserId_(targetUserId);
  if (!found) throw new Error('ไม่พบข้อมูลผู้ใช้งาน');
  if (String(found.user.userId || '').trim().toLowerCase() === 'phukao') {
    throw new Error('ไม่สามารถลบบัญชีผู้ดูแลระบบหลักได้');
  }
  if (String(actingUserId || '').trim().toLowerCase() === String(targetUserId || '').trim().toLowerCase()) {
    throw new Error('ไม่สามารถลบบัญชีที่กำลังใช้งานอยู่ได้');
  }

  found.sheet.deleteRow(found.rowIndex);
  SpreadsheetApp.flush();
  return 'Success';
}

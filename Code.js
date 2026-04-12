// ============================================================
// MARKAZ AL ASAS — Code.gs  (Full Fixed Version)
// STEP 1: Replace YOUR_SPREADSHEET_ID_HERE with your Sheet ID
// STEP 2: Run fixAll() once
// STEP 3: Deploy as NEW VERSION
// ============================================================

var SS_ID = '1OL6ITLAuk3gpJ-8VnmRHvyPd5wYNBm4iXjlASjL7qEU';

/* ── Format a date value from spreadsheet into DD/MM/YYYY ── */
function fmtGasDate(val) {
  if (!val || val === '') return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    var dd   = String(d.getDate()).padStart(2, '0');
    var mm   = String(d.getMonth() + 1).padStart(2, '0');
    var yyyy = d.getFullYear();
    return dd + '/' + mm + '/' + yyyy;
  } catch(e) { return String(val); }
}

// ── WEB APP ──────────────────────────────────────────────────
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  var files = { admin:'admin', parent:'parent', teacher:'teacher', student:'student' };
  return HtmlService.createHtmlOutputFromFile(files[page] || 'index')
    .setTitle('Markaz Al Asas Academy')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── PING ─────────────────────────────────────────────────────
function ping() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheets = ss.getSheets().map(function(s){ return s.getName(); });
    return { ok:true, name:ss.getName(), sheets:sheets, count:sheets.length };
  } catch(e) {
    return { ok:false, error:e.message };
  }
}

// ── LOGIN ─────────────────────────────────────────────────────
function handleLoginFromClient(username, password, role) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success:false, message:'Users sheet missing. Run fixAll() first.' };
    var rows = sheet.getDataRange().getValues();
    var u = String(username).trim().toLowerCase();
    var p = String(password).trim();
    var r = String(role).trim().toLowerCase();
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row[0]) continue;
      if (String(row[1]).trim().toLowerCase() === u &&
          String(row[2]).trim() === p &&
          String(row[3]).trim().toLowerCase() === r) {
        return { success:true, name:String(row[4]), role:String(row[3]),
                 email:String(row[5]), token:Utilities.getUuid() };
      }
    }
    return { success:false, message:'Wrong username, password or role' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET STUDENTS ──────────────────────────────────────────────
function getStudents(classFilter) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return { success:false, message:'Students sheet not found. Run fixAll().' };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    Logger.log('Students sheet: lastRow=' + lastRow + ' lastCol=' + lastCol);

    if (lastRow < 1) return { success:true, data:[], total:0 };

    var rows = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 16)).getValues();
    Logger.log('Row 0: ' + JSON.stringify(rows[0]));
    if (rows.length > 1) Logger.log('Row 1: ' + JSON.stringify(rows[1]));

    var firstCell = String(rows[0][0]).trim();
    var start = (firstCell === 'StudentID') ? 1 : 0;
    Logger.log('firstCell=' + firstCell + ' start=' + start);

    var out = [];
    for (var i = start; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0] && !r[2]) continue;
      if (String(r[0]).trim() === 'StudentID') continue;

      var s = {
        StudentID:   String(r[0]  || ''),
        AdmissionNo: String(r[1]  || ''),
        Name:        String(r[2]  || ''),
        NameArabic:  String(r[3]  || ''),
        DOB:         fmtGasDate(r[4]),
        Gender:      String(r[5]  || ''),
        Class:       String(r[6]  || ''),
        Section:     String(r[7]  || ''),
        FatherName:  String(r[8]  || ''),
        MotherName:  String(r[9]  || ''),
        Phone:       String(r[10] || ''),
        Address:     String(r[11] || ''),
        Email:       String(r[12] || ''),
        Status:      String(r[15] || 'Active')
      };

      if (s.Status === 'Inactive') continue;
      if (classFilter && classFilter !== '' && classFilter !== 'All Classes' &&
          s.Class !== classFilter) continue;
      out.push(s);
    }

    Logger.log('getStudents returning ' + out.length + ' students');
    return { success:true, data:out, total:out.length };
  } catch(e) {
    Logger.log('getStudents ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── GET TEACHERS ──────────────────────────────────────────────
function getTeachers() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Teachers');
    if (!sheet) return { success:false, message:'Teachers sheet not found. Run fixAll().' };
    var rows = sheet.getDataRange().getValues();
    if (rows.length < 1) return { success:true, data:[] };
    var start = (String(rows[0][0]).trim() === 'TeacherID') ? 1 : 0;
    var out = [];
    for (var i = start; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      out.push({
        TeacherID:   String(rows[i][0] || ''),
        Name:        String(rows[i][1] || ''),
        Designation: String(rows[i][2] || ''),
        Subject:     String(rows[i][3] || ''),
        Phone:       String(rows[i][4] || ''),
        Email:       String(rows[i][5] || ''),
        Status:      String(rows[i][10] || 'Active')
      });
    }
    return { success:true, data:out };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET ADMISSIONS ────────────────────────────────────────────
function getAdmissions() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Admissions');

    if (!sheet) {
      Logger.log('getAdmissions: Admissions sheet NOT found');
      return { success:false, message:'Admissions sheet not found. Run fixAll() first.' };
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    Logger.log('getAdmissions: lastRow=' + lastRow + ' lastCol=' + lastCol);

    if (lastRow < 1) {
      Logger.log('getAdmissions: sheet is empty');
      return { success:true, data:[], total:0 };
    }

    // Read all rows with at least 14 columns
    var numCols = Math.max(lastCol, 17);
    var rows = sheet.getRange(1, 1, lastRow, numCols).getValues();

    Logger.log('getAdmissions Row 0: ' + JSON.stringify(rows[0]));
    if (rows.length > 1) Logger.log('getAdmissions Row 1: ' + JSON.stringify(rows[1]));

    // Detect and skip header row
    var firstCell = String(rows[0][0]).trim();
    var start = (firstCell === 'ApplicationID') ? 1 : 0;
    Logger.log('getAdmissions: firstCell=' + firstCell + ' start=' + start);

    var out = [];
    for (var i = start; i < rows.length; i++) {
      var r = rows[i];
      // Skip completely empty rows
      if (!r[0] && !r[1]) continue;
      // Skip stray header rows
      if (String(r[0]).trim() === 'ApplicationID') continue;

      // Auto-detect schema by column count:
      // OLD schema (<=14 cols): ...Address | PrevSchool | '' | SubmittedAt | Status
      //   0-9 same, 10=PrevSchool, 11='', 12='' , 13=SubmittedAt, 14=Status
      // NEW schema (17 cols):   ...Address | PrevMadrasa | PrevRegNo | AcYear | Photo | Docs | SubmittedAt | Status
      //   0-9 same, 10=PrevMadrasa, 11=PrevRegNo, 12=AcYear, 13=Photo, 14=Docs, 15=SubmittedAt, 16=Status
      var nc = r.length;
      var isNewSchema = (nc >= 17);
      var photoVal = '';
      var submittedVal = '';
      var statusVal   = 'Pending';
      var prevMadrasa = '';
      var prevRegNo   = '';
      var academicYear= '';

      if (isNewSchema) {
        prevMadrasa  = String(r[10] || '');
        prevRegNo    = String(r[11] || '');
        academicYear = String(r[12] || '');
        photoVal     = String(r[13] || '');
        submittedVal = fmtGasDate(r[15]);
        statusVal    = String(r[16] || 'Pending');
      } else {
        // Old schema: 10=PrevSchool, 11=empty, 12=empty, 13=SubmittedAt, 14=Status
        prevMadrasa  = String(r[10] || '');
        prevRegNo    = '';
        academicYear = '';
        photoVal     = '';
        submittedVal = fmtGasDate(r[13] || r[12]);
        statusVal    = String(nc > 14 ? r[14] : (nc > 13 ? r[13] : 'Pending')) || 'Pending';
        // If status looks like a date/timestamp, it's in wrong place
        if (statusVal && statusVal.match(/^\d{4}-|^\/|^\d{2}\//)) {
          statusVal = 'Pending';
        }
      }

      out.push({
        ApplicationID:    String(r[0]  || ''),
        StudentName:      String(r[1]  || ''),
        DOB:              fmtGasDate(r[2]),
        Gender:           String(r[3]  || ''),
        ApplyingForClass: String(r[4]  || ''),
        FatherName:       String(r[5]  || ''),
        MotherName:       String(r[6]  || ''),
        Phone:            String(r[7]  || ''),
        Email:            String(r[8]  || ''),
        Address:          String(r[9]  || ''),
        PreviousMadrasa:  prevMadrasa,
        PreviousSchool:   prevMadrasa,
        PrevRegNo:        prevRegNo,
        AcademicYear:     academicYear,
        Photo:            photoVal,
        SubmittedAt:      submittedVal,
        Status:           statusVal
      });
    }

    Logger.log('getAdmissions: returning ' + out.length + ' records');
    return { success:true, data:out, total:out.length };
  } catch(e) {
    Logger.log('getAdmissions ERROR: ' + e.message + ' | Stack: ' + e.stack);
    return { success:false, message:'getAdmissions error: ' + e.message };
  }
}

// ── GET DASHBOARD STATS ───────────────────────────────────────
function getDashboardStats() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
    var stuSheet = ss.getSheetByName('Students');
    var tchSheet = ss.getSheetByName('Teachers');
    var admSheet = ss.getSheetByName('Admissions');

    var stuCount = 0;
    if (stuSheet && stuSheet.getLastRow() > 1) {
      var stuRows = stuSheet.getRange(2, 16, stuSheet.getLastRow() - 1, 1).getValues();
      stuRows.forEach(function(r) {
        if (String(r[0]).trim() !== 'Inactive') stuCount++;
      });
    }

    var tchCount = tchSheet ? Math.max(0, tchSheet.getLastRow() - 1) : 0;

    var pendingAdm = 0;
    if (admSheet && admSheet.getLastRow() > 1) {
      var numCols = Math.max(admSheet.getLastColumn(), 14);
      admSheet.getRange(2, 1, admSheet.getLastRow() - 1, numCols).getValues().forEach(function(r) {
        if (!r[0]) return;
        var status = String(r[13] || 'Pending').trim();
        if (status === 'Pending') pendingAdm++;
      });
    }

    return { success:true, stats:{
      totalStudents:  stuCount,
      totalTeachers:  tchCount,
      attendanceRate: 0,
      pendingAdmissions: pendingAdm,
      todayDate: today
    }};
  } catch(e) {
    Logger.log('getDashboardStats ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── GET SCRIPT URL ────────────────────────────────────────────
function getScriptUrl() {
  try { return ScriptApp.getService().getUrl(); } catch(e) { return ''; }
}

// ── ADD STUDENT ───────────────────────────────────────────────
function addStudentDirect(data) {
  try {
    if (!data || !data.name || !data.name.trim())
      return { success:false, message:'Student name is required' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return { success:false, message:'Students sheet not found. Run fixAll().' };
    var id = 'STU' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMddHHmmss');
    sheet.appendRow([
      id, id, data.name, '', data.dob||'', data.gender||'',
      data.class||'', data.section||'', data.fatherName||'', data.motherName||'',
      data.phone||'', data.address||'', data.email||'', '',
      new Date().toISOString(), 'Active'
    ]);
    Logger.log('Student added: ' + id + ' - ' + data.name);
    return { success:true, studentId:id, message:'Student added: ' + data.name };
  } catch(e) {
    Logger.log('addStudentDirect ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── ADD TEACHER ───────────────────────────────────────────────
function addTeacherDirect(data) {
  try {
    if (!data || !data.name || !data.name.trim())
      return { success:false, message:'Teacher name is required' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Teachers');
    if (!sheet) return { success:false, message:'Teachers sheet not found. Run fixAll().' };
    var id = 'TCH' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMddHHmmss');
    sheet.appendRow([
      id, data.name, data.designation||'', data.subject||'',
      data.phone||'', data.email||'', data.joinDate||'', '', '',
      new Date().toISOString(), 'Active'
    ]);
    return { success:true, teacherId:id, message:'Teacher added: ' + data.name };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── MARK ATTENDANCE ───────────────────────────────────────────
function markAttendanceDirect(data) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Attendance');
    if (!sheet) return { success:false, message:'Attendance sheet not found' };
    var date = data.date || Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
    var records = data.records || [];
    records.forEach(function(r) {
      if (r.studentId) sheet.appendRow([
        date, r.studentId, r.status||'P', data.classId||'', '', new Date().toISOString()
      ]);
    });
    return { success:true, message:'Attendance saved: ' + records.length + ' students' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── ADD NEWS ──────────────────────────────────────────────────
function addNewsDirect(data) {
  try {
    if (!data || !data.title) return { success:false, message:'Title required' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('News');
    if (!sheet) return { success:false, message:'News sheet not found' };
    sheet.appendRow([
      data.title, data.content||'', data.category||'General', '',
      new Date().toISOString(), data.author||'Admin', 'Active'
    ]);
    return { success:true, message:'News published: ' + data.title };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── ADD EVENT ─────────────────────────────────────────────────
function addEventDirect(data) {
  try {
    if (!data || !data.title) return { success:false, message:'Title required' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Events');
    if (!sheet) return { success:false, message:'Events sheet not found' };
    sheet.appendRow([
      data.title, data.date||'', data.time||'', data.venue||'',
      data.description||'', data.category||'General', '', new Date().toISOString(), 'Active'
    ]);
    return { success:true, message:'Event added: ' + data.title };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── ADD FEE ───────────────────────────────────────────────────
function addFeeDirect(data) {
  try {
    if (!data || !data.studentId || !data.amount)
      return { success:false, message:'Student ID and amount required' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Fees');
    if (!sheet) return { success:false, message:'Fees sheet not found' };
    var id = 'RCP' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMddHHmmss');
    sheet.appendRow([
      id, data.studentId, data.studentName||'', data.class||'',
      data.feeType||'Monthly Fee', data.amount, data.month||'',
      data.academicYear||'2024-25', new Date().toISOString(), 'Admin', 'Paid'
    ]);
    return { success:true, receiptId:id, message:'Fee recorded: ' + id };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── ADD GALLERY ───────────────────────────────────────────────
function testGallery() {
  // Run this function directly in the Apps Script editor to test gallery
  var result = addGalleryDirect({
    title: 'Test Photo',
    imageUrl: 'https://via.placeholder.com/400x300.jpg?text=Gallery+Test',
    category: 'Test'
  });
  Logger.log('testGallery result: ' + JSON.stringify(result));
  var gallery = getGallery();
  Logger.log('Gallery items: ' + gallery.data.length);
}

// ── GET UPLOAD CONFIG ─────────────────────────────────────────
// Called by browser BEFORE uploading — returns folder ID for
// direct Drive REST upload (bypasses google.script.run size limit)
function getGalleryUploadConfig() {
  try {
    var folderName = 'Markaz Al Asas Gallery';
    var iter = DriveApp.getFoldersByName(folderName);
    var folder;
    if (iter.hasNext()) {
      folder = iter.next();
    } else {
      folder = DriveApp.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    var folderId = folder.getId();
    var token    = ScriptApp.getOAuthToken();
    Logger.log('getGalleryUploadConfig: folderId=' + folderId);
    return { success:true, folderId:folderId, token:token };
  } catch(e) {
    Logger.log('getGalleryUploadConfig ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── SAVE GALLERY URL ──────────────────────────────────────────
// Called AFTER browser uploads file to Drive directly
// Just saves the title + Drive URL to the sheet
function saveGalleryItem(data) {
  try {
    if (!data || !data.title) return { success:false, message:'Title required' };
    if (!data.url)            return { success:false, message:'URL required' };

    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Gallery');
    if (!sheet) {
      sheet = ss.insertSheet('Gallery');
      sheet.appendRow(['Title','Photo','Category','Description','Date','UploadedBy']);
    }

    // Make the file publicly viewable
    try {
      var fileId = data.fileId || '';
      if (fileId) {
        var file = DriveApp.getFileById(fileId);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    } catch(shareErr) {
      Logger.log('Share warning: ' + shareErr.message);
    }

    sheet.appendRow([
      data.title,
      data.url,
      data.category || 'Gallery',
      '',
      Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'),
      'Admin'
    ]);

    Logger.log('saveGalleryItem OK: ' + data.title + ' => ' + data.url);
    return { success:true, message:'Photo saved: ' + data.title };
  } catch(e) {
    Logger.log('saveGalleryItem ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}



function addGalleryDirect(data) {
  try {
    // ── Validate ──────────────────────────────────────────────
    if (!data) return { success:false, message:'No data received' };
    if (!data.title) return { success:false, message:'Title required' };

    var photo = data.photo || data.imageUrl || '';
    Logger.log('addGalleryDirect: title=' + data.title +
               ' | photoLen=' + photo.length +
               ' | isBase64=' + (photo.indexOf('data:image') === 0));

    // ── Get sheet ─────────────────────────────────────────────
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Gallery');
    if (!sheet) {
      // Auto-create if missing
      sheet = ss.insertSheet('Gallery');
      sheet.appendRow(['Title','Photo','Category','Description','Date','UploadedBy']);
      Logger.log('Gallery sheet created');
    }

    var savedUrl = '';

    // ── Case 1: base64 image → save to Google Drive ───────────
    if (photo && photo.indexOf('data:image') === 0) {
      var commaIdx = photo.indexOf(',');
      if (commaIdx < 0) return { success:false, message:'Invalid base64 format' };

      var meta     = photo.substring(5, commaIdx);          // e.g. "image/jpeg;base64"
      var mimeType = meta.split(';')[0];                    // e.g. "image/jpeg"
      var ext      = mimeType.split('/')[1] || 'jpg';
      var b64data  = photo.substring(commaIdx + 1);

      Logger.log('Decoding base64: mime=' + mimeType + ' b64len=' + b64data.length);

      var bytes = Utilities.base64Decode(b64data);
      var blob  = Utilities.newBlob(bytes, mimeType,
                    data.title.replace(/[^a-zA-Z0-9 ]/g,'_') + '.' + ext);

      // Get or create "Markaz Al Asas Gallery" folder
      var folderIter = DriveApp.getFoldersByName('Markaz Al Asas Gallery');
      var folder = folderIter.hasNext()
                   ? folderIter.next()
                   : DriveApp.createFolder('Markaz Al Asas Gallery');

      // Make folder public
      try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
      catch(shareErr) { Logger.log('Folder share warning: ' + shareErr.message); }

      // Save file
      var file   = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
      catch(shareErr) { Logger.log('File share warning: ' + shareErr.message); }

      savedUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
      Logger.log('Saved to Drive: ' + savedUrl);

    // ── Case 2: plain URL ─────────────────────────────────────
    } else if (photo && photo.length > 10) {
      savedUrl = photo;
      Logger.log('Using URL directly: ' + savedUrl.substring(0,60));

    } else {
      Logger.log('No photo provided, saving title only');
    }

    // ── Write row ─────────────────────────────────────────────
    sheet.appendRow([
      data.title,
      savedUrl,
      data.category || 'Gallery',
      data.description || '',
      Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'),
      'Admin'
    ]);

    Logger.log('Gallery row saved OK: ' + data.title);
    return { success:true, message:'Photo saved: ' + data.title, url: savedUrl };

  } catch(e) {
    Logger.log('addGalleryDirect EXCEPTION: ' + e.message + ' | stack: ' + e.stack);
    return { success:false, message:e.message };
  }
}

// ── ADD MARKS ─────────────────────────────────────────────────
function addMarksDirect(data) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Marks');
    if (!sheet) return { success:false, message:'Marks sheet not found' };
    var records = data.records || [];
    if (records.length === 0) return { success:false, message:'No records provided' };
    records.forEach(function(r) {
      var pct = r.maxMarks ? Math.round((r.marksObtained / r.maxMarks) * 100) : 0;
      var grade = pct>=90?'A+':pct>=80?'A':pct>=70?'B+':pct>=60?'B':pct>=50?'C':pct>=35?'D':'F';
      sheet.appendRow([
        r.studentId, data.examName||'', data.academicYear||'2024-25',
        data.class||'', r.subject||'', r.maxMarks||100, r.marksObtained||0,
        grade, '', new Date().toISOString()
      ]);
    });
    return { success:true, message:'Marks saved' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── APPROVE ADMISSION ─────────────────────────────────────────
function approveAdmissionDirect(data) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Admissions');
    if (!sheet) return { success:false, message:'Admissions sheet not found' };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success:false, message:'No admissions found' };
    var numCols = Math.max(sheet.getLastColumn(), 14);
    var rows = sheet.getRange(1, 1, lastRow, numCols).getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.applicationId).trim()) {
        sheet.getRange(i + 1, 14).setValue(data.status || 'Approved');
        Logger.log('Admission ' + data.applicationId + ' -> ' + data.status);
        return { success:true, message:'Admission ' + (data.status || 'Approved') };
      }
    }
    return { success:false, message:'Application not found: ' + data.applicationId };
  } catch(e) {
    Logger.log('approveAdmissionDirect ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── ADD COMMITTEE MEMBER ──────────────────────────────────────
function addCommitteeDirect(data) {
  try {
    if (!data || !data.name) return { success:false, message:'Name required' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Committee');
    if (!sheet) return { success:false, message:'Committee sheet not found' };
    sheet.appendRow([
      data.name, data.role||'', data.department||'',
      data.phone||'', data.email||'', data.photo||'', '', '',
      new Date().toISOString(), 'Active'
    ]);
    return { success:true, message:'Member added: ' + data.name };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── SUBMIT ADMISSION (public form) ────────────────────────────
function submitAdmissionDirect(data) {
  try {
    if (!data || !data.studentName || !data.phone)
      return { success:false, message:'Name and phone required' };

    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Admissions');
    if (!sheet) return { success:false, message:'Admissions sheet not found. Run fixAll().' };

    var id = 'APP' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMddHHmmss');

    // Handle photo — save to Drive if base64, store URL
    var photoUrl = '';
    if (data.photo && data.photo.indexOf('data:image') === 0) {
      try {
        var comma    = data.photo.indexOf(',');
        var mime     = data.photo.substring(5, comma).split(';')[0];
        var ext      = mime.split('/')[1] || 'jpg';
        var bytes    = Utilities.base64Decode(data.photo.substring(comma + 1));
        var blob     = Utilities.newBlob(bytes, mime,
                         id + '_' + (data.studentName||'student').replace(/[^a-zA-Z0-9]/g,'_') + '.' + ext);

        var folderIter = DriveApp.getFoldersByName('Markaz Al Asas Admissions');
        var folder = folderIter.hasNext()
                     ? folderIter.next()
                     : DriveApp.createFolder('Markaz Al Asas Admissions');
        try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}

        var file = folder.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        photoUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
        Logger.log('Admission photo saved: ' + photoUrl);
      } catch(photoErr) {
        Logger.log('Admission photo upload warning: ' + photoErr.message);
        photoUrl = '';  // continue without photo
      }
    }

    // Schema cols:
    // ApplicationID | StudentName | DOB | Gender | ApplyingForClass |
    // FatherName | MotherName | Phone | Email | Address |
    // PreviousMadrasa | PrevRegNo | AcademicYear | Photo | Documents | SubmittedAt | Status
    sheet.appendRow([
      id,
      data.studentName,
      data.dob             || '',
      data.gender          || '',
      data.applyingForClass|| '',
      data.fatherName      || '',
      data.motherName      || '',
      data.phone,
      data.email           || '',
      data.address         || '',
      data.previousSchool  || '',   // col 11: Previous Madrasa name
      data.prevRegNo       || '',   // col 12: Reg No (numeric)
      data.academicYear    || '2024-25', // col 13
      photoUrl,                     // col 14: Drive URL
      '',                           // col 15: Documents
      new Date().toISOString(),     // col 16: Submitted at
      'Pending'                     // col 17: Status
    ]);

    Logger.log('Admission submitted: ' + id + ' for ' + data.studentName);
    return { success:true, applicationId:id, message:'Application submitted successfully! ID: ' + id };
  } catch(e) {
    Logger.log('submitAdmissionDirect ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── HANDLE POST FROM CLIENT (index.html public form) ─────────
function handlePostFromClient(payload) {
  try {
    if (!payload) return { success:false, message:'No data received' };
    var action = payload.action || '';
    if (action === 'submitAdmission') {
      return submitAdmissionDirect(payload);
    }
    return { success:false, message:'Unknown action: ' + action };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── SAVE SETTINGS ─────────────────────────────────────────────
function saveSettingsDirect(data) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { success:false, message:'Settings sheet not found' };
    var settings = data.settings || {};
    var rows = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    Object.keys(settings).forEach(function(key) {
      var found = false;
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(settings[key]);
          sheet.getRange(i + 1, 3).setValue(now);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([key, settings[key], now]);
    });
    return { success:true, message:'Settings saved' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET STUDENT BY ID ─────────────────────────────────────────
function getStudentById(studentId) {
  try {
    Logger.log('getStudentById: searching for=' + studentId);
    if (!studentId || String(studentId).trim() === '') {
      return { success:false, message:'No student ID provided' };
    }
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return { success:false, message:'Students sheet not found' };
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return { success:false, message:'Students sheet is empty' };
    var rows = sheet.getRange(1, 1, lastRow, 16).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[0]).trim() === 'StudentID') continue;
      var id0 = String(r[0] || '').trim();
      var id1 = String(r[1] || '').trim();
      var search = String(studentId).trim();
      if (id0 === search || id1 === search) {
        Logger.log('getStudentById FOUND at row ' + i);
        return { success:true, data:{
          StudentID:   String(r[0]  || ''),
          AdmissionNo: String(r[1]  || ''),
          Name:        String(r[2]  || ''),
          NameArabic:  String(r[3]  || ''),
          DOB:         fmtGasDate(r[4]),
          Gender:      String(r[5]  || ''),
          Class:       String(r[6]  || ''),
          Section:     String(r[7]  || ''),
          FatherName:  String(r[8]  || ''),
          MotherName:  String(r[9]  || ''),
          Phone:       String(r[10] || ''),
          Address:     String(r[11] || ''),
          Email:       String(r[12] || ''),
          Status:      String(r[15] || 'Active')
        }};
      }
    }
    Logger.log('getStudentById NOT FOUND: ' + studentId);
    return { success:false, message:'Student not found: ' + studentId };
  } catch(e) {
    Logger.log('getStudentById ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── UPDATE STUDENT ────────────────────────────────────────────
function updateStudentDirect(data) {
  try {
    if (!data || !data.studentId) return { success:false, message:'Student ID required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return { success:false, message:'Students sheet not found' };
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return { success:false, message:'No students found' };
    var rows = sheet.getRange(1, 1, lastRow, 16).getValues();
    var sid  = String(data.studentId).trim();
    for (var i = 0; i < rows.length; i++) {
      var id0 = String(rows[i][0] || '').trim();
      var id1 = String(rows[i][1] || '').trim();
      if (id0 === sid || id1 === sid) {
        var r = i + 1;
        if (data.name)        sheet.getRange(r, 3).setValue(data.name);
        if (data.dob)         sheet.getRange(r, 5).setValue(data.dob);
        if (data.gender)      sheet.getRange(r, 6).setValue(data.gender);
        if (data.class)       sheet.getRange(r, 7).setValue(data.class);
        if (data.section      !== undefined) sheet.getRange(r, 8).setValue(data.section);
        if (data.fatherName   !== undefined) sheet.getRange(r, 9).setValue(data.fatherName);
        if (data.motherName   !== undefined) sheet.getRange(r, 10).setValue(data.motherName);
        if (data.phone        !== undefined) sheet.getRange(r, 11).setValue(data.phone);
        if (data.address      !== undefined) sheet.getRange(r, 12).setValue(data.address);
        if (data.email        !== undefined) sheet.getRange(r, 13).setValue(data.email);
        if (data.status)      sheet.getRange(r, 16).setValue(data.status);
        Logger.log('Student updated: ' + sid + ' row=' + r);
        return { success:true, message:'Student updated: ' + (data.name || sid) };
      }
    }
    return { success:false, message:'Student not found: ' + sid };
  } catch(e) {
    Logger.log('updateStudentDirect ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── DELETE STUDENT (sets Status=Inactive) ─────────────────────
function deleteStudentDirect(data) {
  try {
    if (!data || !data.studentId) return { success:false, message:'Student ID required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return { success:false, message:'Students sheet not found' };
    var lastRow = sheet.getLastRow();
    var rows = sheet.getRange(1, 1, lastRow, 16).getValues();
    var sid  = String(data.studentId).trim();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === sid || String(rows[i][1]).trim() === sid) {
        sheet.getRange(i + 1, 16).setValue('Inactive');
        Logger.log('Student set Inactive: ' + sid);
        return { success:true, message:'Student removed' };
      }
    }
    return { success:false, message:'Student not found: ' + sid };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET FEE STATUS ────────────────────────────────────────────
function getFeeStatus(studentId) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Fees');
    if (!sheet || sheet.getLastRow() < 2) return { success:true, data:[] };
    var rows  = sheet.getRange(1, 1, sheet.getLastRow(), 11).getValues();
    var start = String(rows[0][0]).trim() === 'ReceiptID' ? 1 : 0;
    var sid   = String(studentId || '').trim();
    var out   = [];
    for (var i = start; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (sid && String(rows[i][1]).trim() !== sid) continue;
      out.push({
        ReceiptID:    String(rows[i][0]  || ''),
        StudentID:    String(rows[i][1]  || ''),
        StudentName:  String(rows[i][2]  || ''),
        Class:        String(rows[i][3]  || ''),
        FeeType:      String(rows[i][4]  || ''),
        Amount:       rows[i][5] || 0,
        Month:        String(rows[i][6]  || ''),
        AcademicYear: String(rows[i][7]  || ''),
        Status:       String(rows[i][10] || 'Paid')
      });
    }
    return { success:true, data:out };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET ATTENDANCE SUMMARY ────────────────────────────────────
function getAttendanceSummary(data) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Attendance');
    if (!sheet || sheet.getLastRow() < 2) return { success:true, data:{} };
    var rows  = sheet.getRange(1, 1, sheet.getLastRow(), 6).getValues();
    var start = String(rows[0][0]).trim() === 'Date' ? 1 : 0;
    var sid   = data && data.studentId ? String(data.studentId).trim() : '';
    var summary = {};
    for (var i = start; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      var rsid   = String(rows[i][1] || '').trim();
      var status = String(rows[i][2] || 'P').trim();
      if (sid && rsid !== sid) continue;
      if (!summary[rsid]) summary[rsid] = { P:0, A:0, L:0, total:0, percentage:0 };
      summary[rsid][status] = (summary[rsid][status] || 0) + 1;
      summary[rsid].total++;
    }
    Object.keys(summary).forEach(function(id) {
      var s = summary[id];
      s.percentage = s.total > 0 ? Math.round((s.P / s.total) * 100) : 0;
    });
    return { success:true, data:summary };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET ATTENDANCE FOR STUDENT ────────────────────────────────
function getAttendanceForStudent(studentId) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Attendance');
    if (!sheet || sheet.getLastRow() < 2) return { success:true, data:[] };
    var rows  = sheet.getRange(1, 1, sheet.getLastRow(), 6).getValues();
    var start = String(rows[0][0]).trim() === 'Date' ? 1 : 0;
    var sid   = String(studentId || '').trim();
    var out   = [];
    for (var i = start; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (sid && String(rows[i][1] || '').trim() !== sid) continue;
      out.push({
        date:      fmtGasDate(rows[i][0]),
        studentId: String(rows[i][1] || ''),
        status:    String(rows[i][2] || 'P'),
        classId:   String(rows[i][3] || '')
      });
    }
    return { success:true, data:out };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET MARKS FOR CLASS ───────────────────────────────────────
function getMarksForClass(data) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Marks');
    if (!sheet || sheet.getLastRow() < 2) return { success:true, data:[] };
    var rows  = sheet.getRange(1, 1, sheet.getLastRow(), 10).getValues();
    var start = String(rows[0][0]).trim() === 'StudentID' ? 1 : 0;
    var stuSheet = ss.getSheetByName('Students');
    var stuMap = {};
    if (stuSheet) {
      stuSheet.getDataRange().getValues().slice(1).forEach(function(r) {
        if (r[0]) stuMap[String(r[0]).trim()] = { name:String(r[2]||''), class:String(r[6]||''), admNo:String(r[1]||'') };
      });
    }
    var out = [];
    for (var i = start; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var sid  = String(r[0]).trim();
      var exam = String(r[1] || '').trim();
      var yr   = String(r[2] || '').trim();
      var cls  = String(r[3] || '').trim();
      if (data.examName     && exam !== data.examName)     continue;
      if (data.academicYear && yr   !== data.academicYear) continue;
      if (data.class && data.class !== '' && cls !== data.class) continue;
      var stu = stuMap[sid] || {};
      out.push({
        StudentID:     sid,
        AdmissionNo:   stu.admNo || sid,
        Name:          stu.name  || sid,
        Class:         cls || stu.class || '',
        Subject:       String(r[4] || ''),
        MaxMarks:      r[5] || 100,
        MarksObtained: r[6] || 0,
        Grade:         String(r[7] || ''),
        ExamName:      exam,
        AcademicYear:  yr
      });
    }
    return { success:true, data:out };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── LINK PARENT TO STUDENT ────────────────────────────────────
function linkParentStudent(data) {
  try {
    if (!data || !data.parent || !data.studentId)
      return { success:false, message:'Parent and student ID required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { success:false, message:'Settings sheet not found' };
    var key = 'parent_link_' + String(data.parent).replace(/[@.]/g,'_') + '_' + String(data.studentId);
    var val = JSON.stringify({
      parent:       data.parent,
      studentId:    data.studentId,
      relationship: data.relationship || 'Parent',
      createdAt:    new Date().toISOString()
    });
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) {
        sheet.getRange(i+1, 2).setValue(val);
        sheet.getRange(i+1, 3).setValue(new Date().toISOString());
        return { success:true, message:'Link updated' };
      }
    }
    sheet.appendRow([key, val, new Date().toISOString()]);
    Logger.log('Parent linked: ' + data.parent + ' -> ' + data.studentId);
    return { success:true, message:'Parent linked to student!' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET PARENT LINKS ──────────────────────────────────────────
function getParentLinks() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { success:true, data:[] };
    var rows = sheet.getDataRange().getValues();
    var out  = [];
    var stuSheet = ss.getSheetByName('Students');
    var stuMap = {};
    if (stuSheet) {
      stuSheet.getDataRange().getValues().slice(1).forEach(function(r) {
        if (r[0]) stuMap[String(r[0]).trim()] = { name:String(r[2]||''), class:String(r[6]||'') };
        if (r[1]) stuMap[String(r[1]).trim()] = { name:String(r[2]||''), class:String(r[6]||'') };
      });
    }
    rows.slice(1).forEach(function(r) {
      if (String(r[0]).indexOf('parent_link_') === 0) {
        try {
          var d   = JSON.parse(String(r[1]));
          var stu = stuMap[d.studentId] || {};
          out.push({
            ParentUsername: d.parent,
            StudentID:      d.studentId,
            StudentName:    stu.name || d.studentId,
            Class:          stu.class || '',
            Relationship:   d.relationship || 'Parent'
          });
        } catch(pe) {}
      }
    });
    return { success:true, data:out };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET LINKED CHILDREN (by parent username) ──────────────────
function getLinkedChildren(parentUsername) {
  try {
    if (!parentUsername) return { success:true, data:[] };
    var ss        = SpreadsheetApp.openById(SS_ID);
    var settSheet = ss.getSheetByName('Settings');
    if (!settSheet) return { success:true, data:[] };

    var uname = String(parentUsername).trim().toLowerCase();
    var rows  = settSheet.getDataRange().getValues();
    var studentIds = [];

    rows.slice(1).forEach(function(r) {
      var key = String(r[0] || '');
      if (key.indexOf('parent_link_') !== 0) return;
      try {
        var val = JSON.parse(String(r[1] || '{}'));
        var linkParent = String(val.parent || '').trim().toLowerCase();
        if (linkParent === uname) {
          studentIds.push(String(val.studentId || '').trim());
        }
      } catch(pe) {}
    });

    if (studentIds.length === 0) return { success:true, data:[] };

    var stuSheet = ss.getSheetByName('Students');
    if (!stuSheet) return { success:true, data:[] };
    var stuRows  = stuSheet.getDataRange().getValues();
    var stuStart = String(stuRows[0][0]).trim() === 'StudentID' ? 1 : 0;
    var children = [];

    studentIds.forEach(function(sid) {
      for (var i = stuStart; i < stuRows.length; i++) {
        var r = stuRows[i];
        if (!r[0]) continue;
        if (String(r[0]).trim() === sid || String(r[1]).trim() === sid) {
          children.push({
            StudentID:   String(r[0]  || ''),
            AdmissionNo: String(r[1]  || ''),
            Name:        String(r[2]  || ''),
            NameArabic:  String(r[3]  || ''),
            DOB:         fmtGasDate(r[4]),
            Gender:      String(r[5]  || ''),
            Class:       String(r[6]  || ''),
            Section:     String(r[7]  || ''),
            FatherName:  String(r[8]  || ''),
            MotherName:  String(r[9]  || ''),
            Phone:       String(r[10] || ''),
            Address:     String(r[11] || ''),
            Email:       String(r[12] || ''),
            Status:      String(r[15] || 'Active')
          });
          break;
        }
      }
    });

    Logger.log('getLinkedChildren: ' + parentUsername + ' -> ' + children.length + ' children');
    return { success:true, data:children };
  } catch(e) {
    Logger.log('getLinkedChildren ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── GET CHILD DATA by phone/email ─────────────────────────────
function getChildData(data) {
  try {
    if (!data) return { success:true, data:[] };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return { success:true, data:[] };
    var rows     = sheet.getDataRange().getValues();
    var stuStart = String(rows[0][0]).trim() === 'StudentID' ? 1 : 0;
    var children = [];
    var phone    = String(data.phone || '').trim();
    var email    = String(data.email || '').trim().toLowerCase();
    for (var i = stuStart; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rPhone  = String(r[10] || '').replace(/\D/g, '').trim();
      var rEmail  = String(r[12] || '').trim().toLowerCase();
      var pMatch  = phone && (rPhone === phone.replace(/\D/g, '') || rPhone.endsWith(phone.slice(-10)));
      var eMatch  = email && rEmail === email;
      if (pMatch || eMatch) {
        children.push({
          StudentID:   String(r[0]  || ''),
          AdmissionNo: String(r[1]  || ''),
          Name:        String(r[2]  || ''),
          Class:       String(r[6]  || ''),
          Section:     String(r[7]  || ''),
          FatherName:  String(r[8]  || ''),
          MotherName:  String(r[9]  || ''),
          Phone:       String(r[10] || ''),
          Email:       String(r[12] || ''),
          DOB:         fmtGasDate(r[4]),
          Status:      String(r[15] || 'Active')
        });
      }
    }
    return { success:true, data:children };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── CREATE STUDENT LOGIN ──────────────────────────────────────
function createStudentLogin(data) {
  try {
    if (!data || !data.username || !data.password)
      return { success:false, message:'Username and password required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success:false, message:'Users sheet not found' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === data.username.trim().toLowerCase()) {
        return { success:false, message:'Username "' + data.username + '" already exists.' };
      }
    }
    var userId = 'STU_' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMddHHmmss');
    sheet.appendRow([userId, data.username.trim(), data.password.trim(), 'student',
      data.name || data.username, data.email || '', true]);
    var settings = ss.getSheetByName('Settings');
    if (settings && data.studentId) {
      settings.appendRow(['student_login_' + data.studentId, data.username, new Date().toISOString()]);
    }
    Logger.log('Student login created: ' + data.username);
    return { success:true, message:'Login created: ' + data.username };
  } catch(e) {
    Logger.log('createStudentLogin ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── BULK CREATE STUDENT LOGINS ────────────────────────────────
function bulkCreateStudentLogins(data) {
  try {
    var ss       = SpreadsheetApp.openById(SS_ID);
    var stuSheet = ss.getSheetByName('Students');
    var usrSheet = ss.getSheetByName('Users');
    if (!stuSheet || !usrSheet) return { success:false, message:'Required sheets not found' };
    var stuRows = stuSheet.getDataRange().getValues();
    var usrRows = usrSheet.getDataRange().getValues();
    var existing = {};
    usrRows.slice(1).forEach(function(r){ if(r[1]) existing[String(r[1]).toLowerCase()] = true; });
    var stuStart = String(stuRows[0][0]).trim() === 'StudentID' ? 1 : 0;
    var created = 0, skipped = 0;
    var cls = data && data.class ? data.class : '';
    stuRows.slice(stuStart).forEach(function(r) {
      if (!r[0]) return;
      if (cls && String(r[6]).trim() !== cls) return;
      if (String(r[15] || '').trim() === 'Inactive') return;
      var admno = String(r[1] || r[0] || '').trim();
      var phone = String(r[10] || '').replace(/\D/g,'').trim();
      var uname = admno.toLowerCase().replace(/[^a-z0-9_]/g,'');
      if (!uname || !phone) { skipped++; return; }
      if (existing[uname]) { skipped++; return; }
      var uid = 'STU_' + admno;
      usrSheet.appendRow([uid, uname, phone.slice(-10)||'123456', 'student', String(r[2]||''), String(r[12]||''), true]);
      existing[uname] = true;
      created++;
    });
    Logger.log('Bulk login: created=' + created + ' skipped=' + skipped);
    return { success:true, count:created, skipped:skipped, message:'Created ' + created + ' logins, ' + skipped + ' skipped' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── GET ALL USERS ─────────────────────────────────────────────
function getAllUsers() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success:true, data:[] };
    var rows  = sheet.getDataRange().getValues();
    var start = String(rows[0][0]).trim() === 'UserID' ? 1 : 0;
    var out   = rows.slice(start).filter(function(r){ return r[0]; }).map(function(r) {
      return {
        UserID:   String(r[0] || ''),
        Username: String(r[1] || ''),
        Role:     String(r[3] || ''),
        Name:     String(r[4] || ''),
        Email:    String(r[5] || ''),
        Active:   r[6] === true || String(r[6]).toLowerCase() === 'true'
      };
    });
    return { success:true, data:out };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── RESET USER PASSWORD ───────────────────────────────────────
function resetUserPassword(data) {
  try {
    if (!data || !data.username || !data.newPassword)
      return { success:false, message:'Username and new password required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success:false, message:'Users sheet not found' };
    var rows  = sheet.getDataRange().getValues();
    var uname = String(data.username).trim().toLowerCase();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === uname) {
        sheet.getRange(i + 1, 3).setValue(data.newPassword.trim());
        return { success:true, message:'Password reset for: ' + data.username };
      }
    }
    return { success:false, message:'User not found: ' + data.username };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── TOGGLE USER ACTIVE ────────────────────────────────────────
function toggleUserActive(data) {
  try {
    if (!data || !data.username) return { success:false, message:'Username required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success:false, message:'Users sheet not found' };
    var rows  = sheet.getDataRange().getValues();
    var uname = String(data.username).trim().toLowerCase();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === uname) {
        sheet.getRange(i + 1, 7).setValue(data.active === true);
        return { success:true, message:(data.active ? 'Activated' : 'Deactivated') + ': ' + data.username };
      }
    }
    return { success:false, message:'User not found' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── TEST FUNCTIONS — Run from Apps Script Editor to debug ─────

function testGetStudents() {
  Logger.log('=== testGetStudents ===');
  var res = getStudents('');
  Logger.log('success=' + res.success + ' count=' + (res.data ? res.data.length : 'N/A'));
  if (res.data && res.data.length > 0) Logger.log('First: ' + JSON.stringify(res.data[0]));
  if (!res.success) Logger.log('ERROR: ' + res.message);
}

function testGetAdmissions() {
  Logger.log('=== testGetAdmissions ===');
  var ss    = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('Admissions');
  if (!sheet) { Logger.log('ERROR: No Admissions sheet'); return; }
  Logger.log('lastRow=' + sheet.getLastRow() + ' lastCol=' + sheet.getLastColumn());
  if (sheet.getLastRow() > 0) {
    var all = sheet.getDataRange().getValues();
    all.forEach(function(row, i) {
      Logger.log('Row ' + (i+1) + ': ' + JSON.stringify(row.map(function(v){ return String(v).substring(0,20); })));
    });
  }
  var res = getAdmissions();
  Logger.log('Result: success=' + res.success + ' count=' + (res.data ? res.data.length : 'N/A'));
  if (res.data && res.data.length > 0) Logger.log('First: ' + JSON.stringify(res.data[0]));
  if (!res.success) Logger.log('ERROR: ' + res.message);
}

function testDashboard() {
  Logger.log('=== testDashboard ===');
  var res = getDashboardStats();
  Logger.log('Result: ' + JSON.stringify(res));
}


// ══════════════════════════════════════════════════════════════
// COMMITTEE & PUBLIC DASHBOARD FUNCTIONS
// ══════════════════════════════════════════════════════════════

// ── GET COMMITTEE ─────────────────────────────────────────────
function getCommittee() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Committee');
    if (!sheet || sheet.getLastRow() < 1) return { success: true, data: [] };

    var numCols = Math.max(sheet.getLastColumn(), 10);
    var rows    = sheet.getRange(1, 1, sheet.getLastRow(), numCols).getValues();

    // Skip header row if present
    var start = (String(rows[0][0]).trim() === 'Name') ? 1 : 0;
    var out   = [];

    for (var i = start; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;                                    // skip empty rows
      if (String(r[9] || 'Active').trim() === 'Inactive') continue; // skip inactive

      out.push({
        Name:       String(r[0] || ''),
        Role:       String(r[1] || ''),
        Department: String(r[2] || ''),
        Phone:      String(r[3] || ''),
        Email:      String(r[4] || ''),
        Photo:      String(r[5] || ''),  // base64 image stored here
        JoinDate:   String(r[6] || ''),
        Bio:        String(r[7] || ''),
        Status:     String(r[9] || 'Active')
      });
    }

    Logger.log('getCommittee: returned ' + out.length + ' members');
    return { success: true, data: out };
  } catch(e) {
    Logger.log('getCommittee ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── DELETE COMMITTEE MEMBER ───────────────────────────────────
function deleteCommitteeMember(data) {
  try {
    if (!data || !data.name) return { success: false, message: 'Name required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Committee');
    if (!sheet) return { success: false, message: 'Committee sheet not found' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.name).trim()) {
        sheet.getRange(i + 1, 10).setValue('Inactive');
        Logger.log('Committee member removed: ' + data.name);
        return { success: true, message: 'Member removed: ' + data.name };
      }
    }
    return { success: false, message: 'Member not found: ' + data.name };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── GET PUBLIC DASHBOARD SETTINGS ────────────────────────────
function getPublicDashboardSettings() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Settings');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: {} };

    var settings = {};
    sheet.getDataRange().getValues().slice(1).forEach(function(r) {
      var key = String(r[0] || '');
      if (key.indexOf('public_') !== 0) return;
      try { settings[key] = JSON.parse(String(r[1] || '{}')); }
      catch(e) { settings[key] = String(r[1] || ''); }
    });
    return { success: true, data: settings };
  } catch(e) {
    Logger.log('getPublicDashboardSettings ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── SAVE PUBLIC DASHBOARD SETTINGS ───────────────────────────
function savePublicDashboardSettings(data) {
  try {
    var ss      = SpreadsheetApp.openById(SS_ID);
    var sheet   = ss.getSheetByName('Settings');
    if (!sheet) return { success: false, message: 'Settings sheet not found. Run fixAll().' };

    var updates = data.settings || {};
    var now     = new Date().toISOString();

    // ── Handle logo: if base64, save to Drive first ───────────
    if (updates.public_identity && updates.public_identity.logo &&
        updates.public_identity.logo.indexOf('data:image') === 0) {
      try {
        var b64str   = updates.public_identity.logo;
        var comma    = b64str.indexOf(',');
        var mime     = b64str.substring(5, comma).split(';')[0];
        var ext      = mime.split('/')[1] || 'png';
        var bytes    = Utilities.base64Decode(b64str.substring(comma + 1));
        var blob     = Utilities.newBlob(bytes, mime, 'madrasa_logo.' + ext);

        var folderIter = DriveApp.getFoldersByName('Markaz Al Asas Gallery');
        var folder = folderIter.hasNext()
          ? folderIter.next()
          : DriveApp.createFolder('Markaz Al Asas Gallery');
        try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}

        var file = folder.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}

        // Replace base64 with Drive URL
        var driveUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
        updates.public_identity.logo = driveUrl;
        Logger.log('Logo saved to Drive: ' + driveUrl);
      } catch(logoErr) {
        Logger.log('Logo Drive upload warning: ' + logoErr.message);
        // Remove logo rather than fail entirely — don't let it block saving other settings
        updates.public_identity.logo = '';
      }
    }

    // ── Write all settings to sheet ───────────────────────────
    var rows = sheet.getDataRange().getValues();

    Object.keys(updates).forEach(function(key) {
      var val = updates[key];

      // Serialize objects to JSON
      var v = (typeof val === 'object') ? JSON.stringify(val) : String(val || '');

      // Safety: truncate if still too long (should never happen after logo fix)
      if (v.length > 49000) {
        Logger.log('WARNING: value for ' + key + ' is ' + v.length + ' chars, truncating');
        v = v.substring(0, 49000);
      }

      var found = false;
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === key) {
          sheet.getRange(i + 1, 2).setValue(v);
          sheet.getRange(i + 1, 3).setValue(now);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([key, v, now]);
    });

    Logger.log('savePublicDashboardSettings OK: ' + Object.keys(updates).join(', '));
    return {
      success: true,
      message: 'Settings saved',
      logoUrl: (updates.public_identity && updates.public_identity.logo) || ''
    };
  } catch(e) {
    Logger.log('savePublicDashboardSettings ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── GET ALL PUBLIC DATA (single call from index.html) ─────────
function getPublicData() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);

    // 1. Settings (ticker, stats, events, contact, adm classes)
    var settSheet = ss.getSheetByName('Settings');
    var settings  = {};
    if (settSheet && settSheet.getLastRow() > 1) {
      settSheet.getDataRange().getValues().slice(1).forEach(function(r) {
        var key = String(r[0] || '');
        if (key.indexOf('public_') !== 0) return;
        try { settings[key] = JSON.parse(String(r[1] || '{}')); }
        catch(e) { settings[key] = String(r[1] || ''); }
      });
    }

    // 2. Committee members (with photos)
    var cmRes   = getCommittee();
    var committee = cmRes.data || [];

    // 3. Gallery items
    var galSheet = ss.getSheetByName('Gallery');
    var gallery  = [];
    if (galSheet && galSheet.getLastRow() > 1) {
      galSheet.getDataRange().getValues().slice(1).forEach(function(r) {
        if (!r[0]) return;
        gallery.push({
          Title:    String(r[0] || ''),
          Photo:    String(r[1] || ''),   // Drive URL
          ImageURL: String(r[1] || ''),   // alias for backward compat
          Category: String(r[2] || ''),
          Date:     String(r[4] || '')
        });
      });
    }

    // 4. News items
    var newsSheet = ss.getSheetByName('News');
    var news      = [];
    if (newsSheet && newsSheet.getLastRow() > 1) {
      newsSheet.getDataRange().getValues().slice(1).forEach(function(r) {
        if (!r[0]) return;
        if (String(r[6] || 'Active').trim() !== 'Active') return;
        news.push({
          Title:    String(r[0] || ''),
          Content:  String(r[1] || ''),
          Category: String(r[2] || ''),
          Date:     String(r[4] || ''),
          Author:   String(r[5] || '')
        });
      });
    }

    // Extract identity from settings for easy access
    var identity = settings.public_identity || {};

    Logger.log('getPublicData: committee=' + committee.length + ' gallery=' + gallery.length + ' news=' + news.length);
    return {
      success:   true,
      settings:  settings,
      identity:  identity,
      committee: committee,
      gallery:   gallery,
      news:      news
    };
  } catch(e) {
    Logger.log('getPublicData ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── TEST ALL PUBLIC FUNCTIONS ─────────────────────────────────
function testPublicFunctions() {
  Logger.log('=== testPublicFunctions ===');
  var cm = getCommittee();
  Logger.log('getCommittee: success=' + cm.success + ' count=' + (cm.data ? cm.data.length : 0));
  if (cm.data && cm.data.length > 0) {
    Logger.log('First member: ' + JSON.stringify({name:cm.data[0].Name, role:cm.data[0].Role, hasPhoto:cm.data[0].Photo.length > 10}));
  }
  var pd = getPublicData();
  Logger.log('getPublicData: success=' + pd.success + ' committee=' + (pd.committee||[]).length + ' gallery=' + (pd.gallery||[]).length);
  var settings = getPublicDashboardSettings();
  Logger.log('Settings keys: ' + Object.keys(settings.data || {}).join(', '));
}


// ── UPDATE TEACHER ────────────────────────────────────────────
function updateTeacherDirect(data) {
  try {
    if (!data || !data.teacherId) return { success:false, message:'Teacher ID required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Teachers');
    if (!sheet) return { success:false, message:'Teachers sheet not found' };
    var rows = sheet.getDataRange().getValues();
    // Schema: TeacherID,Name,Designation,Subject,Phone,Email,JoinDate,AssignedClasses,Photo,Status,...
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.teacherId).trim()) {
        var row = i + 1;
        if (data.name)            sheet.getRange(row, 2).setValue(data.name);
        if (data.designation)     sheet.getRange(row, 3).setValue(data.designation);
        if (data.subject)         sheet.getRange(row, 4).setValue(data.subject);
        if (data.phone)           sheet.getRange(row, 5).setValue(data.phone);
        if (data.email)           sheet.getRange(row, 6).setValue(data.email);
        if (data.joinDate)        sheet.getRange(row, 7).setValue(data.joinDate);
        if (data.assignedClasses !== undefined) sheet.getRange(row, 8).setValue(data.assignedClasses);
        if (data.photo)           sheet.getRange(row, 9).setValue(data.photo);
        Logger.log('Teacher updated: ' + data.teacherId);
        return { success:true, message:'Teacher updated: ' + data.name };
      }
    }
    return { success:false, message:'Teacher not found: ' + data.teacherId };
  } catch(e) {
    Logger.log('updateTeacherDirect ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── DELETE TEACHER ────────────────────────────────────────────
function deleteTeacherDirect(data) {
  try {
    if (!data || !data.teacherId) return { success:false, message:'Teacher ID required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Teachers');
    if (!sheet) return { success:false, message:'Teachers sheet not found' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.teacherId).trim()) {
        // Mark inactive instead of deleting
        sheet.getRange(i + 1, 10).setValue('Inactive');
        return { success:true, message:'Teacher removed' };
      }
    }
    return { success:false, message:'Teacher not found' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

// ── UPDATE COMMITTEE MEMBER ───────────────────────────────────
function updateCommitteeMember(data) {
  try {
    if (!data || !data.origName) return { success:false, message:'Original name required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Committee');
    if (!sheet) return { success:false, message:'Committee sheet not found' };
    var rows = sheet.getDataRange().getValues();
    // Schema: Name,Role,Department,Phone,Email,Photo,JoinDate,Bio,CreatedAt,Status
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.origName).trim()) {
        var row = i + 1;
        sheet.getRange(row, 1).setValue(data.name        || rows[i][0]);
        sheet.getRange(row, 2).setValue(data.role        !== undefined ? data.role        : rows[i][1]);
        sheet.getRange(row, 3).setValue(data.department  !== undefined ? data.department  : rows[i][2]);
        sheet.getRange(row, 4).setValue(data.phone       !== undefined ? data.phone       : rows[i][3]);
        sheet.getRange(row, 5).setValue(data.email       !== undefined ? data.email       : rows[i][4]);
        if (data.photo) sheet.getRange(row, 6).setValue(data.photo);
        sheet.getRange(row, 8).setValue(data.bio         !== undefined ? data.bio         : rows[i][7]);
        Logger.log('Committee member updated: ' + data.origName);
        return { success:true, message:'Member updated: ' + (data.name || data.origName) };
      }
    }
    return { success:false, message:'Member not found: ' + data.origName };
  } catch(e) {
    Logger.log('updateCommitteeMember ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── GET GALLERY ───────────────────────────────────────────────
function getGallery() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Gallery');
    if (!sheet || sheet.getLastRow() < 1) return { success:true, data:[] };
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 6);
    var rows = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    // Skip header row if present
    var start = (String(rows[0][0]).trim().toLowerCase() === 'title') ? 1 : 0;
    var out = [];
    for (var i = start; i < rows.length; i++) {
      var title = String(rows[i][0] || '').trim();
      if (!title) continue;                          // skip blank rows
      var imageData = String(rows[i][1] || '').trim(); // col 2: base64 OR url
      out.push({
        GalleryID:   String(i + 1),                  // use actual row index as stable ID
        Title:       title,
        Photo:       imageData,                      // base64 or URL — same field
        ImageURL:    imageData,
        Category:    String(rows[i][2] || ''),
        Description: String(rows[i][3] || ''),
        Date:        String(rows[i][4] || '')
      });
    }
    Logger.log('getGallery: returning ' + out.length + ' items');
    return { success:true, data:out };
  } catch(e) {
    Logger.log('getGallery ERROR: ' + e.message);
    return { success:false, message:e.message };
  }
}

// ── DELETE GALLERY ITEM ───────────────────────────────────────
function deleteGalleryItem(data) {
  try {
    if (!data) return { success:false, message:'Data required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Gallery');
    if (!sheet) return { success:false, message:'Gallery sheet not found' };
    var gid = String(data.galleryId || '');
    var rows = sheet.getDataRange().getValues();
    // Try to match by row number (GalleryID = row index stored in col 7)
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][6] || (i + 1)) === gid || String(i + 1) === gid) {
        sheet.deleteRow(i + 1);
        return { success:true, message:'Gallery item deleted' };
      }
    }
    // Fallback: delete by title
    if (data.title) {
      for (var j = 1; j < rows.length; j++) {
        if (String(rows[j][0]).trim() === String(data.title).trim()) {
          sheet.deleteRow(j + 1);
          return { success:true, message:'Gallery item deleted' };
        }
      }
    }
    return { success:false, message:'Gallery item not found' };
  } catch(e) {
    return { success:false, message:e.message };
  }
}


// ══════════════════════════════════════════════════════════════
// OTP — SEND & VERIFY (Gmail Email OTP)
// Parents enter their email → receive 6-digit OTP in inbox
// No API keys or external services needed — uses Google account
// ══════════════════════════════════════════════════════════════

var OTP_EXPIRY_MS = 10 * 60 * 1000; // 10 minutes

function sendOtp(data) {
  try {
    if (!data || !data.email)
      return { success:false, message:'Email address required' };

    var email = String(data.email).trim().toLowerCase();

    // Basic email format check
    if (!email.match(/^[^@\s]+@[^@\s]+\.[^@\s]+$/))
      return { success:false, message:'Enter a valid email address' };

    // ── Generate 6-digit OTP ──────────────────────────────────
    var otp    = String(Math.floor(100000 + Math.random() * 900000));
    var expiry = new Date().getTime() + OTP_EXPIRY_MS;

    // ── Hash + store server-side (never expose OTP to client) ─
    var salt    = Utilities.getUuid();
    var hashArr = Utilities.computeDigest(
                    Utilities.DigestAlgorithm.SHA_256,
                    email + ':' + otp + ':' + salt,
                    Utilities.Charset.UTF_8);
    var hash = hashArr.map(function(b){
                 return ('0' + (b & 0xFF).toString(16)).slice(-2);
               }).join('');

    PropertiesService.getScriptProperties()
      .setProperty('OTP_' + hash, otp + ':' + expiry);

    // ── Send OTP email via GmailApp ───────────────────────────
    var ss     = SpreadsheetApp.openById(SS_ID);
    var settingsSheet = ss.getSheetByName('Settings');
    var madrasaName = 'Markaz Al Asas Academy';
    if (settingsSheet) {
      var srows = settingsSheet.getDataRange().getValues();
      for (var si = 0; si < srows.length; si++) {
        if (String(srows[si][0]).trim() === 'public_identity') {
          try {
            var id = JSON.parse(srows[si][1]);
            madrasaName = id.nameEn || id.nameShort || madrasaName;
          } catch(pe) {}
          break;
        }
      }
    }

    var subject = madrasaName + ' \u2014 Admission Application OTP';
    var body =
      'Assalamu Alaikum,\n\n' +
      'Your One-Time Password (OTP) for the admission application at ' +
      madrasaName + ' is:\n\n' +
      '    ' + otp + '\n\n' +
      'This OTP is valid for 10 minutes.\n' +
      'Do not share this OTP with anyone.\n\n' +
      'If you did not request this, please ignore this email.\n\n' +
      'Regards,\n' + madrasaName;

    var htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;' +
      'border:1px solid #f0e0f5;border-radius:14px;overflow:hidden">' +
      '<div style="background:linear-gradient(135deg,#1a0a2e,#2d0845);padding:20px;text-align:center">' +
      '<div style="font-family:serif;font-size:20px;font-weight:900;color:#F5C518">' + madrasaName + '</div>' +
      '<div style="font-size:12px;color:rgba(255,255,255,.6);margin-top:4px">Admission Application Verification</div>' +
      '</div>' +
      '<div style="padding:30px;text-align:center">' +
      '<p style="color:#555;font-size:14px;margin-bottom:20px">Your One-Time Password (OTP) is:</p>' +
      '<div style="background:#fdf5ff;border:2px solid rgba(233,30,140,.3);border-radius:12px;' +
      'padding:16px 24px;display:inline-block;margin-bottom:20px">' +
      '<span style="font-size:36px;font-weight:900;letter-spacing:10px;color:#E91E8C">' + otp + '</span>' +
      '</div>' +
      '<p style="color:#888;font-size:12px">Valid for <strong>10 minutes</strong>. Do not share this OTP.</p>' +
      '</div>' +
      '<div style="background:#f9f9f9;padding:14px;text-align:center;font-size:11px;color:#aaa">' +
      'If you did not request this, please ignore this email.' +
      '</div></div>';

    GmailApp.sendEmail(email, subject, body, { htmlBody: htmlBody });

    Logger.log('OTP email sent to ' + email + ' | OTP=' + otp);
    return {
      success:   true,
      hash:      hash,
      message:   'OTP sent to ' + email,
      emailSent: true,
      maskedEmail: email.replace(/(.{2})[^@]+(@.+)/, '$1****$2')
    };

  } catch(e) {
    Logger.log('sendOtp ERROR: ' + e.message);
    return { success:false, message:'Failed to send OTP: ' + e.message };
  }
}

function verifyOtp(data) {
  try {
    if (!data || !data.hash || !data.otp || !data.email)
      return { success:false, message:'Missing verification data' };

    var store  = PropertiesService.getScriptProperties();
    var saved  = store.getProperty('OTP_' + data.hash);
    if (!saved)
      return { success:false, message:'OTP expired or not found. Please request a new OTP.' };

    var parts    = saved.split(':');
    var savedOtp = parts[0];
    var expiry   = parseInt(parts[1]);

    if (new Date().getTime() > expiry) {
      store.deleteProperty('OTP_' + data.hash);
      return { success:false, message:'OTP has expired. Please request a new one.' };
    }

    if (String(data.otp).trim() !== savedOtp)
      return { success:false, message:'Incorrect OTP. Please try again.' };

    // Valid — delete immediately so it cannot be reused
    store.deleteProperty('OTP_' + data.hash);
    Logger.log('OTP verified: email=' + data.email);
    return { success:true, message:'Email verified successfully' };

  } catch(e) {
    Logger.log('verifyOtp ERROR: ' + e.message);
    return { success:false, message:'Verification error: ' + e.message };
  }
}


// ── FIX ALL — Run once to set up all sheets ───────────────────
function fixAll() {
  Logger.log('fixAll() starting...');
  var ss = SpreadsheetApp.openById(SS_ID);
  var schemas = {
    Students:   ['StudentID','AdmissionNo','Name','NameArabic','DOB','Gender','Class','Section','FatherName','MotherName','Phone','Address','Email','Photo','CreatedAt','Status'],
    Teachers:   ['TeacherID','Name','Designation','Subject','Phone','Email','JoinDate','Qualification','Photo','CreatedAt','Status'],
    Attendance: ['Date','StudentID','Status','ClassID','TeacherID','CreatedAt'],
    Marks:      ['StudentID','ExamName','AcademicYear','Class','Subject','MaxMarks','MarksObtained','Grade','TeacherID','CreatedAt'],
    Admissions: ['ApplicationID','StudentName','DOB','Gender','ApplyingForClass','FatherName','MotherName','Phone','Email','Address','PreviousMadrasa','PrevRegNo','AcademicYear','Photo','Documents','SubmittedAt','Status'],
    Gallery:    ['Title','Photo','Category','Description','Date','UploadedBy'],
    News:       ['Title','Content','Category','Image','CreatedAt','Author','Status'],
    Events:     ['Title','Date','Time','Venue','Description','Category','Image','CreatedAt','Status'],
    Users:      ['UserID','Username','Password','Role','Name','Email','Active'],
    Committee:  ['Name','Role','Department','Phone','Email','Photo','JoinDate','Bio','CreatedAt','Status'],
    Fees:       ['ReceiptID','StudentID','StudentName','Class','FeeType','Amount','Month','AcademicYear','PaidAt','CollectedBy','Status'],
    AdmitCards: ['CardID','StudentID','StudentName','Class','AdmissionNo','ExamName','ExamDate','Venue','GeneratedAt'],
    Settings:   ['Key','Value','UpdatedAt']
  };

  Object.keys(schemas).forEach(function(name) {
    var sheet   = ss.getSheetByName(name);
    var headers = schemas[name];
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      var r = sheet.getRange(1, 1, 1, headers.length);
      r.setBackground('#E91E8C'); r.setFontColor('#fff'); r.setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log('Created sheet: ' + name);
      return;
    }
    var rows = sheet.getDataRange().getValues();
    if (rows.length === 0) {
      sheet.appendRow(headers);
      Logger.log('Added header to empty: ' + name);
      return;
    }
    if (String(rows[0][0]).trim() !== headers[0]) {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      var r2 = sheet.getRange(1, 1, 1, headers.length);
      r2.setBackground('#E91E8C'); r2.setFontColor('#fff'); r2.setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log('Inserted header: ' + name);
    } else {
      Logger.log('OK: ' + name);
    }
  });

  // Default users
  var us = ss.getSheetByName('Users');
  var existing = us.getDataRange().getValues().slice(1).map(function(r){ return String(r[1]).toLowerCase(); });
  var defs = [
    ['USR001','admin','admin123','admin','Administrator','admin@alasas.edu',true],
    ['USR002','teacher01','teacher123','teacher','Demo Teacher','teacher@alasas.edu',true],
    ['USR003','parent01','parent123','parent','Demo Parent','parent@alasas.edu',true],
    ['USR004','student01','student123','student','Demo Student','student@alasas.edu',true]
  ];
  defs.forEach(function(row) {
    if (existing.indexOf(row[1]) === -1) { us.appendRow(row); Logger.log('Added user: ' + row[1]); }
  });

  // Ensure all users have Active = true
  var last = us.getLastRow();
  if (last > 1) {
    var rng  = us.getRange(2, 7, last - 1, 1);
    var vals = rng.getValues();
    vals.forEach(function(r, i){ vals[i][0] = true; });
    rng.setValues(vals);
  }

  Logger.log('fixAll() complete!');
  Logger.log('Default logins: admin/admin123 | teacher01/teacher123 | parent01/parent123 | student01/student123');
  return 'Done! Run testGetAdmissions() to verify, then deploy a NEW VERSION.';
}
// ── GET COMMITTEE ─────────────────────────────────────────────
function getCommittee() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Committee');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    var rows = sheet.getRange(1, 1, sheet.getLastRow(), 10).getValues();
    var start = String(rows[0][0]).trim() === 'Name' ? 1 : 0;
    var out = [];
    for (var i = start; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (String(rows[i][9] || 'Active').trim() === 'Inactive') continue;
      out.push({
        Name:       String(rows[i][0] || ''),
        Role:       String(rows[i][1] || ''),
        Department: String(rows[i][2] || ''),
        Phone:      String(rows[i][3] || ''),
        Email:      String(rows[i][4] || ''),
        Photo:      String(rows[i][5] || ''),
        Bio:        String(rows[i][7] || ''),
        Status:     String(rows[i][9] || 'Active')
      });
    }
    Logger.log('getCommittee: returning ' + out.length + ' members');
    return { success: true, data: out };
  } catch(e) {
    Logger.log('getCommittee ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── ADD COMMITTEE MEMBER (with photo) ─────────────────────────
function addCommitteeDirect(data) {
  try {
    if (!data || !data.name) return { success: false, message: 'Name required' };
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Committee');
    if (!sheet) return { success: false, message: 'Committee sheet not found. Run fixAll().' };
    // Schema: Name,Role,Department,Phone,Email,Photo,JoinDate,Bio,CreatedAt,Status
    sheet.appendRow([
      data.name,
      data.role        || '',
      data.department  || '',
      data.phone       || '',
      data.email       || '',
      data.photo       || '',   // col 6 — base64 or URL
      data.joinDate    || '',
      data.bio         || '',
      new Date().toISOString(),
      'Active'
    ]);
    Logger.log('Committee member added: ' + data.name);
    return { success: true, message: 'Member added: ' + data.name };
  } catch(e) {
    Logger.log('addCommitteeDirect ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── GET PUBLIC DASHBOARD SETTINGS ────────────────────────────
function getPublicDashboardSettings() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { success: true, data: {} };
    var rows = sheet.getDataRange().getValues();
    var settings = {};
    rows.slice(1).forEach(function(r) {
      var key = String(r[0] || '');
      if (key.indexOf('public_') === 0) {
        try { settings[key] = JSON.parse(String(r[1] || '{}')); }
        catch(e) { settings[key] = String(r[1] || ''); }
      }
    });
    return { success: true, data: settings };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── SAVE PUBLIC DASHBOARD SETTINGS ───────────────────────────
function savePublicDashboardSettings(data) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { success: false, message: 'Settings sheet not found' };
    var rows = sheet.getDataRange().getValues();
    var now  = new Date().toISOString();
    var updates = data.settings || {};
    Object.keys(updates).forEach(function(key) {
      var val  = typeof updates[key] === 'object' ? JSON.stringify(updates[key]) : String(updates[key]);
      var found = false;
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(val);
          sheet.getRange(i + 1, 3).setValue(now);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([key, val, now]);
    });
    return { success: true, message: 'Public dashboard settings saved' };
  } catch(e) {
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
        return { success: true, message: 'Member removed: ' + data.name };
      }
    }
    return { success: false, message: 'Member not found' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
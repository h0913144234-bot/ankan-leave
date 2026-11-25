var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var STAFF_SHEET_NAME = "Staff";
var LEAVES_SHEET_NAME = "Leaves";
var META_SHEET_NAME = "Meta";

function getOrCreateSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

function writeStaffSheet_(ss, staffList) {
  var sheet = getOrCreateSheet_(ss, STAFF_SHEET_NAME, ["name", "block", "role", "emp_id", "phone", "updated_at"]);
  var nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  var rows = [["name", "block", "role", "emp_id", "phone", "updated_at"]];
  (staffList || []).forEach(function (p) {
    if (!p || !p.name) return;
    rows.push([
      String(p.name).trim(),
      String(p.block || "").trim(),
      String(p.role || "").trim(),
      String(p.emp_id || "").trim(),
      String(p.phone || "").trim(),
      nowStr
    ]);
  });
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

function readStaffSheet_(ss) {
  var sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var name = String(values[i][0] || "").trim();
    if (!name) continue;
    list.push({
      name: name,
      block: String(values[i][1] || "").trim(),
      role: String(values[i][2] || "").trim(),
      emp_id: String(values[i][3] || "").trim(),
      phone: String(values[i][4] || "").trim()
    });
  }
  return list;
}

function writePersonLeaves_(ss, staffName, dates) {
  var sheet = getOrCreateSheet_(ss, LEAVES_SHEET_NAME, ["staff_name", "leave_date", "updated_at"]);
  var values = sheet.getDataRange().getValues();
  var keepRows = [["staff_name", "leave_date", "updated_at"]];
  for (var i = 1; i < values.length; i++) {
    var rName = String(values[i][0] || "").trim();
    if (rName && rName !== staffName) {
      keepRows.push([rName, String(values[i][1] || "").trim(), String(values[i][2] || "")]);
    }
  }
  var nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  var uniqueDates = Array.from(new Set((dates || []).filter(Boolean))).sort();
  uniqueDates.forEach(function (d) {
    keepRows.push([staffName, String(d).trim(), nowStr]);
  });
  sheet.clearContents();
  sheet.getRange(1, 1, keepRows.length, 3).setValues(keepRows);
}

function writeAllLeaves_(ss, leavesMap) {
  var sheet = getOrCreateSheet_(ss, LEAVES_SHEET_NAME, ["staff_name", "leave_date", "updated_at"]);
  var nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  var rows = [["staff_name", "leave_date", "updated_at"]];
  Object.keys(leavesMap || {}).forEach(function (name) {
    var cleanName = String(name || "").trim();
    if (!cleanName) return;
    var dates = Array.from(new Set((leavesMap[name] || []).filter(Boolean))).sort();
    dates.forEach(function (d) {
      rows.push([cleanName, String(d).trim(), nowStr]);
    });
  });
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
}

function setMetaValue_(ss, key, val) {
  var sheet = getOrCreateSheet_(ss, META_SHEET_NAME, ["key", "value"]);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === key) {
      sheet.getRange(i + 1, 2).setValue(String(val || ""));
      return;
    }
  }
  sheet.appendRow([key, String(val || "")]);
}

function getMetaValue_(ss, key) {
  var sheet = ss.getSheetByName(META_SHEET_NAME);
  if (!sheet) return "";
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === key) {
      return String(values[i][1] || "").trim();
    }
  }
  return "";
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var leavesSheet = ss.getSheetByName(LEAVES_SHEET_NAME);
    var leavesMap = {};
    if (leavesSheet) {
      var data = leavesSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var name = String(data[i][0] || "").trim();
        var dateStr = String(data[i][1] || "").trim();
        if (!name || !dateStr) continue;
        if (!leavesMap[name]) leavesMap[name] = [];
        if (leavesMap[name].indexOf(dateStr) === -1) leavesMap[name].push(dateStr);
      }
      Object.keys(leavesMap).forEach(function (k) { leavesMap[k].sort(); });
    }

    var staffList = readStaffSheet_(ss);
    var startDate = getMetaValue_(ss, "startDate");

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        staff: staffList,
        leaves: leavesMap,
        startDate: startDate
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var payload = JSON.parse(raw);
    var action = payload.action || "";
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === "sync_person") {
      writePersonLeaves_(ss, String(payload.staffName || "").trim(), payload.dates || []);
    } else if (action === "sync_all") {
      if (payload.leaves) writeAllLeaves_(ss, payload.leaves);
      if (Array.isArray(payload.staff)) writeStaffSheet_(ss, payload.staff);
      if (payload.startDate) setMetaValue_(ss, "startDate", payload.startDate);
    } else if (action === "add_staff") {
      var listAdd = readStaffSheet_(ss);
      var newP = payload.staffItem || {};
      var addName = String(newP.name || "").trim();
      if (addName) {
        var idxAdd = listAdd.findIndex(function (x) { return x.name === addName; });
        var itemAdd = {
          name: addName,
          block: String(newP.block || "").trim(),
          role: String(newP.role || "").trim(),
          emp_id: String(newP.emp_id || "").trim(),
          phone: String(newP.phone || "").trim()
        };
        if (idxAdd >= 0) listAdd[idxAdd] = itemAdd;
        else listAdd.push(itemAdd);
        writeStaffSheet_(ss, listAdd);
      }
    } else if (action === "update_staff") {
      var listUpd = readStaffSheet_(ss);
      var oldName = String(payload.oldName || "").trim();
      var updP = payload.staffItem || {};
      var updName = String(updP.name || "").trim();
      if (updName) {
        var idxUpd = listUpd.findIndex(function (x) { return x.name === oldName; });
        var itemUpd = {
          name: updName,
          block: String(updP.block || "").trim(),
          role: String(updP.role || "").trim(),
          emp_id: String(updP.emp_id || "").trim(),
          phone: String(updP.phone || "").trim()
        };
        if (idxUpd >= 0) listUpd[idxUpd] = itemUpd;
        else listUpd.push(itemUpd);
        writeStaffSheet_(ss, listUpd);

        if (oldName && oldName !== updName) {
          var lSheet = getOrCreateSheet_(ss, LEAVES_SHEET_NAME, ["staff_name", "leave_date", "updated_at"]);
          var lVals = lSheet.getDataRange().getValues();
          var movedDates = [];
          var otherRows = [["staff_name", "leave_date", "updated_at"]];
          for (var i = 1; i < lVals.length; i++) {
            var rn = String(lVals[i][0] || "").trim();
            var rd = String(lVals[i][1] || "").trim();
            var rt = String(lVals[i][2] || "");
            if (rn === oldName) {
              if (rd) movedDates.push(rd);
            } else {
              otherRows.push([rn, rd, rt]);
            }
          }
          var nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
          Array.from(new Set(movedDates)).sort().forEach(function (md) {
            otherRows.push([updName, md, nowStr]);
          });
          lSheet.clearContents();
          lSheet.getRange(1, 1, otherRows.length, 3).setValues(otherRows);
        }
      }
    } else if (action === "delete_staff") {
      var delName = String(payload.staffName || "").trim();
      if (delName) {
        var listDel = readStaffSheet_(ss).filter(function (x) { return x.name !== delName; });
        writeStaffSheet_(ss, listDel);
        var dSheet = getOrCreateSheet_(ss, LEAVES_SHEET_NAME, ["staff_name", "leave_date", "updated_at"]);
        var dVals = dSheet.getDataRange().getValues();
        var kept = [["staff_name", "leave_date", "updated_at"]];
        for (var j = 1; j < dVals.length; j++) {
          var dn = String(dVals[j][0] || "").trim();
          if (dn && dn !== delName) {
            kept.push([dn, String(dVals[j][1] || "").trim(), String(dVals[j][2] || "")]);
          }
        }
        dSheet.clearContents();
        dSheet.getRange(1, 1, kept.length, 3).setValues(kept);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
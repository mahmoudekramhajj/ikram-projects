/**
 * ReportApp.js — Ikram Report Builder v2.0
 * بحث شامل عبر كل الشيتات
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.key) {
    return handleClaudeAPI_(e);
  }

  return HtmlService.createTemplateFromFile('ReportIndex')
    .evaluate()
    .setTitle('Ikram Report Builder')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ══════════════════════════════════════════════
//  API Endpoints
// ══════════════════════════════════════════════

/** قائمة كل الشيتات مع أعمدتها (كشف تلقائي) */
function getAllSheets() {
  return JSON.stringify(getAllSheetsWithColumns());
}

/** جلب البيانات مع فلترة + تجميع + pagination */
function getReportData(paramsJson) {
  var params = JSON.parse(paramsJson);
  return JSON.stringify(fetchReportData(params));
}

/** القيم الفريدة لعمود في شيت */
function getUniqueValues(sheetName, colIndex) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'uv_' + sheetName + '_' + colIndex;
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return JSON.stringify([]);

  var cfg = SHEET_CONFIG[sheetName] || {};
  var startRow = cfg.dataStartRow || DEFAULT_START_ROW;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return JSON.stringify([]);

  var col = colIndex + 1;
  var values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();
  var unique = {};
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0]).trim();
    if (v && v !== 'undefined') unique[v] = true;
  }

  var result = Object.keys(unique).sort();
  var json = JSON.stringify(result);
  if (json.length < 100000) cache.put(cacheKey, json, CACHE_TTL);
  return json;
}

/** تصدير Excel */
function exportReport(paramsJson) {
  var params = JSON.parse(paramsJson);
  params.page = 1;
  params.pageSize = 999999;
  var data = fetchReportData(params);

  var ss = SpreadsheetApp.create('Report_' + Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyyMMdd_HHmm'));
  var sheet = ss.getActiveSheet();
  sheet.setName('Report');

  // Headers
  var colsMeta = data.columnsMeta || data.groupColumns || [];
  var headers = colsMeta.map(function(c) { return c.name; });
  if (headers.length === 0 && data.rows && data.rows.length > 0) {
    headers = data.rows[0].map(function(_, i) { return 'Col ' + (i + 1); });
  }

  sheet.appendRow(headers);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a4d2e').setFontColor('#ffffff').setFontWeight('bold')
    .setFontFamily('Cairo').setHorizontalAlignment('center');

  // Data
  var rows = data.rows || [];
  for (var i = 0; i < rows.length; i++) {
    sheet.appendRow(rows[i]);
  }

  // تنسيق
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setFontFamily('Cairo');
    for (var r = 0; r < rows.length; r++) {
      if (r % 2 === 1) {
        sheet.getRange(r + 2, 1, 1, headers.length).setBackground('#f8f9fa');
      }
    }
  }

  // Totals footer
  if (data.totals && data.totals.length > 0) {
    sheet.appendRow(data.totals);
    var footRow = rows.length + 2;
    sheet.getRange(footRow, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#e8f5e9').setFontFamily('Cairo');
  }

  sheet.setFrozenRows(1);
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
  }

  SpreadsheetApp.flush();

  var fileId = ss.getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';

  // حذف تلقائي بعد ساعة
  ScriptApp.newTrigger('deleteExportFile_').timeBased().after(3600000).create();
  PropertiesService.getScriptProperties().setProperty('deleteFileId', fileId);

  return JSON.stringify({ success: true, url: url, fileName: ss.getName() + '.xlsx' });
}

function deleteExportFile_() {
  try {
    var fileId = PropertiesService.getScriptProperties().getProperty('deleteFileId');
    if (fileId) {
      DriveApp.getFileById(fileId).setTrashed(true);
      PropertiesService.getScriptProperties().deleteProperty('deleteFileId');
    }
  } catch (e) {}
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'deleteExportFile_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ══════════════════════════════════════════════
//  Templates
// ══════════════════════════════════════════════

var TEMPLATES_SHEET = 'Report Templates';

function saveTemplate(name, configJson) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TEMPLATES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TEMPLATES_SHEET);
    sheet.appendRow(['Name', 'Config', 'Created', 'LastUsed']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.getRange(i + 1, 2).setValue(configJson);
      sheet.getRange(i + 1, 4).setValue(new Date());
      return JSON.stringify({ success: true, updated: true });
    }
  }

  sheet.appendRow([name, configJson, new Date(), new Date()]);
  return JSON.stringify({ success: true, updated: false });
}

function loadTemplates() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TEMPLATES_SHEET);
  if (!sheet) return JSON.stringify([]);

  var data = sheet.getDataRange().getValues();
  var templates = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      templates.push({
        name: data[i][0],
        created: data[i][2] ? Utilities.formatDate(new Date(data[i][2]), 'Asia/Riyadh', 'yyyy-MM-dd') : '',
        lastUsed: data[i][3] ? Utilities.formatDate(new Date(data[i][3]), 'Asia/Riyadh', 'yyyy-MM-dd') : ''
      });
    }
  }
  return JSON.stringify(templates);
}

function loadTemplate(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TEMPLATES_SHEET);
  if (!sheet) return JSON.stringify(null);

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.getRange(i + 1, 4).setValue(new Date());
      return data[i][1] || JSON.stringify(null);
    }
  }
  return JSON.stringify(null);
}

function deleteTemplate(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TEMPLATES_SHEET);
  if (!sheet) return JSON.stringify({ success: false });

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.deleteRow(i + 1);
      return JSON.stringify({ success: true });
    }
  }
  return JSON.stringify({ success: false });
}

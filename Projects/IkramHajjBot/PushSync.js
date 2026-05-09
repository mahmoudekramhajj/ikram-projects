// ============================================
// PushSync — يدفع البيانات من Google Sheets
// إلى HajjBotServer (Node.js / Railway)
// ============================================
// إعدادات:
//   Script Properties:
//     SERVER_URL    = https://your-server.railway.app
//     SYNC_TOKEN    = نفس قيمة SYNC_TOKEN في السيرفر
//     TENANT_ID     = ikram (أو أي معرّف tenant)
// ============================================

var PUSH_SHEETS = [
  { sheet: 'Presonal Details', mode: 'full' },
  { sheet: 'الباقات',          mode: 'full' },
  { sheet: 'الطيران',          mode: 'full' },
  { sheet: 'B2C',              mode: 'full' },
  { sheet: 'BotSessions',      mode: 'full' }
];

// ============================================
// دفع كامل لكل الشيتات (يُشغَّل يدوياً أو عند البداية)
// ============================================
function pushAllDataToServer() {
  var props = PropertiesService.getScriptProperties();
  var serverUrl = props.getProperty('SERVER_URL');
  var syncToken = props.getProperty('SYNC_TOKEN');
  var tenantId  = props.getProperty('TENANT_ID') || 'ikram';

  if (!serverUrl || !syncToken) {
    throw new Error('Set SERVER_URL and SYNC_TOKEN in Script Properties first');
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var results = [];

  for (var i = 0; i < PUSH_SHEETS.length; i++) {
    var conf = PUSH_SHEETS[i];
    var sheet = ss.getSheetByName(conf.sheet);
    if (!sheet) {
      results.push({ sheet: conf.sheet, ok: false, error: 'sheet not found' });
      continue;
    }
    var rows = sheet.getDataRange().getValues();
    var res = pushSheet_(serverUrl, syncToken, tenantId, conf.sheet, rows, 'full');
    results.push({ sheet: conf.sheet, rows: rows.length, response: res });
    Utilities.sleep(500); // breathing room
  }

  Logger.log('pushAllDataToServer results: ' + JSON.stringify(results, null, 2));
  return results;
}

// ============================================
// دفع شيت واحد
// ============================================
function pushSheet_(serverUrl, syncToken, tenantId, sheetName, rows, mode, keyCol) {
  var payload = {
    tenant_id: tenantId,
    sheet_name: sheetName,
    mode: mode || 'full',
    rows: rows
  };
  if (mode === 'delta' && keyCol != null) payload.key_col = keyCol;

  var resp = UrlFetchApp.fetch(serverUrl + '/api/sync/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Sync-Token': syncToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return { code: resp.getResponseCode(), body: resp.getContentText().substring(0, 500) };
}

// ============================================
// Trigger كل 30 دقيقة (إعداد مرة واحدة)
// ============================================
function setupPushSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pushAllDataToServer') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('pushAllDataToServer')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('Trigger created: pushAllDataToServer every 30 min');
}

// ============================================
// إعداد سريع (اعرض النموذج ثم املأه)
// ============================================
function setupSyncProperties_TEMPLATE() {
  var props = PropertiesService.getScriptProperties();
  // عدّل القيم ثم شغّل الدالة:
  props.setProperty('SERVER_URL', 'https://your-server.railway.app');
  props.setProperty('SYNC_TOKEN', 'replace-with-strong-random-token');
  props.setProperty('TENANT_ID', 'ikram');
  Logger.log('Properties saved. Now run pushAllDataToServer().');
}

// ============================================
// اختبار سريع — push شيت واحد فقط
// ============================================
function testPushSinglePD_() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rows = ss.getSheetByName('Presonal Details').getDataRange().getValues();
  var res = pushSheet_(
    props.getProperty('SERVER_URL'),
    props.getProperty('SYNC_TOKEN'),
    props.getProperty('TENANT_ID') || 'ikram',
    'Presonal Details',
    rows,
    'full'
  );
  Logger.log(JSON.stringify(res));
}

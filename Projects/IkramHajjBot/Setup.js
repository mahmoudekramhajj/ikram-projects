// ============================================
// إعداد Webhook + اختبار + debug
// ============================================

function setWebhook() {
  var webAppUrl = 'https://script.google.com/macros/s/AKfycbydPqakby_TjKOUMMXAtTiynEiu_96bZ7W6xSCr11EAf7f1-UnL5duR-_cc0u5xjwZlhg/exec';
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/setWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: webAppUrl }),
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

function getWebhookInfo() {
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/getWebhookInfo', { muteHttpExceptions: true });
  Logger.log(res.getContentText());
}

function testBot() {
  Logger.log('=== Test 1: BotSessions ===');
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BotSessions');
    Logger.log('Sheet found: ' + (sheet ? 'YES' : 'NO'));
    Logger.log('Sheet name: ' + sheet.getName());
    sheet.appendRow(['TEST123', '', '', 'ar', 'test', '', new Date().toISOString(), 0]);
    Logger.log('Write OK');
  } catch(e) {
    Logger.log('Write ERROR: ' + e.message);
  }

  Logger.log('=== Test 2: Journey Sheet ===');
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var jSheet = ss.getSheetByName(JOURNEY_SHEET);
    Logger.log('Journey sheet found: ' + (jSheet ? 'YES' : 'NO'));
    if (jSheet) {
      var headers = jSheet.getRange(1, 1, 1, 15).getValues()[0];
      Logger.log('First 15 headers: ' + headers.join(' | '));
    }
  } catch(e) {
    Logger.log('Journey ERROR: ' + e.message);
  }

  Logger.log('=== Test 3: Telegram ===');
  try {
    var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: '121527633',
        text: '✅ تجربة إرسال من testBot'
      }),
      muteHttpExceptions: true
    });
    Logger.log('TG Response: ' + res.getContentText());
  } catch(e) {
    Logger.log('TG ERROR: ' + e.message);
  }
}

function clearAllCache() {
  CacheService.getScriptCache().removeAll([]);
  // removeAll with empty doesn't work, use a different approach
  var cache = CacheService.getScriptCache();
  // Can't enumerate keys, so just log
  Logger.log('Note: CacheService cannot enumerate keys. Cache will expire naturally (30 min TTL).');
  Logger.log('Workaround: setting TTL to 1 second for known patterns.');
}

function debugHotelMap() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('الفنادق');

  if (!sheet) { Logger.log('ERROR: sheet not found'); return; }

  var data = sheet.getDataRange().getValues();
  Logger.log('Total rows: ' + data.length + ' | cols: ' + data[0].length);

  Logger.log('=== Row 0 (headers) ===');
  for (var c = 0; c < Math.min(8, data[0].length); c++) {
    Logger.log('  [' + c + '] = "' + String(data[0][c]) + '"');
  }

  Logger.log('=== Row 1 (first data) ===');
  for (var c2 = 0; c2 < Math.min(8, data[1].length); c2++) {
    Logger.log('  [' + c2 + '] = "' + String(data[1][c2]).substring(0, 60) + '"');
  }

  Logger.log('=== Test getHotelMapLink_ ===');
  var testNames = ['ديار التقوى', 'فندق ديار التقوى الفندقية', 'بارك بلازا', 'فندق بارك بلازا الفندقية', 'Diyar Al Taqwa Hotel', 'Park Plaza Hotel'];
  for (var t = 0; t < testNames.length; t++) {
    CacheService.getScriptCache().remove('hotelmap_' + testNames[t]);
    var result = getHotelMapLink_(testNames[t]);
    CacheService.getScriptCache().remove('hotelmap_' + testNames[t]);
    Logger.log(testNames[t] + ' → ' + (result ? result : 'NULL'));
  }
}

function debugTransport() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var jSheet = ss.getSheetByName(JOURNEY_SHEET);
  var jData = jSheet.getDataRange().getValues();
  for (var i = 1; i < jData.length; i++) {
    if (String(jData[i][8]).toUpperCase().trim() === '674711081') {
      Logger.log('Journey PackageId: [' + jData[i][1] + '] type: ' + typeof jData[i][1]);
      break;
    }
  }

  var pSheet = ss.getSheetByName('الباقات');
  var pData = pSheet.getDataRange().getValues();
  Logger.log('Row0 headers count: ' + pData[0].length);
  Logger.log('Row1 BN value: [' + pData[1][65] + ']');
  for (var j = 2; j < 7 && j < pData.length; j++) {
    Logger.log('Row' + j + ' B=[' + pData[j][1] + '] type=' + typeof pData[j][1] + ' BN=[' + pData[j][65] + ']');
  }
}

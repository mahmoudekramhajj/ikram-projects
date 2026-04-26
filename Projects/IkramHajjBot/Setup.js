// ============================================
// إعداد Webhook + اختبار + debug
// ============================================

function setWebhook() {
  var webAppUrl = 'https://script.google.com/macros/s/AKfycbyRxUEkiKfLHB9aUZn0m2Mc7jgGgCC9ltdXD5qug2ZDiuDEflzei11oewnR1cR3EcgICg/exec';
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

  Logger.log('=== Test 2: Presonal Details ===');
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var pdSheet = ss.getSheetByName(PERSONAL_SHEET);
    Logger.log('Presonal Details sheet found: ' + (pdSheet ? 'YES' : 'NO'));
    if (pdSheet) {
      var headers = pdSheet.getRange(1, 1, 1, 15).getValues()[0];
      Logger.log('First 15 headers: ' + headers.join(' | '));
    }
  } catch(e) {
    Logger.log('Presonal Details ERROR: ' + e.message);
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
  var cache = CacheService.getScriptCache();
  Logger.log('Note: CacheService cannot enumerate keys. Cache will expire naturally (5 min TTL).');
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

  var pdSheet = ss.getSheetByName(PERSONAL_SHEET);
  var pdData = pdSheet.getDataRange().getValues();
  for (var i = 1; i < pdData.length; i++) {
    if (String(pdData[i][PD.PASSPORT]).toUpperCase().trim() === '674711081') {
      Logger.log('PD PackageNo: [' + pdData[i][PD.PACKAGE_NO] + '] type: ' + typeof pdData[i][PD.PACKAGE_NO]);
      break;
    }
  }

  var pSheet = ss.getSheetByName(PACKAGES_SHEET);
  var pData = pSheet.getDataRange().getValues();
  Logger.log('Row0 headers count: ' + pData[0].length);
  Logger.log('Row1 BN value: [' + pData[1][PKG.TRANSPORT] + ']');
  for (var j = 2; j < 7 && j < pData.length; j++) {
    Logger.log('Row' + j + ' B=[' + pData[j][PKG.NUSK_NO] + '] type=' + typeof pData[j][PKG.NUSK_NO] + ' BN=[' + pData[j][PKG.TRANSPORT] + ']');
  }
}

function debugFlight() {
  var passport = '25FD36016';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pdSheet = ss.getSheetByName(PERSONAL_SHEET);
  var pdData = pdSheet.getDataRange().getValues();
  var pd = null;
  for (var i = 1; i < pdData.length; i++) {
    if (String(pdData[i][PD.PASSPORT]).toUpperCase().trim() === passport) { pd = pdData[i]; break; }
  }
  if (!pd) { Logger.log('not found'); return; }

  var contractName = String(pd[PD.CONTRACT_NAME] || '').trim();
  var fltSheet = ss.getSheetByName(FLIGHTS_SHEET);
  var fltData = fltSheet.getDataRange().getValues();

  // Find the row
  var row = -1;
  for (var f = 2; f < fltData.length; f++) {
    if (String(fltData[f][FLT.CONTRACT_NAME] || '').trim() === contractName) { row = f; break; }
  }
  if (row === -1) { Logger.log('row not found'); return; }

  // Print columns 35-50 headers (row 0 + row 1) and data
  Logger.log('=== COLUMNS 35-50 (RET1 + RET2 area) ===');
  for (var c = 35; c <= 50; c++) {
    var colLetter = getColLetter_(c);
    Logger.log('[' + c + '] ' + colLetter + ': header0=[' + fltData[0][c] + '] header1=[' + fltData[1][c] + '] data=[' + fltData[row][c] + ']');
  }
}

function getColLetter_(n) {
  var s = '';
  n++;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}
// v86-webhook

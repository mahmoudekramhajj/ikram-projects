// B2C Sync — deprecated
// All functions moved to GDS_B2C_Complete.js (B2C Complete)
// This project is no longer active.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚠️ B2C Sync (قديم)')
    .addItem('ℹ️ تم النقل إلى B2C Complete', 'showMigrationNotice')
    .addToUi();
}

function showMigrationNotice() {
  SpreadsheetApp.getUi().alert(
    '⚠️ هذا السكربت قديم\n\n' +
    'تم دمج جميع الوظائف في سكربت B2C Complete.\n' +
    'استخدم قائمة "🔄 B2C Complete" بدلاً من هذا.'
  );
}

// Old triggers may still call these — stubs prevent errors
function syncB2C() { Logger.log('B2C Sync deprecated — use B2C Complete'); }
function syncNamesOnly() { Logger.log('deprecated'); }
function syncFlightsOnly() { Logger.log('deprecated'); }

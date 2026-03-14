const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s';
const CREDENTIALS_FILE = path.join(__dirname, 'credentials.json');
const OUTPUT_DIR = path.join(__dirname, 'sheets');

async function main() {
  // إعداد المصادقة
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS_FILE,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  console.log('جاري الاتصال بـ Google Sheets...');

  // جلب قائمة جميع الأوراق
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });

  const sheetNames = spreadsheet.data.sheets.map(s => s.properties.title);
  console.log(`تم العثور على ${sheetNames.length} ورقة`);

  // إنشاء مجلد الإخراج
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
  }

  // تصدير كل ورقة
  for (const sheetName of sheetNames) {
    try {
      console.log(`جاري تصدير: ${sheetName}...`);

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'`,
      });

      const rows = response.data.values || [];

      // أول صف هو العناوين
      const headers = rows[0] || [];
      const data = rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((header, i) => {
          obj[header || `col_${i}`] = row[i] || '';
        });
        return obj;
      });

      // حفظ كـ JSON
      const fileName = sheetName.replace(/[\\/:*?"<>|]/g, '_') + '.json';
      const filePath = path.join(OUTPUT_DIR, fileName);
      fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');

      console.log(`  ✓ ${fileName} (${data.length} صف)`);
    } catch (err) {
      console.log(`  ✗ خطأ في ${sheetName}: ${err.message}`);
    }
  }

  // حفظ ملف الفهرس
  const index = {
    lastUpdated: new Date().toISOString(),
    spreadsheetId: SPREADSHEET_ID,
    sheets: sheetNames,
  };
  fs.writeFileSync(path.join(OUTPUT_DIR, '_index.json'), JSON.stringify(index, null, 2));

  console.log('\n✅ تم التصدير بنجاح!');
  console.log(`📁 الملفات محفوظة في: ${OUTPUT_DIR}`);
}

main().catch(err => {
  console.error('خطأ:', err.message);
  process.exit(1);
});

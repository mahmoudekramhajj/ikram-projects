const fs = require('fs');

const guides   = JSON.parse(fs.readFileSync('D:/Ekram Aldyf/Projects/SheetData/sheets/Tour Guide.json', 'utf8'));
const pilgrims = JSON.parse(fs.readFileSync('D:/Ekram Aldyf/Projects/SheetData/sheets/Presonal Details.json', 'utf8'));

// خريطة جواز → باقة
const pkgMap = {};
pilgrims.forEach(p => {
  const passport = (p['رقم جواز السفر'] || '').toString().trim().toUpperCase();
  const pkg      = (p['اسم الباقة'] || '').toString().trim();
  if (passport && pkg) pkgMap[passport] = pkg;
});

// لكل باقة: قائمة المرشدين وعدد حجاجهم
const pkgStats = {};
guides.forEach(g => {
  const guide    = (g['Guide Name'] || '').toString().trim();
  const passport = (g['Passport Number'] || '').toString().trim().toUpperCase();
  const pkg      = pkgMap[passport] || (g['Package Name'] || '').toString().trim();
  if (!guide || !pkg) return;
  if (!pkgStats[pkg]) pkgStats[pkg] = {};
  pkgStats[pkg][guide] = (pkgStats[pkg][guide] || 0) + 1;
});

// ترتيب الباقات حسب عدد الحجاج الإجمالي
const sortedPkgs = Object.entries(pkgStats).map(([pkg, guideMap]) => {
  const total = Object.values(guideMap).reduce((a, b) => a + b, 0);
  return { pkg, guideMap, total };
}).sort((a, b) => b.total - a.total);

console.log('=== إحصائية المرشدين لكل باقة ===\n');
sortedPkgs.forEach(({ pkg, guideMap, total }) => {
  console.log('📦 ' + pkg + ' (' + total + ' حاج)');
  const sortedGuides = Object.entries(guideMap).sort((a, b) => b[1] - a[1]);
  sortedGuides.forEach(([guide, count]) => {
    console.log('   👤 ' + guide + ': ' + count);
  });
  console.log('');
});

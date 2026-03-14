const fs = require('fs');

const guides   = JSON.parse(fs.readFileSync('D:/Ekram Aldyf/Projects/SheetData/sheets/Tour Guide.json', 'utf8'));
const pilgrims = JSON.parse(fs.readFileSync('D:/Ekram Aldyf/Projects/SheetData/sheets/Presonal Details.json', 'utf8'));

// بناء خريطة الجواز → باقة من Presonal Details
const pkgMap = {};
pilgrims.forEach(p => {
  const passport = (p['رقم جواز السفر'] || '').toString().trim().toUpperCase();
  const pkg      = (p['اسم الباقة'] || '').toString().trim();
  if (passport && pkg) pkgMap[passport] = pkg;
});

// إحصاء الحجاج لكل مرشد وباقة
const stats = {};
guides.forEach(g => {
  const guide    = (g['Guide Name'] || '').toString().trim();
  const passport = (g['Passport Number'] || '').toString().trim().toUpperCase();
  const pkg      = pkgMap[passport] || (g['Package Name'] || '').toString().trim();
  if (!guide) return;
  if (!stats[guide]) stats[guide] = { total: 0, packages: {} };
  stats[guide].total++;
  if (pkg) stats[guide].packages[pkg] = (stats[guide].packages[pkg] || 0) + 1;
});

// ترتيب حسب العدد
const sorted = Object.entries(stats).sort((a, b) => b[1].total - a[1].total).slice(0, 20);
console.log('=== أكثر المرشدين حجاجاً ===\n');
sorted.forEach(([guide, data], idx) => {
  console.log((idx+1) + '. ' + guide + ' → ' + data.total + ' حاج');
  const topPkgs = Object.entries(data.packages).sort((a, b) => b[1] - a[1]).slice(0, 3);
  topPkgs.forEach(([pkg, count]) => console.log('     • ' + pkg + ': ' + count));
});

console.log('\nإجمالي المرشدين: ' + Object.keys(stats).length);

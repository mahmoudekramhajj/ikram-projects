// ============================================
// دوال مساعدة — تنسيق التاريخ والوقت والأرقام
// ============================================

function formatDate_(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + pad2_(val.getMonth() + 1) + '-' + pad2_(val.getDate());
  }
  var s = String(val);
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  return s || '-';
}

function formatTime_(val) {
  if (!val) return '-';
  var s = String(val);
  if (s.indexOf('.') !== -1) s = s.split('.')[0];
  var parts = s.split(':');
  if (parts.length >= 2) return parts[0] + ':' + parts[1];
  return s || '-';
}

function pad2_(n) {
  return n < 10 ? '0' + n : '' + n;
}

function parseDate_(dateStr) {
  if (!dateStr) return null;
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return null;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setHours(0, 0, 0, 0);
  return d;
}

function normDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + pad2_(val.getMonth() + 1) + '-' + pad2_(val.getDate());
  }
  var s = String(val);
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  return s;
}

function getTodayString_() {
  var d = new Date();
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}

function getDateOffset_(dateStr, days) {
  var parts = dateStr.split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setDate(d.getDate() + days);
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}

// تنسيق الأرقام بفواصل آلاف
function formatNumber_(n) {
  if (n === null || n === undefined || n === '') return '-';
  return Number(n).toLocaleString('en-US');
}

// حساب نسبة مئوية
function formatPercent_(part, total) {
  if (!total || total === 0) return '0%';
  return Math.round((part / total) * 100) + '%';
}

// اختصار النص الطويل
function truncate_(text, maxLen) {
  if (!text) return '';
  var s = String(text);
  if (s.length <= maxLen) return s;
  return s.substring(0, maxLen - 3) + '...';
}

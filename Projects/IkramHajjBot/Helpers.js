// ============================================
// دوال مساعدة — تنسيق التاريخ والوقت
// ============================================

function formatDate_(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    var d = val;
    return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
  }
  var s = String(val);
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  return s || '-';
}

function formatTime_(val) {
  if (!val) return '-';
  // Date objects (GAS returns time cells as Date with 1899 epoch)
  if (typeof val === 'object' && val.getHours) {
    return pad2_(val.getHours()) + ':' + pad2_(val.getMinutes());
  }
  var s = String(val);
  // Catch stringified Date objects: "Sun Dec 31 1899 07:01:00 GMT+0300"
  if (s.indexOf('1899') !== -1 || s.indexOf('1900') !== -1) {
    var match = s.match(/(d{1,2}):(d{2})/);
    if (match) return pad2_(parseInt(match[1], 10)) + ':' + match[2];
  }
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

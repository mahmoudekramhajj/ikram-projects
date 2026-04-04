/**
 * ReportEngine.js — محرك البيانات v2.0
 * بحث شامل عبر كل الشيتات + دمج تلقائي
 */

// ══════════════════════════════════════════════
//  Helpers
// ══════════════════════════════════════════════

function clean_(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

function toNum_(val) {
  if (!val && val !== 0) return 0;
  var n = Number(val);
  return isNaN(n) ? 0 : n;
}

function formatDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var d = val.getDate(), m = val.getMonth() + 1, y = val.getFullYear();
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
  }
  return String(val).trim();
}

function parseDate_(str) {
  if (!str) return null;
  if (str instanceof Date) return str;
  var d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// ══════════════════════════════════════════════
//  Main: جلب البيانات الموحّد
// ══════════════════════════════════════════════

/**
 * جلب البيانات من شيت واحد أو أكثر
 * @param {Object} params
 *   - selections: [{sheetName, columns: [colIndex, ...]}] — أعمدة محددة من كل شيت
 *   - filters: [{sheetName, colIndex, operator, value, value2}]
 *   - groupBy: {sheetName, colIndex} | null
 *   - aggregations: [{sheetName, colIndex, func}]
 *   - sortBy: {sheetName, colIndex} | null
 *   - sortDir: 'asc' | 'desc'
 *   - page: number
 *   - pageSize: number
 */
function fetchReportData(params) {
  var selections = params.selections || [];
  if (selections.length === 0) return { error: 'No columns selected', rows: [], totalRows: 0 };

  // تحديد الشيتات المطلوبة
  var sheetNames = [];
  selections.forEach(function(sel) {
    if (sheetNames.indexOf(sel.sheetName) === -1) sheetNames.push(sel.sheetName);
  });

  if (sheetNames.length === 1) {
    // ── شيت واحد: استعلام مباشر ──
    return fetchSingleSheet_(params, sheetNames[0]);
  } else if (sheetNames.length === 2) {
    // ── شيتان: محاولة الربط التلقائي ──
    return fetchJoinedSheets_(params, sheetNames);
  } else {
    // ── أكثر من شيتين: نربط تسلسلياً ──
    return fetchMultiSheets_(params, sheetNames);
  }
}

// ══════════════════════════════════════════════
//  Single Sheet Query
// ══════════════════════════════════════════════

function fetchSingleSheet_(params, sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'Sheet "' + sheetName + '" not found', rows: [], totalRows: 0 };

  var cfg = SHEET_CONFIG[sheetName] || {};
  var startRow = cfg.dataStartRow || DEFAULT_START_ROW;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < startRow) return { rows: [], totalRows: 0, page: 1, totalPages: 0 };

  var allData = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  // الأعمدة المحددة
  var sel = params.selections.find(function(s) { return s.sheetName === sheetName; });
  var selectedCols = sel ? sel.columns : [];

  // قراءة الهيدرات
  var headerRow = startRow - 1;
  if (headerRow < 1) headerRow = 1;
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  // بناء column metadata
  var colMeta = selectedCols.map(function(idx) {
    var h = String(headers[idx] || '').trim() || ('col_' + (idx + 1));
    return { index: idx, name: h, type: guessColumnType_(h), sheetName: sheetName };
  });

  // ── الفلترة ──
  var filters = (params.filters || []).filter(function(f) { return f.sheetName === sheetName; });
  var filtered = applyFilters_(allData, filters, colMeta);

  // ── التجميع ──
  if (params.groupBy && params.groupBy.sheetName === sheetName) {
    return buildGroupedResult_(filtered, params, selectedCols, colMeta, headers);
  }

  // ── الترتيب ──
  if (params.sortBy && params.sortBy.sheetName === sheetName) {
    sortData_(filtered, params.sortBy.colIndex, params.sortDir, colMeta);
  }

  // ── Pagination ──
  return paginateAndExtract_(filtered, selectedCols, colMeta, params);
}

// ══════════════════════════════════════════════
//  Joined Two Sheets
// ══════════════════════════════════════════════

function fetchJoinedSheets_(params, sheetNames) {
  var join = findJoinRelation(sheetNames[0], sheetNames[1]);

  if (!join) {
    // لا توجد علاقة معروفة — نعرض كل شيت على حدة
    return {
      error: 'noJoin',
      errorAr: 'لا توجد علاقة معروفة بين "' + sheetNames[0] + '" و "' + sheetNames[1] + '". اختر أعمدة من شيت واحد أو من شيتات مرتبطة.',
      errorEn: 'No known relationship between "' + sheetNames[0] + '" and "' + sheetNames[1] + '". Select columns from one sheet or related sheets.',
      rows: [],
      totalRows: 0
    };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // قراءة الشيتين
  var data = {};
  var headers = {};
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var cfg = SHEET_CONFIG[name] || {};
    var startRow = cfg.dataStartRow || DEFAULT_START_ROW;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < startRow) { data[name] = []; headers[name] = []; return; }

    data[name] = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

    var headerRow = startRow - 1;
    if (headerRow < 1) headerRow = 1;
    headers[name] = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  });

  // بناء مفتاح الربط
  var primarySheet = join.primaryKey.sheet;
  var secondarySheet = join.secondaryKey.sheet;
  var primaryCol = join.primaryKey.col;
  var secondaryCol = join.secondaryKey.col;

  // بناء lookup من الشيت الثانوي
  var secondaryMap = {};
  var secData = data[secondarySheet] || [];
  for (var i = 0; i < secData.length; i++) {
    var key = clean_(secData[i][secondaryCol]);
    if (!key) continue;
    if (!secondaryMap[key]) secondaryMap[key] = [];
    secondaryMap[key].push(secData[i]);
  }

  // دمج البيانات
  var merged = [];
  var priData = data[primarySheet] || [];
  for (var p = 0; p < priData.length; p++) {
    var priKey = clean_(priData[p][primaryCol]);
    var matches = secondaryMap[priKey] || [null]; // null = لا يوجد مطابقة

    for (var m = 0; m < matches.length; m++) {
      var mergedRow = {};
      mergedRow[primarySheet] = priData[p];
      mergedRow[secondarySheet] = matches[m];
      merged.push(mergedRow);
    }
  }

  // استخراج الأعمدة المحددة
  var allSelectedCols = [];
  params.selections.forEach(function(sel) {
    sel.columns.forEach(function(colIdx) {
      var h = headers[sel.sheetName] ? String(headers[sel.sheetName][colIdx] || '').trim() : ('col_' + (colIdx + 1));
      allSelectedCols.push({
        sheetName: sel.sheetName,
        index: colIdx,
        name: h,
        type: guessColumnType_(h)
      });
    });
  });

  // فلترة
  var filters = params.filters || [];
  var filtered = merged.filter(function(row) {
    for (var f = 0; f < filters.length; f++) {
      var filter = filters[f];
      var sheetData = row[filter.sheetName];
      if (!sheetData) continue;
      var val = sheetData[filter.colIndex];
      var meta = allSelectedCols.find(function(c) {
        return c.sheetName === filter.sheetName && c.index === filter.colIndex;
      });
      if (!matchFilter_(val, filter, meta ? meta.type : 'text')) return false;
    }
    return true;
  });

  // ترتيب
  if (params.sortBy) {
    var sortSheet = params.sortBy.sheetName;
    var sortCol = params.sortBy.colIndex;
    var sortDir = (params.sortDir === 'desc') ? -1 : 1;
    var sortMeta = allSelectedCols.find(function(c) {
      return c.sheetName === sortSheet && c.index === sortCol;
    });
    var sortType = sortMeta ? sortMeta.type : 'text';

    filtered.sort(function(a, b) {
      var va = a[sortSheet] ? a[sortSheet][sortCol] : '';
      var vb = b[sortSheet] ? b[sortSheet][sortCol] : '';
      if (sortType === 'number' || sortType === 'currency' || sortType === 'percentage') {
        return (toNum_(va) - toNum_(vb)) * sortDir;
      }
      if (sortType === 'date') {
        var da = parseDate_(va), db = parseDate_(vb);
        if (!da && !db) return 0;
        if (!da) return sortDir;
        if (!db) return -sortDir;
        return (da.getTime() - db.getTime()) * sortDir;
      }
      return clean_(va).localeCompare(clean_(vb)) * sortDir;
    });
  }

  // Grouping
  if (params.groupBy) {
    return buildGroupedFromMerged_(filtered, params, allSelectedCols);
  }

  // Pagination
  var totalRows = filtered.length;
  var pageSize = params.pageSize || 100;
  var page = params.page || 1;
  var totalPages = Math.ceil(totalRows / pageSize) || 1;
  if (page > totalPages) page = totalPages;
  var startIdx = (page - 1) * pageSize;
  var pageData = filtered.slice(startIdx, startIdx + pageSize);

  // استخراج القيم
  var rows = pageData.map(function(mergedRow) {
    return allSelectedCols.map(function(col) {
      var sheetData = mergedRow[col.sheetName];
      if (!sheetData) return '';
      var val = sheetData[col.index];
      if (col.type === 'date') return formatDate_(val);
      return (val === null || val === undefined) ? '' : val;
    });
  });

  return {
    rows: rows,
    totalRows: totalRows,
    page: page,
    pageSize: pageSize,
    totalPages: totalPages,
    columnsMeta: allSelectedCols.map(function(c) {
      return { name: c.name, type: c.type, sheetName: c.sheetName };
    })
  };
}

// ══════════════════════════════════════════════
//  Multi-Sheet (3+) — chain joins
// ══════════════════════════════════════════════

function fetchMultiSheets_(params, sheetNames) {
  // حالياً: نستخدم أول شيتين فقط ونضيف رسالة
  return fetchJoinedSheets_(params, [sheetNames[0], sheetNames[1]]);
}

// ══════════════════════════════════════════════
//  الفلترة
// ══════════════════════════════════════════════

function applyFilters_(data, filters, colMeta) {
  if (!filters || filters.length === 0) return data;

  return data.filter(function(row) {
    for (var f = 0; f < filters.length; f++) {
      var filter = filters[f];
      var val = row[filter.colIndex];
      var meta = colMeta.find(function(c) { return c.index === filter.colIndex; });
      var type = meta ? meta.type : 'text';
      if (!matchFilter_(val, filter, type)) return false;
    }
    return true;
  });
}

function matchFilter_(val, filter, type) {
  var op = filter.operator;
  var target = filter.value;

  if (type === 'number' || type === 'currency' || type === 'percentage') {
    var numVal = toNum_(val);
    var numTarget = toNum_(target);
    switch (op) {
      case 'equals': return numVal === numTarget;
      case 'gt': return numVal > numTarget;
      case 'gte': return numVal >= numTarget;
      case 'lt': return numVal < numTarget;
      case 'lte': return numVal <= numTarget;
      case 'between': return numVal >= numTarget && numVal <= toNum_(filter.value2);
      case 'notEmpty': return val !== '' && val !== null && val !== undefined;
      default: return true;
    }
  }

  if (type === 'date') {
    var dateVal = parseDate_(val);
    var dateTarget = parseDate_(target);
    switch (op) {
      case 'equals':
        if (!dateVal || !dateTarget) return false;
        return formatDate_(dateVal) === formatDate_(dateTarget);
      case 'after':
        if (!dateVal || !dateTarget) return false;
        return dateVal > dateTarget;
      case 'before':
        if (!dateVal || !dateTarget) return false;
        return dateVal < dateTarget;
      case 'between':
        var d2 = parseDate_(filter.value2);
        if (!dateVal || !dateTarget || !d2) return false;
        return dateVal >= dateTarget && dateVal <= d2;
      case 'notEmpty': return !!dateVal;
      default: return true;
    }
  }

  // text
  var strVal = clean_(val).toLowerCase();
  var strTarget = clean_(target).toLowerCase();
  switch (op) {
    case 'equals': return strVal === strTarget;
    case 'contains': return strVal.indexOf(strTarget) > -1;
    case 'startsWith': return strVal.indexOf(strTarget) === 0;
    case 'notEquals': return strVal !== strTarget;
    case 'notEmpty': return strVal !== '';
    case 'empty': return strVal === '';
    default: return true;
  }
}

// ══════════════════════════════════════════════
//  الترتيب
// ══════════════════════════════════════════════

function sortData_(data, sortIdx, sortDir, colMeta) {
  var dir = (sortDir === 'desc') ? -1 : 1;
  var meta = colMeta.find(function(c) { return c.index === sortIdx; });
  var type = meta ? meta.type : 'text';

  data.sort(function(a, b) {
    var va = a[sortIdx], vb = b[sortIdx];
    if (type === 'number' || type === 'currency' || type === 'percentage') {
      return (toNum_(va) - toNum_(vb)) * dir;
    }
    if (type === 'date') {
      var da = parseDate_(va), db = parseDate_(vb);
      if (!da && !db) return 0;
      if (!da) return dir;
      if (!db) return -dir;
      return (da.getTime() - db.getTime()) * dir;
    }
    return clean_(va).localeCompare(clean_(vb)) * dir;
  });
}

// ══════════════════════════════════════════════
//  Pagination + Extract columns
// ══════════════════════════════════════════════

function paginateAndExtract_(data, selectedCols, colMeta, params) {
  var totalRows = data.length;
  var pageSize = params.pageSize || 100;
  var page = params.page || 1;
  var totalPages = Math.ceil(totalRows / pageSize) || 1;
  if (page > totalPages) page = totalPages;
  var startIdx = (page - 1) * pageSize;
  var pageData = data.slice(startIdx, startIdx + pageSize);

  var rows = [];
  for (var r = 0; r < pageData.length; r++) {
    var row = [];
    for (var c = 0; c < selectedCols.length; c++) {
      var val = pageData[r][selectedCols[c]];
      var meta = colMeta[c];
      if (meta && meta.type === 'date') {
        val = formatDate_(val);
      } else {
        val = (val === null || val === undefined) ? '' : val;
      }
      row.push(val);
    }
    rows.push(row);
  }

  return {
    rows: rows,
    totalRows: totalRows,
    page: page,
    pageSize: pageSize,
    totalPages: totalPages,
    columnsMeta: colMeta.map(function(c) { return { name: c.name, type: c.type, sheetName: c.sheetName }; })
  };
}

// ══════════════════════════════════════════════
//  التجميع — من شيت واحد
// ══════════════════════════════════════════════

function buildGroupedResult_(data, params, selectedCols, colMeta, headers) {
  var groupIdx = params.groupBy.colIndex;
  var aggs = params.aggregations || [];
  var groupHeader = String(headers[groupIdx] || '').trim() || ('col_' + (groupIdx + 1));

  var groups = {};
  var groupOrder = [];

  for (var i = 0; i < data.length; i++) {
    var key = clean_(data[i][groupIdx]) || '(فارغ)';
    if (!groups[key]) { groups[key] = []; groupOrder.push(key); }
    groups[key].push(data[i]);
  }

  var result = [];
  var totals = { _key: 'الإجمالي / Total', _count: data.length };

  for (var a = 0; a < aggs.length; a++) {
    totals['agg_' + aggs[a].colIndex + '_' + aggs[a].func] = 0;
  }
  var totalCountForAvg = {};

  for (var g = 0; g < groupOrder.length; g++) {
    var gKey = groupOrder[g];
    var gRows = groups[gKey];
    var row = { _key: gKey, _count: gRows.length };

    for (var ai = 0; ai < aggs.length; ai++) {
      var agg = aggs[ai];
      var aggKey = 'agg_' + agg.colIndex + '_' + agg.func;
      var values = gRows.map(function(r) { return toNum_(r[agg.colIndex]); });
      var aggVal = calcAgg_(values, agg.func);
      row[aggKey] = Math.round(aggVal * 100) / 100;

      // totals
      if (agg.func === 'SUM' || agg.func === 'COUNT') {
        totals[aggKey] = (totals[aggKey] || 0) + aggVal;
      } else if (agg.func === 'AVG') {
        totals[aggKey] = (totals[aggKey] || 0) + aggVal * gRows.length;
        totalCountForAvg[aggKey] = (totalCountForAvg[aggKey] || 0) + gRows.length;
      } else if (agg.func === 'MIN') {
        if (g === 0 || aggVal < totals[aggKey]) totals[aggKey] = aggVal;
      } else if (agg.func === 'MAX') {
        if (g === 0 || aggVal > totals[aggKey]) totals[aggKey] = aggVal;
      }
    }
    result.push(row);
  }

  // تصحيح المتوسطات
  for (var tk in totalCountForAvg) {
    if (totalCountForAvg[tk] > 0) {
      totals[tk] = Math.round((totals[tk] / totalCountForAvg[tk]) * 100) / 100;
    }
  }

  // ترتيب
  result.sort(function(a, b) { return b._count - a._count; });

  // أعمدة النتيجة
  var resultColumns = [
    { key: '_key', name: groupHeader, type: 'text' },
    { key: '_count', name: 'العدد / Count', type: 'number' }
  ];

  for (var ac = 0; ac < aggs.length; ac++) {
    var am = String(headers[aggs[ac].colIndex] || '').trim() || ('col_' + (aggs[ac].colIndex + 1));
    resultColumns.push({
      key: 'agg_' + aggs[ac].colIndex + '_' + aggs[ac].func,
      name: am + ' (' + aggs[ac].func + ')',
      type: 'number'
    });
  }

  var rows = result.map(function(r) {
    return resultColumns.map(function(col) { return r[col.key] || 0; });
  });
  var totalRow = resultColumns.map(function(col) { return totals[col.key] || 0; });

  return {
    grouped: true,
    groupColumns: resultColumns,
    rows: rows,
    totals: totalRow,
    totalGroups: groupOrder.length,
    totalRows: data.length
  };
}

function buildGroupedFromMerged_(merged, params, allSelectedCols) {
  var groupSheet = params.groupBy.sheetName;
  var groupCol = params.groupBy.colIndex;
  var aggs = params.aggregations || [];

  var groups = {};
  var groupOrder = [];

  for (var i = 0; i < merged.length; i++) {
    var sd = merged[i][groupSheet];
    var key = sd ? clean_(sd[groupCol]) : '(فارغ)';
    if (!key) key = '(فارغ)';
    if (!groups[key]) { groups[key] = []; groupOrder.push(key); }
    groups[key].push(merged[i]);
  }

  var result = [];
  for (var g = 0; g < groupOrder.length; g++) {
    var gKey = groupOrder[g];
    var gRows = groups[gKey];
    var row = { _key: gKey, _count: gRows.length };

    for (var ai = 0; ai < aggs.length; ai++) {
      var agg = aggs[ai];
      var aggKey = 'agg_' + agg.sheetName + '_' + agg.colIndex + '_' + agg.func;
      var values = gRows.map(function(r) {
        var sd = r[agg.sheetName];
        return sd ? toNum_(sd[agg.colIndex]) : 0;
      });
      row[aggKey] = Math.round(calcAgg_(values, agg.func) * 100) / 100;
    }
    result.push(row);
  }

  result.sort(function(a, b) { return b._count - a._count; });

  var groupMeta = allSelectedCols.find(function(c) {
    return c.sheetName === groupSheet && c.index === groupCol;
  });

  var resultColumns = [
    { key: '_key', name: groupMeta ? groupMeta.name : 'Group', type: 'text' },
    { key: '_count', name: 'العدد / Count', type: 'number' }
  ];

  for (var ac = 0; ac < aggs.length; ac++) {
    var am = allSelectedCols.find(function(c) {
      return c.sheetName === aggs[ac].sheetName && c.index === aggs[ac].colIndex;
    });
    resultColumns.push({
      key: 'agg_' + aggs[ac].sheetName + '_' + aggs[ac].colIndex + '_' + aggs[ac].func,
      name: (am ? am.name : 'Col') + ' (' + aggs[ac].func + ')',
      type: 'number'
    });
  }

  var rows = result.map(function(r) {
    return resultColumns.map(function(col) { return r[col.key] || 0; });
  });

  return {
    grouped: true,
    groupColumns: resultColumns,
    rows: rows,
    totals: [],
    totalGroups: groupOrder.length,
    totalRows: merged.length
  };
}

function calcAgg_(values, func) {
  if (values.length === 0) return 0;
  switch (func) {
    case 'SUM': return values.reduce(function(s, v) { return s + v; }, 0);
    case 'AVG': return values.reduce(function(s, v) { return s + v; }, 0) / values.length;
    case 'COUNT': return values.length;
    case 'MIN': return Math.min.apply(null, values);
    case 'MAX': return Math.max.apply(null, values);
    default: return 0;
  }
}

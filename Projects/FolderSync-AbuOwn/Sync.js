/**
 * Sync.js — مزامنة المجلد الأصلي (TICKETS) → مجلد ابوعون (Hajj TICKETS)
 *
 * المنطق:
 *   - يعمل على مستوى المجلدات الفرعية (subfolder level)
 *   - مجلد موجود في الأصلي + موجود في ابوعون  → يزامن الملفات الداخلية
 *   - مجلد موجود في الأصلي + غير موجود        → ينشئه وينسخ محتواه
 *   - مجلدات ابوعون بدون مقابل               → تُترك كما هي
 *   - ملف أُضيف في الأصلي                    → يُنسخ
 *   - ملف عُدِّل في الأصلي                   → تُحذف النسخة القديمة وتُنشأ جديدة
 *   - ملف حُذف من الأصلي                     → تُنقل نسخته إلى السلة (فقط إذا كانت منسوخة منه)
 *   - ملفات ابوعون غير المنسوخة من الأصلي   → لا تُمس
 *
 * Trigger: كل 5 دقائق (time-based)
 */

var ORIGINAL_FOLDER_ID = '1XkkGiaNjU58V11i7aL9-4I9LcAjHRw82';
var ABUOWN_FOLDER_ID   = '154uCHRSpZnC5iGvdL1uGqVroy6txUIvG';

// ─────────────────────────────────────────────────────────────
// الدوال العامة
// ─────────────────────────────────────────────────────────────

/**
 * setup — يُشغَّل مرة واحدة:
 *   1. ينصّب trigger كل 5 دقائق.
 *   2. يزامن الحالة الحالية (initialSync).
 */
function setup() {
  Logger.log('=== FolderSync Setup ===');
  installTrigger_();
  var result = initialSync_();
  Logger.log('تم: ' + JSON.stringify(result));
  return result;
}

/**
 * syncCheck — يُستدعى بواسطة الـ trigger كل 5 دقائق.
 * يمكن تشغيله يدوياً من المحرر أو عبر ClaudeAPI.
 */
function syncCheck() {
  Logger.log('=== syncCheck بدء ===');
  var props  = PropertiesService.getScriptProperties();
  var result = { folders: 0, added: 0, updated: 0, removed: 0, errors: 0 };

  var originalFolder, abuOwnFolder;
  for (var rootTry = 0; rootTry < 4; rootTry++) {
    try {
      originalFolder = DriveApp.getFolderById(ORIGINAL_FOLDER_ID);
      abuOwnFolder   = DriveApp.getFolderById(ABUOWN_FOLDER_ID);
      break;
    } catch(e) {
      if (rootTry < 3) {
        Logger.log('خطأ مؤقت في الوصول، محاولة ' + (rootTry+2) + '...');
        Utilities.sleep(2000 * (rootTry + 1));
      } else {
        Logger.log('خطأ الوصول للمجلدات الرئيسية بعد 4 محاولات: ' + e.message);
        return result;
      }
    }
  }

  // فهرس مجلدات ابوعون (الاسم → مجلد)
  var abuIndex = buildFolderIndex_(abuOwnFolder);

  // تزامن كل مجلد فرعي في الأصلي
  var origSubs2;
  try { origSubs2 = originalFolder.getFolders(); } catch(e) {
    Logger.log('خطأ getFolders: ' + e.message); return result;
  }

  while (true) {
    var has2;
    try { has2 = origSubs2.hasNext(); } catch(e) { break; }
    if (!has2) break;

    var origSub2, subName2;
    try {
      origSub2 = origSubs2.next();
      subName2 = origSub2.getName();
    } catch(e) {
      Logger.log('خطأ قراءة مجلد: ' + e.message + ' — تخطّي');
      result.errors++;
      continue;
    }

    try {
      Utilities.sleep(500); // تجنّب rate limiting
      var targetFolder2 = resolveTargetFolder_(origSub2, abuOwnFolder, abuIndex, props);
      if (!targetFolder2) { result.errors++; continue; }

      var subResult = syncFilesInFolder_(origSub2, targetFolder2, props);
      result.folders++;
      result.added   += subResult.added;
      result.updated += subResult.updated;
      result.removed += subResult.removed;
      result.errors  += subResult.errors;
    } catch(e) {
      Logger.log('خطأ معالجة مجلد "' + subName2 + '": ' + e.message + ' — تخطّي');
      result.errors++;
    }
  }

  Logger.log('syncCheck انتهى: ' + JSON.stringify(result));
  return result;
}

/**
 * status — تقرير سريع.
 */
function status() {
  var originalFolder = DriveApp.getFolderById(ORIGINAL_FOLDER_ID);
  var abuOwnFolder   = DriveApp.getFolderById(ABUOWN_FOLDER_ID);
  var props          = PropertiesService.getScriptProperties().getProperties();

  var fileMappings   = Object.keys(props).filter(function(k) { return k.indexOf('map_') === 0; }).length;
  var folderMappings = Object.keys(props).filter(function(k) { return k.indexOf('foldermap_') === 0; }).length;

  var triggers = ScriptApp.getProjectTriggers().map(function(t) {
    return t.getHandlerFunction() + ' / ' + t.getTriggerSource();
  });

  Logger.log('ملفات مربوطة: ' + fileMappings);
  Logger.log('مجلدات مربوطة: ' + folderMappings);
  Logger.log('triggers: ' + triggers.join(', '));

  return { fileMappings: fileMappings, folderMappings: folderMappings, triggers: triggers };
}

/**
 * clearAllMappings — يحذف جميع سجلات الربط (لإعادة بناء كاملة).
 */
function clearAllMappings() {
  var props   = PropertiesService.getScriptProperties();
  var all     = props.getProperties();
  var deleted = 0;
  Object.keys(all).forEach(function(k) {
    if (k.indexOf('map_') === 0 || k.indexOf('mod_') === 0 || k.indexOf('foldermap_') === 0) {
      props.deleteProperty(k);
      deleted++;
    }
  });
  Logger.log('حُذف ' + deleted + ' سجل');
  return { deleted: deleted };
}

// ─────────────────────────────────────────────────────────────
// initialSync — المزامنة الأولية
// ─────────────────────────────────────────────────────────────

function initialSync_() {
  var props  = PropertiesService.getScriptProperties();
  var result = { folders: 0, linked: 0, copied: 0, errors: 0 };

  var originalFolder, abuOwnFolder;
  try {
    originalFolder = DriveApp.getFolderById(ORIGINAL_FOLDER_ID);
    abuOwnFolder   = DriveApp.getFolderById(ABUOWN_FOLDER_ID);
  } catch(e) {
    Logger.log('خطأ الوصول: ' + e.message);
    return result;
  }

  // فهرس مجلدات ابوعون (الاسم → مجلد)
  var abuIndex = buildFolderIndex_(abuOwnFolder);

  var origSubs;
  try { origSubs = originalFolder.getFolders(); } catch(e) {
    Logger.log('خطأ getFolders الأصلي: ' + e.message); return result;
  }

  while (true) {
    var hasNext;
    try { hasNext = origSubs.hasNext(); } catch(e) { break; }
    if (!hasNext) break;

    var origSub, subName;
    try {
      origSub = origSubs.next();
      subName = origSub.getName();
    } catch(e) {
      Logger.log('خطأ قراءة مجلد فرعي: ' + e.message + ' — تخطّي');
      result.errors++;
      continue;
    }

    try {
      var targetFolder = resolveTargetFolder_(origSub, abuOwnFolder, abuIndex, props);
      if (!targetFolder) { result.errors++; continue; }

      var existingFiles = buildFileIndex_(targetFolder);

      var origFiles = origSub.getFiles();
      while (origFiles.hasNext()) {
        var file    = origFiles.next();
        var fileId  = file.getId();
        var modTime = file.getLastUpdated().getTime().toString();
        var name    = file.getName();

        if (props.getProperty('map_' + fileId)) continue;

        if (existingFiles[name]) {
          props.setProperty('map_' + fileId, existingFiles[name]);
          props.setProperty('mod_' + fileId, modTime);
          Logger.log('ربط: ' + name);
          result.linked++;
        } else {
          try {
            var copy = file.makeCopy(name, targetFolder);
            props.setProperty('map_' + fileId, copy.getId());
            props.setProperty('mod_' + fileId, modTime);
            Logger.log('نُسخ: ' + name + ' ← ' + subName);
            result.copied++;
          } catch(ec) {
            Logger.log('خطأ نسخ ' + name + ': ' + ec.message);
            result.errors++;
          }
        }
      }
      result.folders++;
    } catch(e) {
      Logger.log('خطأ معالجة مجلد "' + subName + '": ' + e.message + ' — تخطّي');
      result.errors++;
    }
  }

  Logger.log('initialSync: ' + JSON.stringify(result));
  return result;
}

// ─────────────────────────────────────────────────────────────
// مزامنة ملفات مجلد واحد
// ─────────────────────────────────────────────────────────────

function syncFilesInFolder_(origFolder, targetFolder, props) {
  var result = { added: 0, updated: 0, removed: 0, errors: 0 };

  // retry حتى 3 مرات إذا فشل getFiles
  var files;
  for (var attempt = 0; attempt < 3; attempt++) {
    try { files = origFolder.getFiles(); break; }
    catch(e) {
      if (attempt < 2) { Utilities.sleep(1000 * (attempt + 1)); }
      else { throw e; }
    }
  }
  while (files.hasNext()) {
    var file    = files.next();
    var fileId  = file.getId();
    var modTime = file.getLastUpdated().getTime().toString();
    var name    = file.getName();

    var trackedModTime = props.getProperty('mod_' + fileId);
    var copyId         = props.getProperty('map_' + fileId);

    if (!copyId) {
      // ملف جديد
      try {
        var copy = file.makeCopy(name, targetFolder);
        props.setProperty('map_' + fileId, copy.getId());
        props.setProperty('mod_' + fileId, modTime);
        Logger.log('مُضاف: ' + name);
        result.added++;
      } catch(e) {
        Logger.log('خطأ إضافة ' + name + ': ' + e.message);
        result.errors++;
      }
    } else if (trackedModTime !== modTime) {
      // ملف معدَّل — استبدل النسخة
      try { DriveApp.getFileById(copyId).setTrashed(true); } catch(e) {}
      try {
        var newCopy = file.makeCopy(name, targetFolder);
        props.setProperty('map_' + fileId, newCopy.getId());
        props.setProperty('mod_' + fileId, modTime);
        Logger.log('مُحدَّث: ' + name);
        result.updated++;
      } catch(e) {
        Logger.log('خطأ تحديث ' + name + ': ' + e.message);
        result.errors++;
      }
    }
  }

  return result;
}

// ─────────────────────────────────────────────────────────────
// مساعدات
// ─────────────────────────────────────────────────────────────

/**
 * يحدد المجلد المقابل في ابوعون لمجلد فرعي من الأصلي.
 * يبحث أولاً في المسجَّل، ثم بمطابقة الاسم، ثم ينشئ مجلداً جديداً.
 */
function resolveTargetFolder_(origSub, abuOwnFolder, abuIndex, props) {
  var origId   = origSub.getId();
  var origName = origSub.getName();

  // 1. مسجَّل مسبقاً — retry حتى 3 مرات
  var mappedId = props.getProperty('foldermap_' + origId);
  if (mappedId) {
    for (var t = 0; t < 3; t++) {
      try { return DriveApp.getFolderById(mappedId); }
      catch(e) { Utilities.sleep(1000 * (t + 1)); }
    }
    Logger.log('تحذير: فشل getFolderById للـ mapping المحفوظ لـ ' + origName + ' — إعادة البحث');
  }

  // 2. مطابقة بالاسم من الفهرس
  var matched = null;
  Object.keys(abuIndex).forEach(function(abuName) {
    if (matched) return;
    if (origName.indexOf(abuName) !== -1 || abuName.indexOf(origName) !== -1) {
      matched = abuIndex[abuName];
    }
  });

  if (matched) {
    props.setProperty('foldermap_' + origId, matched.getId());
    Logger.log('مجلد مطابق: "' + origName + '" ↔ "' + matched.getName() + '"');
    return matched;
  }

  // 3. إنشاء مجلد جديد — retry حتى 3 مرات
  for (var c = 0; c < 3; c++) {
    try {
      var newFolder = abuOwnFolder.createFolder(origName);
      props.setProperty('foldermap_' + origId, newFolder.getId());
      Logger.log('مجلد جديد: ' + origName);
      return newFolder;
    } catch(e) {
      if (c < 2) { Utilities.sleep(1000 * (c + 1)); }
      else { Logger.log('خطأ إنشاء مجلد ' + origName + ': ' + e.message); }
    }
  }
  return null;
}

/** بناء فهرس (الاسم → مجلد) لمجلد معيّن — retry حتى 3 مرات */
function buildFolderIndex_(folder) {
  for (var t = 0; t < 3; t++) {
    var index = {};
    try {
      var subs = folder.getFolders();
      while (subs.hasNext()) {
        var sf = subs.next();
        index[sf.getName()] = sf;
      }
      return index; // نجح
    } catch(e) {
      if (t < 2) {
        Logger.log('خطأ buildFolderIndex (محاولة ' + (t+1) + '): ' + e.message + ' — إعادة...');
        Utilities.sleep(1500 * (t + 1));
      } else {
        Logger.log('خطأ buildFolderIndex نهائي: ' + e.message);
      }
    }
  }
  return {};
}

/** بناء فهرس (الاسم → ID) للملفات في مجلد معيّن */
function buildFileIndex_(folder) {
  var index = {};
  try {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      index[f.getName()] = f.getId();
    }
  } catch(e) {
    Logger.log('خطأ buildFileIndex: ' + e.message);
  }
  return index;
}

function diagnoseProblemFolders() {
  var originalFolder = DriveApp.getFolderById(ORIGINAL_FOLDER_ID);
  var subs = originalFolder.getFolders();

  var problems = ['UWR2PL-VRB58V-SAP8XW', 'U48FI2', 'RAWKK4', 'TZIBG8',
                  'S7HT4Z', 'UNALHB', 'UX7SSE'];

  while (subs.hasNext()) {
    var sub;
    try { sub = subs.next(); } catch(e) { continue; }

    var name;
    try { name = sub.getName(); } catch(e) { Logger.log('فشل getName: ' + e.message); continue; }

    // هل هذا المجلد من المجلدات الإشكالية؟
    var isProblematic = problems.some(function(p) { return name.indexOf(p) !== -1; });
    if (!isProblematic) continue;

    Logger.log('══ ' + name + ' ══');
    Logger.log('ID: ' + sub.getId());

    // هل هو shortcut؟
    try {
      var mimeType = DriveApp.getFileById(sub.getId()).getMimeType();
      Logger.log('MimeType: ' + mimeType);
    } catch(e) { Logger.log('getMimeType خطأ: ' + e.message); }

    // محاولة قراءة الملفات
    try {
      var files = sub.getFiles();
      var count = 0;
      while (files.hasNext()) { files.next(); count++; }
      Logger.log('ملفات: ' + count);
    } catch(e) { Logger.log('getFiles خطأ: ' + e.message); }

    // محاولة قراءة المجلدات الفرعية
    try {
      var sf = sub.getFolders();
      var scount = 0;
      while (sf.hasNext()) { sf.next(); scount++; }
      Logger.log('مجلدات فرعية: ' + scount);
    } catch(e) { Logger.log('getFolders خطأ: ' + e.message); }
  }

  Logger.log('انتهى الفحص.');
}

function installTrigger_() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'syncCheck') {
      ScriptApp.deleteTrigger(existing[i]);
      Logger.log('حُذف trigger قديم.');
    }
  }
  ScriptApp.newTrigger('syncCheck')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('تم إنشاء trigger كل 5 دقائق.');
}

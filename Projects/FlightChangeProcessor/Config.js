/**
 * Config.js — إعدادات معالج تغييرات الطيران
 */

var CONFIG = {
  // === Spreadsheet الرئيسي (Ikram Abuown — للقراءة فقط) ===
  MAIN_SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',

  // === Spreadsheet التغييرات (منفصل — يُنشأ تلقائياً) ===
  CHANGES_SPREADSHEET_ID: '',
  CHANGES_SHEET_NAME: 'تغييرات الطيران',

  // === Spreadsheet المقارنة (منفصل — يُنشأ تلقائياً) ===
  COMPARISON_SPREADSHEET_ID: '',
  COMPARISON_SHEET_NAME: 'مقارنة الأسماء',

  // === أسماء الشيتات في Spreadsheet الرئيسي ===
  PD_SHEET: 'Presonal Details',
  FLIGHT_SHEET: 'الطيران',

  // === أعمدة Presonal Details (0-based) ===
  PD: {
    SERIAL: 0,
    GROUP_NUM: 1,
    PILGRIM_TYPE: 2,
    CATEGORY: 3,
    GENDER: 4,
    PASSPORT: 5,
    FIRST_NAME_AR: 8,
    LAST_NAME_AR: 9,
    FIRST_NAME_EN: 10,
    LAST_NAME_EN: 11,
    EMAIL: 13,
    PHONE: 14,
    GUIDE: 15,
    PKG_NUM: 18,
    PKG_NAME: 19,
    FLIGHT_TYPE: 20,
    CONTRACT_NAME: 21,
    TICKET_NUM: 24,
    TICKET_LINK: 25
  },

  // === Google Drive ===
  PARENT_FOLDER_ID: '1hKYyDB1hrW6ZiW6ERSAk0Ujk743wx8HV',
  CHANGES_FOLDER_NAME: 'تغييرات',

  // === Gmail ===
  GMAIL_LABEL: 'TKT',
  PROCESSED_LABEL: 'TKT-Processed',

  // === حدود ===
  MAX_RUNTIME_MS: 5.5 * 60 * 1000,
  MAX_EMAILS_PER_RUN: 500
};

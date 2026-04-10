/**
 * Config.js — إعدادات معالج التذاكر
 */

var CONFIG = {
  // === Google Sheets ===
  SPREADSHEET_ID: '1vb5MsSRg6HcZ5DTFAFv5ySRBPW7fCjwtQTNhm9lOZDo',
  SHEET_NAME: 'Copy of Presonal Details',

  // === أرقام الأعمدة (1-based) ===
  COL_FIRST_NAME: 11,   // K — الاسم الأول (إنجليزي)
  COL_LAST_NAME: 12,    // L — اسم العائلة (إنجليزي)
  COL_CONTRACT: 22,     // V — اسم العقد (يحتوي Booking Ref)
  COL_TICKET_LINK: 26,  // Z — رابط التذكرة
  COL_BOOKING_REF: 34,  // AH — Booking Ref
  COL_TICKET_NUM: 35,   // AI — Ticket Number

  // === Google Drive ===
  TICKETS_FOLDER_ID: '1hKYyDB1hrW6ZiW6ERSAk0Ujk743wx8HV',  // مجلد TKT Ekram
  PROCESSED_FOLDER_NAME: 'تمت المعالجة'  // مجلد فرعي للملفات المعالجة
};

const ITEM_SHEET_NAME = 'รายการของหาย';
const ITEM_HEADERS = ['รหัสรายการ', 'ประเภท', 'ชื่อสิ่งของ', 'หมวดหมู่', 'สถานที่', 'วันที่', 'ช่องทางติดต่อ', 'สถานะ'];
const USER_SHEET_NAME = 'บัญชีผู้ใช้';
const USER_HEADERS = ['อีเมล Google', 'ชื่อ-นามสกุล', 'เบอร์ติดต่อ', 'คณะ/หน่วยงาน', 'สถานะผู้ใช้', 'รหัสนักศึกษา/บุคลากร'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบแจ้งของหายและติดตามของหายในมหาวิทยาลัย');
}

function getAccountState() {
  const email = getActiveUserEmail_();
  if (!email) return { email: '', profile: null };
  return {
    email: email,
    profile: findUserProfile_(email)
  };
}

function getAccountLinks() {
  const appUrl = ScriptApp.getService().getUrl();
  if (!appUrl) throw new Error('กรุณา Deploy Apps Script เป็น Web App ก่อน');
  return {
    logout: 'https://accounts.google.com/Logout?continue=' + encodeURIComponent(appUrl),
    switchAccount: 'https://accounts.google.com/AccountChooser?continue=' + encodeURIComponent(appUrl)
  };
}

function registerAccount(profile) {
  const email = getActiveUserEmail_();
  if (!email) throw new Error('ตรวจสอบบัญชี Google ไม่ได้ กรุณาเปิด Web App ด้วยบัญชี Google');
  if (!profile || typeof profile !== 'object') throw new Error('ไม่พบข้อมูลลงทะเบียน');

  const user = {
    email: email,
    name: requiredText_(profile.name, 'ชื่อ-นามสกุล'),
    phone: requiredText_(profile.phone, 'เบอร์ติดต่อ'),
    faculty: requiredText_(profile.faculty, 'คณะ/หน่วยงาน'),
    role: requiredText_(profile.role, 'สถานะผู้ใช้'),
    identifier: requiredText_(profile.identifier, 'รหัสนักศึกษา/บุคลากร')
  };
  if (!['นักศึกษา', 'บุคลากร'].includes(user.role)) {
    throw new Error('สถานะผู้ใช้ไม่ถูกต้อง');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getUserSheet_();
    const existing = findUserRow_(sheet, email);
    if (existing) throw new Error('บัญชี Google นี้ลงทะเบียนแล้ว กรุณาเข้าสู่ระบบ');
    sheet.appendRow([
      user.email,
      user.name,
      user.phone,
      user.faculty,
      user.role,
      user.identifier
    ]);
    return { success: true, profile: user };
  } finally {
    lock.releaseLock();
  }
}

function getItems() {
  requireRegisteredUser_();
  const sheet = getItemSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, ITEM_HEADERS.length).getDisplayValues()
    .filter(row => row.some(value => value !== ''))
    .map(row => ({
      id: row[0],
      type: row[1],
      itemName: row[2],
      category: row[3],
      location: row[4],
      date: row[5],
      contact: row[6],
      status: row[7]
    }))
    .reverse();
}

function saveItemData(formData) {
  requireRegisteredUser_();
  if (!formData || typeof formData !== 'object') {
    throw new Error('ไม่พบข้อมูลที่ต้องการบันทึก');
  }

  const item = {
    type: requiredText_(formData.type, 'ประเภทการแจ้ง'),
    itemName: requiredText_(formData.itemName, 'ชื่อสิ่งของ'),
    category: optionalText_(formData.category),
    location: requiredText_(formData.location, 'สถานที่'),
    date: requiredText_(formData.date, 'วันที่เกิดเหตุ'),
    contact: requiredText_(formData.contact, 'ช่องทางติดต่อ')
  };
  if (!['แจ้งของหาย', 'แจ้งพบสิ่งของ'].includes(item.type)) {
    throw new Error('ประเภทการแจ้งไม่ถูกต้อง');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getItemSheet_();
    sheet.appendRow([
      Utilities.getUuid(),
      item.type,
      item.itemName,
      item.category,
      item.location,
      item.date,
      item.contact,
      'กำลังติดตาม'
    ]);
    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function markItemReturned(itemId) {
  requireRegisteredUser_();
  const id = requiredText_(itemId, 'รหัสรายการ');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getItemSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: 'ไม่พบรายการที่ต้องการอัปเดต' };
    }

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    const rowOffset = ids.findIndex(row => row[0] === id);
    if (rowOffset === -1) {
      return { success: false, message: 'ไม่พบรายการที่ต้องการอัปเดต' };
    }

    sheet.getRange(rowOffset + 2, ITEM_HEADERS.length).setValue('คืนเจ้าของแล้ว');
    return { success: true, message: 'อัปเดตสถานะคืนเจ้าของแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function getActiveUserEmail_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

function requireRegisteredUser_() {
  const email = getActiveUserEmail_();
  if (!email || !findUserProfile_(email)) {
    throw new Error('กรุณาเข้าสู่ระบบและลงทะเบียนบัญชี Google ก่อนใช้งาน');
  }
}

function findUserProfile_(email) {
  const sheet = getUserSheet_();
  const row = findUserRow_(sheet, email);
  if (!row) return null;
  return {
    email: row[0],
    name: row[1],
    phone: row[2],
    faculty: row[3],
    role: row[4],
    identifier: row[5]
  };
}

function findUserRow_(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getDisplayValues();
  return rows.find(row => row[0].trim().toLowerCase() === email) || null;
}

function getUserSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('ไม่พบ Google Sheets ที่เชื่อมกับ Apps Script นี้ กรุณาเปิด Apps Script จากไฟล์ Google Sheets');
  }
  let sheet = spreadsheet.getSheetByName(USER_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(USER_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(USER_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getItemSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('ไม่พบ Google Sheets ที่เชื่อมกับ Apps Script นี้ กรุณาเปิด Apps Script จากไฟล์ Google Sheets');
  }

  let sheet = spreadsheet.getSheetByName(ITEM_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(ITEM_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ITEM_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function requiredText_(value, label) {
  const text = optionalText_(value);
  if (!text) throw new Error(`กรุณาระบุ${label}`);
  return text;
}

function optionalText_(value) {
  return String(value == null ? '' : value).trim();
}

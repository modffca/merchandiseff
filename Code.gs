// =============================================
// MERCHANDISE MANAGEMENT SYSTEM - Google Apps Script
// Version 1.0
// =============================================

const SPREADSHEET_ID = '1MmJTP9feqeKmp-_b9Bkn2G042dBa4DocUn_VZIyEdXw';
const SHEET_MERCH = 'List Merch';
const SHEET_USERS = 'Users';
const SHEET_REQUESTS = 'Requests';

// =============================================
// MAIN ROUTER - handles all incoming requests
// =============================================
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'getMerchandise') {
      result = getMerchandise();
    } else if (action === 'getJenis') {
      result = getJenis();
    } else {
      result = { success: false, message: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  let result;

  try {
    if (action === 'register') {
      result = registerUser(body);
    } else if (action === 'login') {
      result = loginUser(body);
    } else if (action === 'addRequest') {
      result = addRequest(body);
    } else if (action === 'updateStatus') {
      result = updateRequestStatus(body);
    } else if (action === 'createMerchandise') {
      result = createMerchandise(body);
    } else if (action === 'updateMerchandise') {
      result = updateMerchandise(body);
    } else if (action === 'deleteMerchandise') {
      result = deleteMerchandise(body);
    } else {
      result = { success: false, message: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================
// HELPER FUNCTIONS
// =============================================
function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function generateId(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5).toUpperCase();
}

function hashPassword(password) {
  // Simple hash using Utilities - for production use a proper hashing library
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// =============================================
// MERCHANDISE FUNCTIONS
// =============================================
function getMerchandise() {
  const sheet = getSheet(SHEET_MERCH);
  const items = sheetToObjects(sheet);
  return { success: true, data: items };
}

function getJenis() {
  const sheet = getSheet(SHEET_MERCH);
  const items = sheetToObjects(sheet);
  const jenis = [...new Set(items.map(i => i.jenis).filter(Boolean))];
  return { success: true, data: jenis };
}

function createMerchandise(body) {
  const sheet = getSheet(SHEET_MERCH);
  const id = generateId('MRC');
  const timestamp = new Date().toISOString();
  sheet.appendRow([id, body.nama, body.jenis, body.qty, body.foto_url || '', body.note || '', timestamp]);
  return { success: true, message: 'Merchandise created', id };
}

function updateMerchandise(body) {
  const sheet = getSheet(SHEET_MERCH);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === body.id) {
      const timestamp = new Date().toISOString();
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([[
        body.id,
        body.nama,
        body.jenis,
        body.qty,
        body.foto_url || data[i][headers.indexOf('foto_url')],
        body.note || '',
        timestamp
      ]]);
      return { success: true, message: 'Merchandise updated' };
    }
  }
  return { success: false, message: 'Item not found' };
}

function deleteMerchandise(body) {
  const sheet = getSheet(SHEET_MERCH);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === body.id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Merchandise deleted' };
    }
  }
  return { success: false, message: 'Item not found' };
}

// =============================================
// USER / AUTH FUNCTIONS
// =============================================
function registerUser(body) {
  const sheet = getSheet(SHEET_USERS);
  const users = sheetToObjects(sheet);

  // Check if email already exists
  const exists = users.find(u => u.email === body.email);
  if (exists) {
    return { success: false, message: 'Email sudah terdaftar' };
  }

  const id = generateId('USR');
  const hashedPassword = hashPassword(body.password);
  sheet.appendRow([id, body.nama, body.email, hashedPassword, body.no_telp || '', 'user']);

  return { success: true, message: 'Registrasi berhasil', userId: id };
}

function loginUser(body) {
  const sheet = getSheet(SHEET_USERS);
  const users = sheetToObjects(sheet);
  const hashedPassword = hashPassword(body.password);

  const user = users.find(u => u.email === body.email && u.password === hashedPassword);
  if (!user) {
    return { success: false, message: 'Email atau password salah' };
  }

  // Return user data WITHOUT password
  return {
    success: true,
    message: 'Login berhasil',
    user: {
      id: user.id,
      nama: user.nama,
      email: user.email,
      no_telp: user.no_telp,
      role: user.role
    }
  };
}

// =============================================
// REQUEST FUNCTIONS
// =============================================
function addRequest(body) {
  const sheet = getSheet(SHEET_REQUESTS);
  const id = generateId('REQ');
  const timestamp = new Date().toISOString();
  sheet.appendRow([id, body.user_id, body.nama_user, body.barang_id, body.nama_barang, body.qty, 'pending', timestamp]);
  return { success: true, message: 'Request berhasil dikirim', id };
}

function updateRequestStatus(body) {
  const reqSheet = getSheet(SHEET_REQUESTS);
  const reqData = reqSheet.getDataRange().getValues();
  const reqHeaders = reqData[0];
  const idCol = reqHeaders.indexOf('id');
  const statusCol = reqHeaders.indexOf('status');
  const barangIdCol = reqHeaders.indexOf('barang_id');
  const qtyCol = reqHeaders.indexOf('qty');

  for (let i = 1; i < reqData.length; i++) {
    if (reqData[i][idCol] === body.request_id) {
      // Update status
      reqSheet.getRange(i + 1, statusCol + 1).setValue(body.status);

      // If approved, reduce stock
      if (body.status === 'approve') {
        const barangId = reqData[i][barangIdCol];
        const qty = reqData[i][qtyCol];
        reduceStock(barangId, qty);
      }

      return { success: true, message: `Request ${body.status}` };
    }
  }
  return { success: false, message: 'Request not found' };
}

function reduceStock(barangId, qty) {
  const sheet = getSheet(SHEET_MERCH);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const qtyCol = headers.indexOf('qty');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === barangId) {
      const currentQty = parseInt(data[i][qtyCol]);
      const newQty = Math.max(0, currentQty - parseInt(qty));
      sheet.getRange(i + 1, qtyCol + 1).setValue(newQty);
      // Update timestamp
      const lastUpdateCol = headers.indexOf('last_update');
      if (lastUpdateCol >= 0) {
        sheet.getRange(i + 1, lastUpdateCol + 1).setValue(new Date().toISOString());
      }
      return;
    }
  }
}

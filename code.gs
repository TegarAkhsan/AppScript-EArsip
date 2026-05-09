/**
 * E-Arsip Digital System
 * Based on provided flowchart
 */

const CONFIG = {
  SPREADSHEET_ID: '1_PoQaAhM-Tz0ojG4Qc1D23mnSKN0IfFiY-R8ZYN6nAE',
  FOLDER_MASUK: '1X4D10qzOM9LjoirlW5xmk7HisVUtD-T0',
  FOLDER_KELUAR: '12pCSn3ohmQQbCkFqDD8D7CrQXnr8LOgW',
  SHEET_USERS: 'Users',
  SHEET_ARCHIVES: 'Archives'
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('E-Arsip Digital')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * DATABASE INITIALIZATION
 */
function getDb() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // Setup Sheets
  const sheets = [CONFIG.SHEET_USERS, CONFIG.SHEET_ARCHIVES];
  sheets.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name === CONFIG.SHEET_USERS) {
        sheet.appendRow(['Email', 'Password', 'PIN', 'Nama Desa', 'Alamat Desa', 'OTP']);
        // Add Default Admin
        sheet.appendRow(['admin@earsip.com', 'admin123', '123456', 'Pusat Digital', 'Jl. Arsitektur No. 1', '']);
      } else if (name === CONFIG.SHEET_ARCHIVES) {
        sheet.appendRow(['ID', 'Timestamp', 'Nomor', 'Nama Pemilik', 'Jenis', 'FileName', 'FileID', 'FileType']);
      }
    }
  });
  return ss;
}

/**
 * AUTHENTICATION LOGIC
 */
function login(email, password) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1] === password) {
      return { success: true, user: { email: data[i][0], namaDesa: data[i][3], alamatDesa: data[i][4] } };
    }
  }
  return { success: false, message: 'Email atau Password salah' };
}

function register(email, password, pin) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  // Check if exists
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) return { success: false, message: 'Email sudah terdaftar' };
  }
  
  sheet.appendRow([email, password, pin, 'Desa Contoh', 'Alamat Contoh', '']);
  return { success: true };
}

function verifyPin(email, pin) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][2] == pin) {
      return { success: true };
    }
  }
  return { success: false, message: 'PIN salah' };
}

/**
 * FORGOT PASSWORD & OTP
 */
function sendOtp(email) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      const otp = Math.floor(100000 + Math.random() * 900000).toString();
      sheet.getRange(i + 1, 6).setValue(otp);
      
      // Simulate sending email (in real scenario use MailApp)
      // MailApp.sendEmail(email, 'OTP Lupa Password', 'Kode OTP Anda: ' + otp);
      return { success: true, otp: otp }; // Returning OTP for demo purposes, normally wouldn't
    }
  }
  return { success: false, message: 'Email tidak ditemukan' };
}

function resetPassword(email, otp, newPassword) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][5] == otp) {
      sheet.getRange(i + 1, 2).setValue(newPassword);
      sheet.getRange(i + 1, 6).setValue(''); // Clear OTP
      return { success: true };
    }
  }
  return { success: false, message: 'OTP tidak valid' };
}

/**
 * ARCHIVE LOGIC
 */
function uploadArsip(obj) {
  try {
    const folderId = obj.jenis === 'Masuk' ? CONFIG.FOLDER_MASUK : CONFIG.FOLDER_KELUAR;
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, obj.fileName);
    const file = folder.createFile(blob);
    
    // Set sharing to anyone with link for easy preview/download within the app
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVES);
    const id = 'ARS-' + new Date().getTime();
    
    sheet.appendRow([
      id,
      new Date(),
      obj.nomor,
      obj.namaPemilik,
      obj.jenis,
      obj.fileName,
      file.getId(),
      obj.mimeType
    ]);
    
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function searchArsip(query) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVES);
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  const q = query.toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const nomor = data[i][2].toString().toLowerCase();
    const nama = data[i][3].toString().toLowerCase();
    const jenis = data[i][4].toString().toLowerCase();
    
    if (nomor.includes(q) || nama.includes(q) || jenis.includes(q)) {
      results.push({
        id: data[i][0],
        timestamp: data[i][1],
        nomor: data[i][2],
        namaPemilik: data[i][3],
        jenis: data[i][4],
        fileName: data[i][5],
        fileId: data[i][6],
        mimeType: data[i][7],
        viewUrl: `https://drive.google.com/file/d/${data[i][6]}/preview`,
        downloadUrl: `https://drive.google.com/uc?export=download&id=${data[i][6]}`,
        printUrl: `https://drive.google.com/file/d/${data[i][6]}/view?usp=sharing`
      });
    }
  }
  return results;
}

function updateArsipName(id, newName) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 4).setValue(newName); // Assuming we edit Nama Pemilik as per flowchart "Edit nama dokumen"
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * SETTINGS LOGIC
 */
function updateProfile(email, namaDesa, alamatDesa) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 4).setValue(namaDesa);
      sheet.getRange(i + 1, 5).setValue(alamatDesa);
      return { success: true };
    }
  }
  return { success: false };
}

function changePin(email, oldPin, newPin) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][2] == oldPin) {
      sheet.getRange(i + 1, 3).setValue(newPin);
      return { success: true };
    }
  }
  return { success: false, message: 'PIN lama salah' };
}

function changePassword(email, oldPass, newPass) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1] == oldPass) {
      sheet.getRange(i + 1, 2).setValue(newPass);
      return { success: true };
    }
  }
  return { success: false, message: 'Password lama salah' };
}

/**
 * HELPERS
 */
function getFolder() {
  if (CONFIG.FOLDER_ID) return DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const folders = DriveApp.getFoldersByName('E-Arsip-Uploads');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('E-Arsip-Uploads');
}

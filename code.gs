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
    
    // Attempt to set sharing, but don't crash if organization policies restrict it
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      console.warn('Gagal mengatur sharing: ' + e.toString());
    }
    
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVES);
    const id = 'ARS-' + new Date().getTime();
    
    // Faster way to append data
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 8).setValues([[
      id,
      new Date(),
      obj.nomor,
      obj.namaPemilik,
      obj.jenis,
      obj.fileName,
      file.getId(),
      obj.mimeType
    ]]);
    
    SpreadsheetApp.flush(); // Commit changes immediately
    
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function searchArsip(query) {
  try {
    const ss = getDb();
    SpreadsheetApp.flush(); // Ensure all writes are committed
    
    // Robust sheet selection
    const sheetName = CONFIG.SHEET_ARCHIVES.trim().toLowerCase();
    const sheet = ss.getSheets().find(s => s.getName().trim().toLowerCase() === sheetName);
    
    if (!sheet) {
      console.error('Sheet not found by name:', CONFIG.SHEET_ARCHIVES);
      return { error: 'Sheet "Archives" tidak ditemukan di Spreadsheet.' };
    }
    
    const rawData = sheet.getDataRange().getValues();
    console.log('Raw data length:', rawData.length);
    
    // Filter out truly empty rows
    const data = rawData.filter(row => row[0] !== "" && row[0] !== null && row[0] !== undefined);
    console.log('Filtered data length (with headers):', data.length);
    
    if (data.length <= 1) {
      console.log('No data rows found after filtering.');
      return [];
    }
    
    const results = [];
    const q = (query || "").toLowerCase().trim();
    
    // Process from newest (bottom) to oldest (top), skipping header at index 0
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const nomor = (row[2] || "").toString();
      const nama = (row[3] || "").toString();
      const jenis = (row[4] || "").toString();
      
      // Match query
      if (!q || nomor.toLowerCase().includes(q) || nama.toLowerCase().includes(q) || jenis.toLowerCase().includes(q)) {
        results.push({
          id: row[0].toString(),
          timestamp: row[1],
          nomor: nomor || '-',
          namaPemilik: nama || '-',
          jenis: jenis || '-',
          fileName: row[5] || 'Tanpa Nama',
          fileId: row[6],
          mimeType: row[7],
          viewUrl: row[6] ? `https://drive.google.com/file/d/${row[6]}/preview` : '#',
          downloadUrl: row[6] ? `https://drive.google.com/uc?export=download&id=${row[6]}` : '#',
          printUrl: row[6] ? `https://drive.google.com/file/d/${row[6]}/view?usp=sharing` : '#'
        });
      }
      
      // Limit to 20 for search results, 10 for dashboard
      const limit = q ? 20 : 10;
      if (results.length >= limit) break;
    }
    
    console.log('Returning results:', results.length);
    return results;
  } catch (e) {
    console.error('Search Error Detail:', e.stack);
    return { error: 'Gagal memuat data: ' + e.toString() };
  }
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

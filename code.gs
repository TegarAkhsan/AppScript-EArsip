/**
 * E-Arsip Digital System
 * Based on provided flowchart
 */

const CONFIG = {
  SPREADSHEET_ID: '1_PoQaAhM-Tz0ojG4Qc1D23mnSKN0IfFiY-R8ZYN6nAE',
  FOLDER_ARCHIVE: '1X4D10qzOM9LjoirlW5xmk7HisVUtD-T0', // Using FOLDER_MASUK as default
  SHEET_USERS: 'Users',
  SHEET_ARCHIVES: 'Archives',
  SHEET_TRASH: 'Trash'
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('BERBETA')
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
  const sheets = [CONFIG.SHEET_USERS, CONFIG.SHEET_ARCHIVES, CONFIG.SHEET_TRASH];
  sheets.forEach(name => {
    // Robust search to avoid "sheet already exists" errors due to casing or trailing spaces
    const targetName = name.trim().toLowerCase();
    let sheet = ss.getSheets().find(s => s.getName().trim().toLowerCase() === targetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name === CONFIG.SHEET_USERS) {
        sheet.appendRow(['Email', 'Password', 'PIN', 'Nama Desa', 'Alamat Desa', 'OTP']);
        // Add Default Admin
        sheet.appendRow(['admin@earsip.com', 'admin123', '123456', 'Pusat Digital', 'Jl. Arsitektur No. 1', '']);
      } else if (name === CONFIG.SHEET_ARCHIVES) {
        sheet.appendRow(['ID', 'Timestamp', 'Nomor', 'Nama Pemilik', 'FileName', 'FileID', 'FileType']);
      } else if (name === CONFIG.SHEET_TRASH) {
        sheet.appendRow(['ID', 'Timestamp', 'Nomor', 'Nama Pemilik', 'FileName', 'FileID', 'FileType', 'DeletedAt']);
      }
    }
  });

  // Auto-setup trash cleanup trigger
  try {
    setupTrashTrigger();
  } catch (err) {
    console.warn('Gagal setup trigger: ' + err.toString());
  }

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

function loginWithPin(pin) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  // Search for user by PIN
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == pin) {
      return { 
        success: true, 
        user: { email: data[i][0], namaDesa: data[i][3], alamatDesa: data[i][4] } 
      };
    }
  }
  return { success: false, message: 'PIN tidak valid atau tidak terdaftar' };
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
    const folderId = CONFIG.FOLDER_ARCHIVE;
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
    sheet.getRange(lastRow + 1, 1, 1, 7).setValues([[
      id,
      new Date(),
      obj.nomor,
      obj.namaPemilik,
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
      
      // Match query
      if (!q || nomor.toLowerCase().includes(q) || nama.toLowerCase().includes(q)) {
        const hasJenis = row.length >= 8;
        results.push({
          id: row[0].toString(),
          timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1].toString(),
          nomor: nomor || '-',
          namaPemilik: nama || '-',
          fileName: (hasJenis ? row[5] : row[4]) || 'Tanpa Nama',
          fileId: hasJenis ? row[6] : row[5],
          mimeType: hasJenis ? row[7] : row[6],
          viewUrl: (hasJenis ? row[6] : row[5]) ? `https://drive.google.com/file/d/${hasJenis ? row[6] : row[5]}/preview` : '#',
          downloadUrl: (hasJenis ? row[6] : row[5]) ? `https://drive.google.com/uc?export=download&id=${hasJenis ? row[6] : row[5]}` : '#',
          printUrl: (hasJenis ? row[6] : row[5]) ? `https://drive.google.com/file/d/${hasJenis ? row[6] : row[5]}/view?usp=sharing` : '#'
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

function updateArsip(id, newNomor, newName) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVES);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 3).setValue(newNomor); // Column C
        sheet.getRange(i + 1, 4).setValue(newName);  // Column D
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, message: 'ID tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteArsip(id) {
  try {
    const ss = getDb();
    const sheetArchives = ss.getSheetByName(CONFIG.SHEET_ARCHIVES);
    const sheetTrash = ss.getSheetByName(CONFIG.SHEET_TRASH);
    const data = sheetArchives.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        const row = data[i];
        
        // Append to trash sheet
        const lastRow = sheetTrash.getLastRow();
        sheetTrash.getRange(lastRow + 1, 1, 1, 8).setValues([[
          row[0], // ID
          row[1], // Timestamp
          row[2], // Nomor
          row[3], // Nama Pemilik
          row[4], // FileName
          row[5], // FileID
          row[6], // FileType
          new Date() // DeletedAt
        ]]);
        
        // Delete from active archives
        sheetArchives.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Arsip berhasil dipindahkan ke Sampah.' };
      }
    }
    return { success: false, message: 'ID tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
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
 * TRASH MANAGEMENT LOGIC
 */
function getTrashItems() {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_TRASH);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const results = [];
    const now = new Date();
    
    // Read from newest to oldest
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (!row[0]) continue;
      
      const deletedAtDate = row[7] instanceof Date ? row[7] : new Date(row[7]);
      const diffTime = Math.abs(now - deletedAtDate);
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      const remainingDays = Math.max(0, 30 - diffDays);
      
      results.push({
        id: row[0].toString(),
        timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1].toString(),
        nomor: (row[2] || "").toString() || '-',
        namaPemilik: (row[3] || "").toString() || '-',
        fileName: (row[4] || "").toString() || 'Tanpa Nama',
        fileId: row[5],
        mimeType: row[6],
        deletedAt: deletedAtDate.toISOString(),
        remainingDays: remainingDays
      });
    }
    return results;
  } catch (e) {
    console.error('getTrashItems error:', e.toString());
    return [];
  }
}

function restoreArsip(id) {
  try {
    const ss = getDb();
    const sheetArchives = ss.getSheetByName(CONFIG.SHEET_ARCHIVES);
    const sheetTrash = ss.getSheetByName(CONFIG.SHEET_TRASH);
    const data = sheetTrash.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        const row = data[i];
        
        // Append back to Archives sheet
        const lastRow = sheetArchives.getLastRow();
        sheetArchives.getRange(lastRow + 1, 1, 1, 7).setValues([[
          row[0], // ID
          row[1], // Timestamp
          row[2], // Nomor
          row[3], // Nama Pemilik
          row[4], // FileName
          row[5], // FileID
          row[6]  // FileType
        ]]);
        
        // Remove from Trash sheet
        sheetTrash.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, message: 'ID tidak ditemukan di Sampah' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deletePermanently(id) {
  try {
    const ss = getDb();
    const sheetTrash = ss.getSheetByName(CONFIG.SHEET_TRASH);
    const data = sheetTrash.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        // Delete the file from Drive
        const fileId = data[i][5];
        if (fileId) {
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (err) {
            console.warn('Gagal menghapus file di Drive:', err.toString());
          }
        }
        
        // Remove from Trash sheet
        sheetTrash.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, message: 'ID tidak ditemukan di Sampah' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function cleanupTrash() {
  try {
    const ss = getDb();
    const sheetTrash = ss.getSheetByName(CONFIG.SHEET_TRASH);
    if (!sheetTrash) return;
    
    const data = sheetTrash.getDataRange().getValues();
    if (data.length <= 1) return;
    
    const now = new Date();
    let rowsDeleted = 0;
    
    // Process from bottom to top so that index deletion doesn't shift unprocessed rows
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (!row[0]) continue;
      
      const deletedAtDate = row[7] instanceof Date ? row[7] : new Date(row[7]);
      const diffTime = Math.abs(now - deletedAtDate);
      const diffDays = diffTime / (1000 * 60 * 60 * 24);
      
      if (diffDays > 30) {
        // Delete file from Drive
        const fileId = row[5];
        if (fileId) {
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (err) {
            console.warn('Gagal menghapus file Drive saat cleanup:', err.toString());
          }
        }
        
        // Delete row
        sheetTrash.deleteRow(i + 1);
        rowsDeleted++;
      }
    }
    
    if (rowsDeleted > 0) {
      SpreadsheetApp.flush();
      console.log(`Cleaned up ${rowsDeleted} trash items older than 30 days.`);
    }
  } catch (e) {
    console.error('cleanupTrash error:', e.toString());
  }
}

function setupTrashTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let hasTrigger = false;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cleanupTrash') {
      hasTrigger = true;
      break;
    }
  }
  
  if (!hasTrigger) {
    ScriptApp.newTrigger('cleanupTrash')
      .timeBased()
      .everyDays(1)
      .create();
  }
}

function getFolder() {
  const folders = DriveApp.getFoldersByName('E-Arsip-Uploads');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('E-Arsip-Uploads');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Buku Tamu Dinkes Kab. HSU')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Buat sheet untuk data tamu jika belum ada
    var tamuSheet = ss.getSheetByName('DataTamu');
    if (!tamuSheet) {
      tamuSheet = ss.insertSheet('DataTamu');
      var headers = ['ID', 'Tanggal', 'Nama', 'Institusi/Asal', 'Jenis Institusi', 'Keperluan', 'Email', 'Nomor Telepon', 'Catatan Tambahan'];
      tamuSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      tamuSheet.setFrozenRows(1);
    }
    
    // Buat sheet untuk admin credentials jika belum ada
    var adminSheet = ss.getSheetByName('AdminCredentials');
    if (!adminSheet) {
      adminSheet = ss.insertSheet('AdminCredentials');
      var adminHeaders = ['Username', 'Password', 'Nama Lengkap'];
      adminSheet.getRange(1, 1, 1, adminHeaders.length).setValues([adminHeaders]).setFontWeight('bold');
      
      // Tambahkan admin default
      var defaultAdmin = ['admin', 'admin123', 'Administrator'];
      adminSheet.getRange(2, 1, 1, defaultAdmin.length).setValues([defaultAdmin]);
      adminSheet.setFrozenRows(1);
    }
    
    return {success: true, message: 'Setup berhasil dilakukan!'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Fungsi untuk menambahkan data tamu
function tambahTamu(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('DataTamu');
    
    // Generate ID (timestamp)
    var id = new Date().getTime().toString();
    var tanggal = new Date().toLocaleString('id-ID');
    
    // Menyiapkan data untuk dimasukkan ke sheet
    var rowData = [
      id,
      tanggal,
      data.nama,
      data.institusi,
      data.jenisInstitusi,
      data.keperluan,
      data.email || '-',
      data.telepon || '-',
      data.catatan || '-'
    ];
    
    // Menambahkan data ke sheet
    sheet.appendRow(rowData);
    
    return {success: true, message: 'Data tamu berhasil ditambahkan!'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Fungsi untuk mendapatkan semua data tamu
function getAllTamu() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('DataTamu');
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {success: true, data: []};
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var result = data.map(function(row) {
      var obj = {};
      for (var i = 0; i < headers.length; i++) {
        obj[headers[i]] = row[i];
      }
      return obj;
    });
    
    return {success: true, data: result};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Fungsi untuk menghapus data tamu
function hapusTamu(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('DataTamu');
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(1, 1, lastRow, 1).getValues(); // Ambil kolom ID
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return {success: true, message: 'Data tamu berhasil dihapus!'};
      }
    }
    
    return {success: false, message: 'ID tamu tidak ditemukan!'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Fungsi untuk login admin
function loginAdmin(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('AdminCredentials');
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(1, 1, lastRow, 3).getValues(); // Username, Password, Nama Lengkap
    
    for (var i = 1; i < data.length; i++) { // Mulai dari 1 untuk skip header
      if (data[i][0] == username && data[i][1] == password) {
        return {
          success: true, 
          message: 'Login berhasil!',
          admin: {
            username: data[i][0],
            nama: data[i][2]
          }
        };
      }
    }
    
    return {success: false, message: 'Username atau password salah!'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Fungsi untuk mendapatkan statistik
function getStatistik() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('DataTamu');
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: true, 
        totalTamu: 0,
        tamuMingguIni: 0,  // Changed from tamuHariIni to tamuMingguIni
        jenisInstitusi: {
          perorangan: 0,
          institusi: 0
        }
      };
    }
    
    // Ambil semua data
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Hitung total tamu
    var totalTamu = data.length;
    
    // Get today's date and date 7 days ago
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(today.getDate() - 7);
    sevenDaysAgo.setHours(0, 0, 0, 0);
    
    // Hitung tamu 7 hari terakhir
    var tamuMingguIni = 0;  // Changed from tamuHariIni to tamuMingguIni
    
    // Hitung berdasarkan jenis institusi
    var perorangan = 0;
    var institusi = 0;
    
    for (var i = 0; i < data.length; i++) {
      var tanggal = new Date(data[i][1]);
      if (tanggal >= sevenDaysAgo && tanggal <= today) {
        tamuMingguIni++;
      }
      
      if (data[i][4] === 'Perorangan') {
        perorangan++;
      } else if (data[i][4] === 'Mewakili Institusi') {
        institusi++;
      }
    }
    
    return {
      success: true,
      totalTamu: totalTamu,
      tamuMingguIni: tamuMingguIni,  // Changed from tamuHariIni to tamuMingguIni
      jenisInstitusi: {
        perorangan: perorangan,
        institusi: institusi
      }
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

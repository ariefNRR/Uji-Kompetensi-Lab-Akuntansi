// ==========================================
// 1. FUNGSI UTAMA (SETUP & TAMPILAN)
// ==========================================
function doGet(e) {
  pastikanSemuaSheetAda();
  // Jika diakses dengan parameter ?api=check atau file Index.html tidak ada di Apps Script,
  // kembalikan respon JSON sebagai penanda bahwa API aktif.
  if (e && e.parameter && e.parameter.api === 'check') {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "aktif", pesan: "API LMS CBT Akuntansi UNY siap digunakan!" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('LMS CBT Akuntansi - UNY')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "aktif", pesan: "Backend API LMS CBT Akuntansi UNY aktif dan siap dihubungkan ke Netlify." }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 1B. ENDPOINT API UNTUK FRONTEND NETLIFY
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Data permintaan kosong.");
    }
    var req = JSON.parse(e.postData.contents);
    var action = req.action;
    var args = req.args || [];

    var allowedFunctions = {
      registerUser: registerUser,
      loginUser: loginUser,
      getSecurityQuestion: getSecurityQuestion,
      resetPassword: resetPassword,
      getDashboardStats: getDashboardStats,
      getAllUsers: getAllUsers,
      addUserManual: addUserManual,
      getDaftarFolder: getDaftarFolder,
      tambahFolderBaru: tambahFolderBaru,
      getIsiMateri: getIsiMateri,
      tambahMateriBaru: tambahMateriBaru,
      hapusDataMateri: hapusDataMateri,
      updateMateri: updateMateri,
      getAllUjian: getAllUjian,
      tambahUjianBaru: tambahUjianBaru,
      updateUjian: updateUjian,
      hapusUjian: hapusUjian,
      simpanHasilUjian: simpanHasilUjian,
      getRekapNilai: getRekapNilai,
      catatPelanggaranTab: catatPelanggaranTab,
      getLogKecurangan: getLogKecurangan,
      hapusSemuaLogKecurangan: hapusSemuaLogKecurangan,
      uploadGambarSoal: uploadGambarSoal,
      getAllSoal: getAllSoal,
      getSoalByKategori: getSoalByKategori,
      getDaftarKategoriSoal: getDaftarKategoriSoal,
      tambahBanyakSoal: tambahBanyakSoal,
      updateSoal: updateSoal,
      hapusSoal: hapusSoal
    };

    if (!allowedFunctions[action]) {
      throw new Error("Fungsi tidak ditemukan atau tidak diizinkan: " + action);
    }

    var result = allowedFunctions[action].apply(null, args);
    return ContentService
      .createTextOutput(JSON.stringify({ result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message || err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function pastikanSemuaSheetAda() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ["Users", "FolderMateri", "IsiMateri", "Ujian", "Soal", "HasilUjian", "LogKecurangan"];
  var headers = [
    ["ID_User", "Nama", "Email", "Password", "Role", "Pertanyaan", "Jawaban"],
    ["ID Folder", "Nama Folder", "Tanggal Buat"],
    ["ID Materi", "Nama Folder", "Tipe", "Judul", "Link/Konten", "Tanggal Buat"],
    ["Nama Ujian", "Waktu Mulai", "Waktu Selesai", "Durasi", "Token", "Status", "Bank Soal"],
    ["Kategori", "Pertanyaan", "A", "B", "C", "D", "E", "Kunci", "Gambar"],
    ["ID_Hasil", "Nama Ujian", "Nama Peserta", "Email Peserta", "Benar", "Salah", "Total Soal", "Nilai", "Waktu Selesai", "Pindah Tab"],
    ["ID_Log", "Waktu", "Nama Ujian", "Nama Peserta", "Email Peserta", "Keterangan", "Status"]
  ];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = ss.getSheetByName(sheets[i]);
    if (!sheet) {
      sheet = ss.insertSheet(sheets[i]);
      sheet.appendRow(headers[i]);
      if (sheets[i] === "Users") {
        sheet.appendRow(["1", "Admin Lab", "admin@uny.ac.id", "admin123", "admin", "Apa nama hewan peliharaan pertama Anda?", "kucing"]);
      }
    } else {
      // Pastikan kolom baru (Gambar di Soal, Pindah Tab di HasilUjian) ditambahkan jika sheet lama sudah ada
      if (sheets[i] === "Soal" && sheet.getLastColumn() < 9) {
        sheet.getRange(1, 9).setValue("Gambar");
      }
      if (sheets[i] === "HasilUjian" && sheet.getLastColumn() < 10) {
        sheet.getRange(1, 10).setValue("Pindah Tab");
      }
    }
  }
}

// ==========================================
// 2. FUNGSI LOGIN, REGISTER & LUPA PASSWORD
// ==========================================
function registerUser(nama, email, pass, pert, jaw) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) {
    pastikanSemuaSheetAda();
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  }
  
  var data = sheet.getDataRange().getValues();
  var inputEmail = email.toString().trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim().toLowerCase() === inputEmail) { 
      return { status: "gagal", pesan: "Email sudah terdaftar!" };
    }
  }
  
  var newId = "USR" + new Date().getTime();
  sheet.appendRow([
    newId,
    nama.toString().trim(), 
    inputEmail, 
    pass.toString().trim(), 
    "peserta", 
    pert ? pert.toString().trim() : "-", 
    jaw ? jaw.toString().trim().toLowerCase() : "-"
  ]);
  
  return { status: "sukses" };
}

function loginUser(email, pass) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) {
    pastikanSemuaSheetAda();
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  }
  
  var data = sheet.getDataRange().getValues();
  var inputEmail = email.toString().trim().toLowerCase();
  var inputPass = pass.toString().trim();

  for (var i = 1; i < data.length; i++) {
    if (data[i][2] !== undefined && data[i][3] !== undefined) {
      var sheetEmail = data[i][2].toString().trim().toLowerCase(); 
      var sheetPass = data[i][3].toString().trim();                
      
      if (sheetEmail === inputEmail && sheetPass === inputPass) {
        var userNama = data[i][1] ? data[i][1].toString() : "Pengguna"; 
        var userRole = data[i][4] ? data[i][4].toString().toLowerCase().trim() : "peserta"; 
        
        return { 
          status: "sukses", 
          user: { 
            nama: userNama, 
            role: userRole 
          } 
        };
      }
    }
  }
  return { status: "gagal", pesan: "Email atau password salah!" };
}

function getSecurityQuestion(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) return { status: "gagal", pesan: "Sheet tidak ditemukan!" };
  
  var data = sheet.getDataRange().getValues();
  var inputEmail = email.toString().trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim().toLowerCase() === inputEmail) { 
      return { status: "sukses", pertanyaan: data[i][5] || "Pertanyaan Keamanan" }; 
    }
  }
  return { status: "gagal", pesan: "Email tidak ditemukan!" };
}

function resetPassword(email, jawaban, passBaru) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) return { status: "gagal", pesan: "Terjadi kesalahan sistem." };
  
  var data = sheet.getDataRange().getValues();
  var inputEmail = email.toString().trim().toLowerCase();
  var inputJawaban = jawaban.toString().trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim().toLowerCase() === inputEmail) { 
      if (data[i][6] && data[i][6].toString().trim().toLowerCase() === inputJawaban) { 
        sheet.getRange(i + 1, 4).setValue(passBaru.toString().trim()); 
        return { status: "sukses", pesan: "Password berhasil diubah!" };
      } else {
        return { status: "gagal", pesan: "Jawaban keamanan salah!" };
      }
    }
  }
  return { status: "gagal", pesan: "Terjadi kesalahan sistem." };
}

// ==========================================
// 3. FUNGSI DASHBOARD & KELOLA PENGGUNA
// ==========================================
function getDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var peserta = 0, admin = 0;
  var sheetUsers = ss.getSheetByName("Users");
  if (sheetUsers) {
    var dataUsers = sheetUsers.getDataRange().getValues();
    for (var i = 1; i < dataUsers.length; i++) {
      if (dataUsers[i][2] && dataUsers[i][2].toString().trim() !== "") { 
        var role = dataUsers[i][4] ? dataUsers[i][4].toString().toLowerCase().trim() : "peserta"; 
        if (role === "admin") admin++; else peserta++;
      }
    }
  }
  var ujian = Math.max(0, (ss.getSheetByName("Ujian") ? ss.getSheetByName("Ujian").getLastRow() - 1 : 0)); 
  var soal = Math.max(0, (ss.getSheetByName("Soal") ? ss.getSheetByName("Soal").getLastRow() - 1 : 0));
  return { peserta: peserta, admin: admin, materi: ujian, soal: soal };
}

function getAllUsers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim() !== "") { 
      result.push({ 
        nama: data[i][1] || "-",                             
        email: data[i][2].toString(),                        
        password: data[i][3] || "-",                         
        role: data[i][4] ? data[i][4].toString().toLowerCase().trim() : "peserta" 
      });
    }
  }
  return result;
}

function addUserManual(nama, email, pass, role) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  var inputEmail = email.toString().trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim().toLowerCase() === inputEmail) { 
      return { status: "gagal", pesan: "Email sudah digunakan!" };
    }
  }
  var newId = "USR" + new Date().getTime();
  sheet.appendRow([newId, nama, inputEmail, pass, role, "Dibuat oleh Admin", "admin"]);
  return { status: "sukses" };
}

// ==========================================
// 4. FUNGSI KELOLA MATERI (FOLDER & ISI)
// ==========================================
function getDaftarFolder() {
  var sheetFolder = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FolderMateri");
  var sheetIsi = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IsiMateri");
  var dataFolder = sheetFolder ? sheetFolder.getDataRange().getValues() : [];
  var dataIsi = sheetIsi ? sheetIsi.getDataRange().getValues() : [];
  var result = [];
  
  for (var i = 1; i < dataFolder.length; i++) {
    var namaFolder = dataFolder[i][1];
    var jumlahItem = 0;
    for (var j = 1; j < dataIsi.length; j++) {
      if (dataIsi[j][1] === namaFolder) jumlahItem++;
    }
    result.push({ nama: namaFolder, jumlah: jumlahItem });
  }
  return result;
}

function tambahFolderBaru(namaFolder) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FolderMateri");
  var id = "FLD-" + new Date().getTime();
  var tanggal = Utilities.formatDate(new Date(), "GMT+7", "dd MMM yyyy");
  sheet.appendRow([id, namaFolder, tanggal]);
  return { status: "sukses" };
}

function getIsiMateri(namaFolder) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IsiMateri");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === namaFolder.toString().trim()) {
      var tgl = data[i][5];
      if (tgl instanceof Date) {
        tgl = Utilities.formatDate(tgl, "GMT+7", "dd MMM yyyy");
      } else {
        tgl = tgl ? tgl.toString() : "-";
      }

      result.push({
        id: data[i][0] ? data[i][0].toString() : "",
        tipe: data[i][2] ? data[i][2].toString() : "TEKS",
        judul: data[i][3] ? data[i][3].toString() : "-",
        konten: data[i][4] ? data[i][4].toString() : "",
        tanggal: tgl
      });
    }
  }
  return result;
}

function tambahMateriBaru(namaFolder, tipe, judul, konten) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IsiMateri");
  var id = "MTR-" + new Date().getTime();
  var tanggal = Utilities.formatDate(new Date(), "GMT+7", "dd MMM yyyy");
  sheet.appendRow([id, namaFolder, tipe, judul, konten, tanggal]);
  return { status: "sukses" };
}

function hapusDataMateri(idMateri) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IsiMateri");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === idMateri.toString()) {
      sheet.deleteRow(i + 1);
      return { status: "sukses" };
    }
  }
  return { status: "gagal", pesan: "Materi tidak ditemukan" };
}

function updateMateri(idMateri, tipe, judulBaru, kontenBaru) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IsiMateri");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === idMateri.toString()) {
      sheet.getRange(i + 1, 3).setValue(tipe);       
      sheet.getRange(i + 1, 4).setValue(judulBaru);  
      sheet.getRange(i + 1, 5).setValue(kontenBaru); 
      return { status: "sukses" };
    }
  }
  return { status: "gagal", pesan: "Materi tidak ditemukan" };
}

// ==========================================
// 5. FUNGSI UJIAN (CRUD), REKAP NILAI & LOG KECURANGAN
// ==========================================
function formatDateForInput(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd'T'HH:mm");
  }
  return val.toString();
}

function getAllUjian() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ujian");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({ 
      row: i + 1,
      nama: data[i][0].toString(), 
      mulai: formatDateForInput(data[i][1]),
      selesai: formatDateForInput(data[i][2]),
      durasi: data[i][3], 
      token: data[i][4] ? data[i][4].toString() : "", 
      status: data[i][5] ? data[i][5].toString() : "Aktif",
      bankSoal: data[i][6] ? data[i][6].toString() : "-"
    });
  }
  return result;
}

function tambahUjianBaru(nama, mulai, selesai, durasi, token, bankSoal) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ujian");
  sheet.appendRow([nama, mulai, selesai, durasi, token, "Aktif", bankSoal]);
  return { status: "sukses" };
}

function updateUjian(row, nama, mulai, selesai, durasi, token, status, bankSoal) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ujian");
  if (!sheet || row < 2) return { status: "gagal", pesan: "Data ujian tidak valid." };
  sheet.getRange(row, 1, 1, 7).setValues([[nama, mulai, selesai, durasi, token, status, bankSoal]]);
  return { status: "sukses" };
}

function hapusUjian(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ujian");
  if (!sheet || row < 2) return { status: "gagal" };
  sheet.deleteRow(row);
  return { status: "sukses" };
}

function simpanHasilUjian(namaUjian, namaPeserta, emailPeserta, benar, salah, totalSoal, nilai, pindahTab) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HasilUjian");
  if (!sheet) {
    pastikanSemuaSheetAda();
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HasilUjian");
  }
  var id = "HSL-" + new Date().getTime();
  var waktu = Utilities.formatDate(new Date(), "GMT+7", "dd MMM yyyy HH:mm");
  sheet.appendRow([id, namaUjian, namaPeserta, emailPeserta, benar, salah, totalSoal, nilai, waktu, pindahTab || 0]);
  return { status: "sukses" };
}

function getRekapNilai(namaUjian) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HasilUjian");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === namaUjian.toString().trim()) {
      result.push({
        namaPeserta: data[i][2],
        emailPeserta: data[i][3],
        benar: data[i][4],
        salah: data[i][5],
        totalSoal: data[i][6],
        nilai: data[i][7],
        waktu: data[i][8] ? data[i][8].toString() : "-",
        pindahTab: data[i][9] ? Number(data[i][9]) : 0
      });
    }
  }
  return result;
}

// FITUR BARU: CATAT & AMBIL LOG PERPINDAHAN TAB (INDIKASI KECURANGAN)
function catatPelanggaranTab(namaUjian, namaPeserta, emailPeserta, jumlahKe) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogKecurangan");
  if (!sheet) {
    pastikanSemuaSheetAda();
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogKecurangan");
  }
  var id = "LOG-" + new Date().getTime();
  var waktu = Utilities.formatDate(new Date(), "GMT+7", "dd MMM yyyy HH:mm:ss");
  var ket = "Terdeteksi berpindah tab / meninggalkan jendela ujian (ke-" + jumlahKe + ")";
  sheet.appendRow([id, waktu, namaUjian, namaPeserta, emailPeserta, ket, "Baru"]);
  return { status: "sukses" };
}

function getLogKecurangan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogKecurangan");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  // Urutkan dari yang terbaru (bawah ke atas)
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][0]) continue;
    result.push({
      id: data[i][0].toString(),
      waktu: data[i][1] ? data[i][1].toString() : "-",
      namaUjian: data[i][2] ? data[i][2].toString() : "-",
      namaPeserta: data[i][3] ? data[i][3].toString() : "-",
      emailPeserta: data[i][4] ? data[i][4].toString() : "-",
      keterangan: data[i][5] ? data[i][5].toString() : "-",
      status: data[i][6] ? data[i][6].toString() : "Baru"
    });
  }
  return result;
}

function hapusSemuaLogKecurangan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogKecurangan");
  if (!sheet) return { status: "sukses" };
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  return { status: "sukses" };
}

// ==========================================
// 6. FUNGSI BANK SOAL (UPLOAD GAMBAR, MULTI-SOAL & CRUD)
// ==========================================
function uploadGambarSoal(base64Data, fileName) {
  try {
    var folderName = "Gambar_Bank_Soal_CBT";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    var splitBase = base64Data.split(',');
    var contentType = splitBase[0].split(':')[1].split(';')[0];
    var decoded = Utilities.base64Decode(splitBase[1]);
    var blob = Utilities.newBlob(decoded, contentType, "SOAL_" + new Date().getTime() + "_" + fileName);
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var urlGambar = "https://lh3.googleusercontent.com/d/" + file.getId();
    return { status: "sukses", url: urlGambar };
  } catch (e) {
    return { status: "gagal", pesan: e.toString() };
  }
}

function getAllSoal() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Soal");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    result.push({ 
      row: i + 1,
      kategori: data[i][0] ? data[i][0].toString().trim() : "", 
      pertanyaan: data[i][1] ? data[i][1].toString() : "", 
      a: data[i][2] ? data[i][2].toString() : "",
      b: data[i][3] ? data[i][3].toString() : "",
      c: data[i][4] ? data[i][4].toString() : "",
      d: data[i][5] ? data[i][5].toString() : "",
      e: data[i][6] ? data[i][6].toString() : "",
      kunci: data[i][7] ? data[i][7].toString() : "A",
      gambar: data[i][8] ? data[i][8].toString() : ""
    });
  }
  return result;
}

function getSoalByKategori(kategori) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Soal");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === kategori.toString().trim()) {
      result.push({
        pertanyaan: data[i][1] ? data[i][1].toString() : "",
        a: data[i][2] ? data[i][2].toString() : "",
        b: data[i][3] ? data[i][3].toString() : "",
        c: data[i][4] ? data[i][4].toString() : "",
        d: data[i][5] ? data[i][5].toString() : "",
        e: data[i][6] ? data[i][6].toString() : "",
        kunci: data[i][7] ? data[i][7].toString().trim().toUpperCase() : "A",
        gambar: data[i][8] ? data[i][8].toString() : ""
      });
    }
  }
  return result;
}

function getDaftarKategoriSoal() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Soal");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var kategoriUnik = [];
  for (var i = 1; i < data.length; i++) {
    var kat = data[i][0] ? data[i][0].toString().trim() : "";
    if (kat !== "" && kategoriUnik.indexOf(kat) === -1) {
      kategoriUnik.push(kat);
    }
  }
  return kategoriUnik;
}

function tambahBanyakSoal(kategori, daftarSoal) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Soal");
  var rows = [];
  for (var i = 0; i < daftarSoal.length; i++) {
    var s = daftarSoal[i];
    rows.push([kategori.trim(), s.tanya, s.a, s.b, s.c, s.d, s.e, s.kunci, s.gambar || ""]);
  }
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  }
  return { status: "sukses" };
}

function updateSoal(row, kategori, tanya, a, b, c, d, e, kunci, gambar) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Soal");
  if (!sheet || row < 2) return { status: "gagal" };
  sheet.getRange(row, 1, 1, 9).setValues([[kategori.trim(), tanya, a, b, c, d, e, kunci, gambar || ""]]);
  return { status: "sukses" };
}

function hapusSoal(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Soal");
  if (!sheet || row < 2) return { status: "gagal" };
  sheet.deleteRow(row);
  return { status: "sukses" };
}
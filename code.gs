function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || "presensi"; 

    // --- 1. LOGIKA PRESENSI ---
    if (action === "presensi") {
      var username = data.username || "Username Kosong";
      var nama = data.nama || "Nama Kosong";
      var jenis = data.jenis || "masuk";
      var fotoBase64 = data.foto || "";
      var lokasi = data.lokasi || "Tidak Diketahui";

      // Membentuk format dd/mm/yyyy dan hh.mm
      var dateObj = new Date();
      var dd = ("0" + dateObj.getDate()).slice(-2);
      var mm = ("0" + (dateObj.getMonth() + 1)).slice(-2);
      var yyyy = dateObj.getFullYear();
      
      // TAMBAHAN: Menyisipkan tanda petik tunggal (') di depan variabel
      var tanggal = "'" + dd + "/" + mm + "/" + yyyy;

      var hh = ("0" + dateObj.getHours()).slice(-2);
      var mnt = ("0" + dateObj.getMinutes()).slice(-2);
      
      // TAMBAHAN: Menyisipkan tanda petik tunggal (') di depan variabel
      var jam = "'" + hh + "." + mnt;

      var base64Data = fotoBase64.split(",")[1]; 
      var fileName = "Absen_" + jenis + "_" + username + "_" + dateObj.getTime() + ".jpg";
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', fileName);

      // MASUKKAN ID FOLDER GOOGLE DRIVE ANDA DI SINI
      var folderId = "10BORea1rdDCs8SBeZCmZUsssrtSY_J27"; 
      var folder = DriveApp.getFolderById(folderId);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataAbsen");
      // Sesuai permintaan: Nama, tanggal, jam, jenis, lokasi pos, link foto
      sheet.appendRow([nama, tanggal, jam, jenis, lokasi, file.getUrl()]);

      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // --- 2. LOGIKA LOGIN ---
    else if (action === "login") {
      var username = data.username;
      var password = data.password;
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSatpam");
      var rows = sheet.getDataRange().getValues();
      
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] == username) {
          if (rows[i][2] == password) {
            return ContentService.createTextOutput(JSON.stringify({ status: "success", nama: rows[i][1] })).setMimeType(ContentService.MimeType.JSON);
          } else {
            return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Password salah!" })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Username tidak ditemukan!" })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- 3. LOGIKA MANAJEMEN PEGAWAI & WAJAH ---
    else if (action === "add_user") {
      // 1. Ambil data dan paksa username jadi huruf kecil & buang spasi ujungnya
      var username = (data.username || "").toString().trim().toLowerCase();
      var nama = (data.nama || "").toString().trim();
      var password = (data.password || "").toString();

      // 2. Cek kelengkapan data
      if (!username || !nama || !password) {
         return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Data tidak lengkap!"})).setMimeType(ContentService.MimeType.JSON);
      }
      
      // 3. Cek larangan pakai nama admin
      if (username === "admin") {
         return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Username 'admin' dilarang!"})).setMimeType(ContentService.MimeType.JSON);
      }
      
      // 4. Cek spasi di tengah username
      if (/\s/.test(username)) {
         return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Username tidak boleh ada spasi!"})).setMimeType(ContentService.MimeType.JSON);
      }

      // 5. Cek panjang password
      if (password.length < 6) {
         return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Password minimal 6 karakter!"})).setMimeType(ContentService.MimeType.JSON);
      }

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSatpam");
      var rows = sheet.getDataRange().getValues();
      
      // 6. Cek apakah Username sudah pernah dipakai orang lain
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString().toLowerCase() === username) {
          return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Username sudah terpakai! Pilih yang lain."})).setMimeType(ContentService.MimeType.JSON);
        }
      }

      // 7. Jika lolos semua ujian di atas, baru simpan ke Google Sheets
      sheet.appendRow([username, nama, password, ""]);
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "delete_user") {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSatpam");
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.username) { sheet.deleteRow(i + 1); break; }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "simpan_wajah") {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSatpam");
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.username) {
          sheet.getRange(i + 1, 4).setValue(data.descriptor);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    // --- 4. LOGIKA MANAJEMEN LOKASI ---
    else if (action === "add_location") {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataLokasi");
      sheet.appendRow([new Date().getTime(), data.nama_lokasi, data.latitude, data.longitude, data.radius]);
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "delete_location") {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataLokasi");
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.id) { sheet.deleteRow(i + 1); break; }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "update_radius") {
      var id = data.id;
      var radius = data.radius;
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataLokasi"); // Catatan: Sesuaikan dengan nama sheet tempat Anda menyimpan data koordinat pos
      var rows = sheet.getDataRange().getValues();
      
      for (var i = 1; i < rows.length; i++) {
        // Mencari baris yang ID-nya cocok
        if (rows[i][0].toString() === id.toString()) {
          // Angka 5 di bawah ini berarti kolom ke-5 (Radius). 
          // Jika di Google Sheets Anda posisi kolom Radius berbeda (misal kolom ke-4), ganti angka 5 menjadi urutan kolom Anda.
          sheet.getRange(i + 1, 5).setValue(radius); 
          return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "ID Pos tidak ditemukan"})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var action = e.parameter.action;

  if (action === "get_descriptor") {
    var username = e.parameter.username;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] == username) {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", descriptor: rows[i][3] })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Wajah belum terdaftar" })).setMimeType(ContentService.MimeType.JSON);
  }
  else if (action === "get_users") {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    var users = [];
    for (var i = 1; i < rows.length; i++) {
      users.push({ username: rows[i][0], nama: rows[i][1], status_wajah: (rows[i][3] && rows[i][3] !== "") ? "Sudah" : "Belum" });
    }
    return ContentService.createTextOutput(JSON.stringify(users)).setMimeType(ContentService.MimeType.JSON);
  }
  else if (action === "get_locations") {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataLokasi");
    var rows = sheet.getDataRange().getValues();
    var locs = [];
    for (var i = 1; i < rows.length; i++) {
      locs.push({ id: rows[i][0], nama_lokasi: rows[i][1], latitude: rows[i][2], longitude: rows[i][3], radius: rows[i][4] });
    }
    return ContentService.createTextOutput(JSON.stringify(locs)).setMimeType(ContentService.MimeType.JSON);
  }
  else if (e.parameter.action === "get_riwayat_user") {
    var usernameReq = e.parameter.username;
    
    // Ganti "DataPresensi" dengan nama tab sheet tempat absen tersimpan
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataPresensi"); 
    var rows = sheet.getDataRange().getValues();
    var riwayat = [];
    
    // Looping dari baris paling bawah (data paling baru) ke atas
    for (var i = rows.length - 1; i > 0; i--) {
      
      // ASUMSI KOLOM (Sesuaikan dengan urutan kolom di Google Sheets Anda)
      // Array dimulai dari 0. Jika Username ada di kolom B, berarti index-nya 1.
      var unameSheet = rows[i][1]; // Kolom B (Username)
      
      if (unameSheet === usernameReq) {
        riwayat.push({
          tanggal: rows[i][0], // Kolom A (Tanggal)
          masuk: rows[i][3],   // Kolom D (Jam Masuk)
          pulang: rows[i][4],  // Kolom E (Jam Pulang)
          lokasi: rows[i][5],  // Kolom F (Lokasi Pos)
          status: rows[i][7] || "Hadir" // Kolom H (Status Tepat Waktu/Terlambat)
        });
      }
      
      // Batasi hanya mengirim 30 data riwayat terakhir agar tidak berat
      if (riwayat.length >= 30) break; 
    }
    
    // Kirim data kembali ke frontend
    return ContentService.createTextOutput(JSON.stringify(riwayat)).setMimeType(ContentService.MimeType.JSON);
  }
  else {
    // Membaca status untuk Dashboard (dicari berdasarkan Nama)
    var nama = e.parameter.nama;
    var responseData = { masuk: "--.-- WIB", keluar: "--.-- WIB", kode_status: 0 };
    if (!nama) return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataAbsen");
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);

    var sekarang = new Date();
    var batas18Jam = new Date(sekarang.getTime() - (18 * 60 * 60 * 1000)); 
    var waktuMasukTerakhir = null;

    // Cari Jam Masuk
    for (var i = data.length - 1; i >= 1; i--) {
      var rowNama = data[i][0], rowTanggal = data[i][1], rowJam = data[i][2], rowJenis = data[i][3];
      
      if (rowNama == nama && rowJenis == 'masuk') {
        var parts = rowTanggal.split("/");
        var timeParts = rowJam.split(".");
        var tgl = new Date(parts[2], parts[1]-1, parts[0], timeParts[0], timeParts[1]);
        
        if (tgl >= batas18Jam) {
          waktuMasukTerakhir = tgl;
          responseData.masuk = rowJam + " WIB";
          responseData.kode_status = 1; 
          break; 
        }
      }
    }

    // Cari Jam Keluar
    if (waktuMasukTerakhir) {
      for (var j = data.length - 1; j >= 1; j--) {
        var rNama = data[j][0], rTgl = data[j][1], rJam = data[j][2], rJns = data[j][3];
        if (rNama == nama && rJns == 'keluar') {
          var p2 = rTgl.split("/");
          var t2 = rJam.split(".");
          var tglKeluar = new Date(p2[2], p2[1]-1, p2[0], t2[0], t2[1]);
          
          if (tglKeluar > waktuMasukTerakhir) {
            responseData.keluar = rJam + " WIB";
            responseData.kode_status = 2; 
            break; 
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);
  }
}

function doOptions(e) { return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT).setHeaders({ "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, GET, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" }); }
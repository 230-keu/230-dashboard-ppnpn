function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e.parameter.action;
  
  // 1. DATA PEGAWAI UNTUK ADMIN
  if (action === "get_users") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    var users = [];
    for (var i = 1; i < rows.length; i++) {
      // Jika kolom Status DNA kosong tapi DNA ada, anggap "Disetujui" (untuk user lama)
      var statusDNA = rows[i][4] ? rows[i][4].toString() : (rows[i][3] ? "Disetujui" : "Belum");
      users.push({
        username: rows[i][0],
        nama: rows[i][1],
        status_wajah: statusDNA,
        foto: rows[i][5] || ""
      });
    }
    return ContentService.createTextOutput(JSON.stringify(users)).setMimeType(ContentService.MimeType.JSON);
  }
  
  else if (action === "get_locations") {
    var sheet = ss.getSheetByName("DataLokasi");
    var rows = sheet.getDataRange().getValues();
    var locs = [];
    for (var i = 1; i < rows.length; i++) {
      locs.push({ id: rows[i][0], nama_lokasi: rows[i][1], latitude: rows[i][2], longitude: rows[i][3], radius: rows[i][4] });
    }
    return ContentService.createTextOutput(JSON.stringify(locs)).setMimeType(ContentService.MimeType.JSON);
  }
  
  else if (action === "get_descriptor") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === e.parameter.username) {
        var statusDNA = rows[i][4] ? rows[i][4].toString() : (rows[i][3] ? "Disetujui" : "Belum");
        if (statusDNA === "Disetujui") {
          return ContentService.createTextOutput(JSON.stringify({status: "success", descriptor: rows[i][3]})).setMimeType(ContentService.MimeType.JSON);
        } else {
          return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Wajah belum disetujui Admin!"})).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  else if (action === "get_riwayat_user") {
    var uname = e.parameter.username;
    var sheet = ss.getSheetByName("DataPresensi");
    var rows = sheet.getDataRange().getValues();
    var riwayat = [];
    
    for (var i = rows.length - 1; i > 0; i--) {
      if (rows[i][1].toString() === uname) {
        var masukTxt = rows[i][3] ? rows[i][3].toString() : "";
        var pulangTxt = rows[i][4] ? rows[i][4].toString() : "";
        
        var statusAbsen = "Hadir";
        var isTL = false;
        var isPSW = false;
        var isShiftPagi = false;
        var isShiftMalam = false;
        
        if (masukTxt !== "") {
          var jamM = parseInt(masukTxt.split(".")[0]);
          var mntM = parseInt(masukTxt.split(".")[1]);
          
          if (jamM >= 4 && jamM < 15) {
            isShiftPagi = true;
            if ((jamM === 7 && mntM > 0) || jamM > 7) isTL = true;
          }
          else if (jamM >= 15 || jamM < 4) {
            isShiftMalam = true;
            if ((jamM === 19 && mntM > 0) || jamM > 19 || jamM < 4) isTL = true;
          }
          
          if (pulangTxt === "") {
            isPSW = true;
            statusAbsen = "PSW (Lupa Absen)";
          } else {
            var jamK = parseInt(pulangTxt.split(".")[0]);
            if (isShiftPagi) { if (jamK >= 4 && jamK < 19) isPSW = true; } 
            else if (isShiftMalam) { if (jamK >= 15 || jamK < 7) isPSW = true; }
          }
        }
        
        if (isTL && isPSW && pulangTxt !== "") statusAbsen = "TL & PSW";
        else if (isTL) statusAbsen = "TL";
        else if (isPSW && pulangTxt !== "") statusAbsen = "PSW";
        else if (masukTxt === "") statusAbsen = "Kosong";

        riwayat.push({ tanggal: rows[i][0], masuk: masukTxt, pulang: pulangTxt, lokasi: rows[i][5], status: statusAbsen });
      }
      if (riwayat.length >= 30) break; 
    }
    return ContentService.createTextOutput(JSON.stringify(riwayat)).setMimeType(ContentService.MimeType.JSON);
  }

  // --- CEK STATUS PRESENSI & STATUS DNA WAJAH ---
  else if (e.parameter.nama && e.parameter.username) {
    var nama = e.parameter.nama;
    var uname = e.parameter.username;
    
    // 1. Cek Status Wajah dari DataSatpam
    var sheetSatpam = ss.getSheetByName("DataSatpam");
    var rowsSatpam = sheetSatpam.getDataRange().getValues();
    var stat_dna = "Belum";
    for(var s=1; s<rowsSatpam.length; s++) {
      if(rowsSatpam[s][0].toString() === uname) {
        stat_dna = rowsSatpam[s][4] ? rowsSatpam[s][4].toString() : (rowsSatpam[s][3] ? "Disetujui" : "Belum");
        break;
      }
    }

    // 2. Cek Status Presensi dari DataPresensi
    var sheet = ss.getSheetByName("DataPresensi");
    var rows = sheet.getDataRange().getValues();
    var status_kode = 0; 
    var jam_masuk = "--.-- WIB";
    var jam_keluar = "--.-- WIB";
    var lokasi_pos_masuk = ""; 
    
    var lastRowIdx = -1;
    for (var i = rows.length - 1; i > 0; i--) {
      if (rows[i][2].toString() === nama) { lastRowIdx = i; break; }
    }
    
    if (lastRowIdx !== -1) {
      var row = rows[lastRowIdx];
      var nowMs = new Date().getTime(); 
      if (!row[4] || row[4].toString().trim() === "") {
        var msSejakMasuk = nowMs - parseFloat(row[9]); 
        var jamSejakMasuk = msSejakMasuk / (1000 * 60 * 60);
        if (jamSejakMasuk < 16) {
          status_kode = 1; 
          jam_masuk = row[3] + " WIB";
          lokasi_pos_masuk = row[5] ? row[5].toString() : ""; 
        } else { status_kode = 0; }
      } 
      else {
        var msSejakKeluar = nowMs - parseFloat(row[10]); 
        var jamSejakKeluar = msSejakKeluar / (1000 * 60 * 60);
        if (jamSejakKeluar < 8) {
          status_kode = 2; 
          jam_masuk = row[3] + " WIB";
          jam_keluar = row[4] + " WIB";
          lokasi_pos_masuk = row[5] ? row[5].toString() : ""; 
        } else { status_kode = 0; }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      kode_status: status_kode, masuk: jam_masuk, keluar: jam_keluar,
      lokasi_masuk: lokasi_pos_masuk, status_dna: stat_dna // Mengirim Status DNA
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === "login") {
    // [Kode Login sama seperti sebelumnya]
    var usr = data.username.toLowerCase().trim();
    var pwd = data.password;
    if (usr === "admin" && pwd === "admin123") { return ContentService.createTextOutput(JSON.stringify({status: "success", nama: "Administrator"})).setMimeType(ContentService.MimeType.JSON); }
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().toLowerCase() === usr && rows[i][2].toString() === pwd) { return ContentService.createTextOutput(JSON.stringify({status: "success", nama: rows[i][1]})).setMimeType(ContentService.MimeType.JSON); }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Username/Password salah!"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  else if (action === "add_user") {
    // [Kode Add User sama, tapi pastikan status awal "Belum"]
    var usr = data.username.toLowerCase().trim();
    if (/\s/.test(usr)) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Spasi dilarang!"})).setMimeType(ContentService.MimeType.JSON);
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().toLowerCase() === usr) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Username terpakai!"})).setMimeType(ContentService.MimeType.JSON);
    }
    sheet.appendRow([usr, data.nama, data.password, "", "Belum", ""]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  else if (action === "delete_user") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) { sheet.deleteRow(i + 1); return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON); }
    }
  }
  
  // --- MANAJEMEN LOKASI ---
  else if (action === "add_location") {
    var sheet = ss.getSheetByName("DataLokasi");
    sheet.appendRow(["POS" + new Date().getTime(), data.nama_lokasi, data.latitude, data.longitude, data.radius]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
  else if (action === "update_radius") {
    var sheet = ss.getSheetByName("DataLokasi");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) { sheet.getRange(i + 1, 5).setValue(data.radius); return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON); }
    }
  }
  else if (action === "delete_location") {
    var sheet = ss.getSheetByName("DataLokasi");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) { sheet.deleteRow(i + 1); return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON); }
    }
  }

  // --- REKAM WAJAH MANUAL OLEH ADMIN ---
  else if (action === "simpan_wajah") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 4).setValue(data.descriptor); // DNA
        sheet.getRange(i + 1, 5).setValue("Disetujui");     // Otomatis Disetujui
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  // --- FITUR BARU: PENGAJUAN WAJAH MANDIRI OLEH PEGAWAI ---
  else if (action === "ajukan_wajah_mandiri") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    var fotoUrl = "";
    
    // Upload Foto Verifikasi Wajah ke Drive
    try {
      if (data.foto && data.foto !== "") {
        var base64Data = data.foto.split(",")[1]; 
        var fileName = "Verifikasi_" + data.username + "_" + new Date().getTime() + ".jpg";
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', fileName);
        var folder = DriveApp.getFolderById("10BORea1rdDCs8SBeZCmZUsssrtSY_J27"); // Folder Drive Anda
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoUrl = file.getUrl();
      }
    } catch(e) { }

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 4).setValue(data.descriptor); // Simpan DNA Sementara
        sheet.getRange(i + 1, 5).setValue("Menunggu");      // Set Status Menunggu ACC
        sheet.getRange(i + 1, 6).setValue(fotoUrl);         // Simpan Link Foto
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan di Database"})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- FITUR BARU: ADMIN ACC / TOLAK WAJAH ---
  else if (action === "acc_wajah") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 5).setValue("Disetujui");
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  else if (action === "tolak_wajah") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 4).clearContent(); // Hapus DNA yang salah
        sheet.getRange(i + 1, 5).setValue("Ditolak"); // Set Status Ditolak
        sheet.getRange(i + 1, 6).clearContent(); // Hapus Link Foto
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  // --- LOGIKA PRESENSI ---
  else if (action === "presensi") {
    var sheet = ss.getSheetByName("DataPresensi");
    var now = new Date();
    var tanggal_wib = Utilities.formatDate(now, "GMT+7", "dd-MM-yyyy");
    var jam_wib = Utilities.formatDate(now, "GMT+7", "HH.mm");
    var time_milidetik = now.getTime();
    
    var fotoUrl = "";
    try {
      if (data.foto && data.foto !== "") {
        var base64Data = data.foto.split(",")[1]; 
        var fileName = "Absen_" + data.jenis + "_" + data.username + "_" + time_milidetik + ".jpg";
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', fileName);
        var folder = DriveApp.getFolderById("10BORea1rdDCs8SBeZCmZUsssrtSY_J27");
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoUrl = file.getUrl();
      }
    } catch(e) { }

    if (data.jenis === "masuk") {
      sheet.appendRow([ tanggal_wib, data.username, data.nama, jam_wib, "", data.lokasi, "", fotoUrl, "", time_milidetik, "" ]);
    } 
    else if (data.jenis === "keluar") {
      var rows = sheet.getDataRange().getValues();
      for (var i = rows.length - 1; i > 0; i--) {
        if (rows[i][1].toString() === data.username) {
          var tsMasuk = parseFloat(rows[i][9]); 
          var selisihJam = (time_milidetik - tsMasuk) / (1000 * 60 * 60);
          if (selisihJam < 16) {
            sheet.getRange(i + 1, 5).setValue(jam_wib);        
            sheet.getRange(i + 1, 7).setValue(data.lokasi);    
            if (fotoUrl !== "") { sheet.getRange(i + 1, 9).setValue(fotoUrl); }
            sheet.getRange(i + 1, 11).setValue(time_milidetik);
          }
          break; 
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
}
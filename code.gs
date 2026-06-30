// --- FUNGSI PENYAPU OTOMATIS GLOBAL (Merapikan 9.2 menjadi 09.20) ---
function formatJam(teks) {
  if (!teks || teks.toString().trim() === "") return "";
  var parts = teks.toString().split(".");
  var j = parts[0].padStart(2, '0');
  var m = (parts[1] || "00").padEnd(2, '0');
  return j + "." + m;
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e.parameter.action;

  // 1. DATA PEGAWAI UNTUK ADMIN
  if (action === "get_users") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    var users = [];
    for (var i = 1; i < rows.length; i++) {
      var statusDNA = rows[i][4] ? rows[i][4].toString() : (rows[i][3] ? "Disetujui" : "Belum");
      users.push({
        username: rows[i][0],
        nama: rows[i][1],
        status_wajah: statusDNA,
        foto: rows[i][5] || "",
        jatah_tahunan: rows[i][6] !== undefined && rows[i][6] !== "" ? rows[i][6] : 6,
        periode: rows[i][8] || "Jan-Des",
        peran: rows[i][9] || "Satpam"
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

  // --- KALKULASI RIWAYAT LOG USER (DASHBOARD MOBIL) ---
  else if (action === "get_riwayat_user") {
    var uname = e.parameter.username;
    var sheetSatpam = ss.getSheetByName("DataSatpam");
    var sData = sheetSatpam ? sheetSatpam.getDataRange().getValues() : [];
    var peran = "Satpam";
    for(var s=1; s<sData.length; s++) {
      if(sData[s][0].toString() === uname) {
        peran = sData[s][9] ? sData[s][9].toString() : "Satpam"; break;
      }
    }

    var sheet = ss.getSheetByName("DataPresensi");
    var rows = sheet.getDataRange().getValues();
    var riwayat = [];
    
    for (var i = rows.length - 1; i > 0; i--) {
      if (rows[i][1].toString() === uname) {
        var masukTxt = rows[i][3] ? formatJam(rows[i][3]) : "";
        var pulangTxt = rows[i][4] ? formatJam(rows[i][4]) : "";
        
        var statusAbsen = "Hadir";
        var isTL = false;
        var isPSW = false;
        var isShiftPagi = false;
        var isShiftMalam = false;
        
        if (masukTxt !== "") {
          var jamM = parseInt(masukTxt.split(".")[0]);
          var mntM = parseInt(masukTxt.split(".")[1]);
          var wkMasuk = jamM * 60 + mntM;
          
          if (peran === "Satpam") {
            if (jamM >= 4 && jamM < 15) {
              isShiftPagi = true;
              if (wkMasuk > 7 * 60) isTL = true;
            }
            else if (jamM >= 15 || jamM < 4) {
              isShiftMalam = true;
              var wktMalam = wkMasuk; if (jamM < 4) wktMalam += 24 * 60;
              if (wktMalam > 19 * 60) isTL = true;
            }
            if (pulangTxt === "") {
              isPSW = true; statusAbsen = "Lupa Absen";
            } else {
              var jamK = parseInt(pulangTxt.split(".")[0]);
              if (isShiftPagi) { if (jamK >= 4 && jamK < 19) isPSW = true; } 
              else if (isShiftMalam) { if (jamK >= 15 || jamK < 7) isPSW = true; }
            }
          } else {
            var batasM = (peran === "Petugas Kebersihan" || peran === "Pengemudi") ? (6 * 60 + 30) : (7 * 60);
            var batasK = 17 * 60;
            if (wkMasuk > batasM) isTL = true;
            if (pulangTxt === "") {
              isPSW = true; statusAbsen = "Lupa Absen";
            } else {
              var jamK = parseInt(pulangTxt.split(".")[0]);
              var mntK = parseInt(pulangTxt.split(".")[1]);
              var wkKeluar = jamK * 60 + mntK;
              if (wkKeluar < batasK) isPSW = true;
            }
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
    
    var sheetSatpam = ss.getSheetByName("DataSatpam");
    var rowsSatpam = sheetSatpam.getDataRange().getValues();
    var stat_dna = "Belum";
    var user_found = false;
    var peran = "Satpam";
    for(var s=1; s<rowsSatpam.length; s++) {
      if(rowsSatpam[s][0].toString() === uname) {
        stat_dna = rowsSatpam[s][4] ? rowsSatpam[s][4].toString() : (rowsSatpam[s][3] ? "Disetujui" : "Belum");
        peran = rowsSatpam[s][9] ? rowsSatpam[s][9].toString() : "Satpam";
        user_found = true;
        break;
      }
    }
    
    if (!user_found) {
      return ContentService.createTextOutput(JSON.stringify({ status_user: "not_found" })).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = ss.getSheetByName("DataPresensi");
    var rows = sheet.getDataRange().getValues();
    
    var status_kode = 0; 
    var jam_masuk = "--.-- WIB";
    var jam_keluar = "--.-- WIB";
    var lokasi_pos_masuk = ""; 
    var menit_telat = 0;
    var durasi_jam = 0;
    
    var lastRowIdx = -1;
    for (var i = rows.length - 1; i > 0; i--) {
      if (rows[i][2].toString() === nama) { lastRowIdx = i; break; }
    }
    
    if (lastRowIdx !== -1) {
      var row = rows[lastRowIdx];
      var nowMs = new Date().getTime(); 
      var wktMasukBersih = formatJam(row[3]);

      if (wktMasukBersih !== "") {
        var jamM = parseInt(wktMasukBersih.split(".")[0]);
        var mntM = parseInt(wktMasukBersih.split(".")[1]);
        
        var batasPagi = 7 * 60;
        if (peran === "Petugas Kebersihan" || peran === "Pengemudi") batasPagi = 6 * 60 + 30; // 06.30

        if (peran === "Satpam") {
          if (jamM >= 4 && jamM < 15) { 
            var waktuMasuk = jamM * 60 + mntM;
            if (waktuMasuk > batasPagi) menit_telat = waktuMasuk - batasPagi;
          } else if (jamM >= 15 || jamM < 4) { 
            var batasMalam = 19 * 60;
            var waktuMasuk = jamM * 60 + mntM;
            if (jamM < 4) waktuMasuk += 24 * 60; 
            if (waktuMasuk > batasMalam) menit_telat = waktuMasuk - batasMalam;
          }
        } else {
          var waktuMasuk = jamM * 60 + mntM;
          if (waktuMasuk > batasPagi) menit_telat = waktuMasuk - batasPagi;
        }
      }

      if (!row[4] || row[4].toString().trim() === "") {
        var msSejakMasuk = nowMs - parseFloat(row[9]);
        var jamSejakMasuk = msSejakMasuk / (1000 * 60 * 60);
        if (jamSejakMasuk < 16) {
          status_kode = 1;
          jam_masuk = wktMasukBersih + " WIB";
          lokasi_pos_masuk = row[5] ? row[5].toString() : ""; 
        } else { status_kode = 0; }
      } 
      else {
        var msSejakKeluar = nowMs - parseFloat(row[10]);
        var jamSejakKeluar = msSejakKeluar / (1000 * 60 * 60);
        if (jamSejakKeluar < 8) {
          status_kode = 2;
          jam_masuk = wktMasukBersih + " WIB";
          jam_keluar = formatJam(row[4]) + " WIB";
          lokasi_pos_masuk = row[5] ? row[5].toString() : "";
          durasi_jam = (parseFloat(row[10]) - parseFloat(row[9])) / (1000 * 60 * 60);
        } else { status_kode = 0; }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      kode_status: status_kode, masuk: jam_masuk, keluar: jam_keluar,
      lokasi_masuk: lokasi_pos_masuk, status_dna: stat_dna, telat: menit_telat, durasi: durasi_jam     
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === "login") {
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
    var usr = data.username.toLowerCase().trim();
    if (/\s/.test(usr)) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Spasi dilarang!"})).setMimeType(ContentService.MimeType.JSON);
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().toLowerCase() === usr) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Username terpakai!"})).setMimeType(ContentService.MimeType.JSON);
    }
    var yr = new Date().getFullYear();
    var defPer = yr + "-01-01|" + yr + "-12-31";
    var peran = data.peran || "Satpam";
    sheet.appendRow([usr, data.nama, data.password, "", "Belum", "", 6, "", defPer, peran]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  else if (action === "delete_user") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) { sheet.deleteRow(i + 1); return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON); }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }
  
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
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Lokasi tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }
  else if (action === "delete_location") {
    var sheet = ss.getSheetByName("DataLokasi");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) { sheet.deleteRow(i + 1); return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON); }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Lokasi tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "simpan_wajah") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 4).setValue(data.descriptor); 
        sheet.getRange(i + 1, 5).setValue("Disetujui");     
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "ajukan_wajah_mandiri") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    var fotoUrl = "";
    try {
      if (data.foto && data.foto !== "") {
        var base64Data = data.foto.split(",")[1]; 
        var fileName = "Verifikasi_" + data.username + "_" + new Date().getTime() + ".jpg";
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', fileName);
        var folder = DriveApp.getFolderById("10BORea1rdDCs8SBeZCmZUsssrtSY_J27"); 
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoUrl = file.getUrl();
      }
    } catch(e) { }

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 4).setValue(data.descriptor); 
        sheet.getRange(i + 1, 5).setValue("Menunggu");      
        sheet.getRange(i + 1, 6).setValue(fotoUrl);         
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan di Database"})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "acc_wajah") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 5).setValue("Disetujui");
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  else if (action === "tolak_wajah") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 4).clearContent(); 
        sheet.getRange(i + 1, 5).setValue("Ditolak"); 
        sheet.getRange(i + 1, 6).clearContent(); 
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "presensi") {
    var sheet = ss.getSheetByName("DataPresensi");
    var now = new Date();
    var tanggal_wib = Utilities.formatDate(now, "GMT+7", "dd-MM-yyyy");
    var jam_wib = "'" + Utilities.formatDate(now, "GMT+7", "HH.mm");
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

  // === FITUR CUTI ===
  else if (action === "ajukan_cuti") {
    var sheet = ss.getSheetByName("DataCuti");
    if(!sheet) {
      sheet = ss.insertSheet("DataCuti");
      sheet.appendRow(["Timestamp", "Username", "Nama", "Jenis Cuti", "Tgl Mulai", "Tgl Selesai", "Jml Hari", "Alasan", "Lampiran", "Status", "Catatan Admin"]);
    }
    
    var time = new Date().getTime();
    var fotoUrl = "";
    if (data.lampiran && data.lampiran !== "") {
      try {
        var base64Data = data.lampiran.split(",")[1]; 
        var fileName = "Surat_" + data.jenis_cuti.replace(/\s/g, "") + "_" + data.username + "_" + time + ".jpg";
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', fileName);
        var folder = DriveApp.getFolderById("10BORea1rdDCs8SBeZCmZUsssrtSY_J27"); 
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoUrl = file.getUrl();
      } catch(e) {}
    }
    
    sheet.appendRow([
      Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy HH:mm:ss"), data.username, data.nama, data.jenis_cuti, data.tgl_mulai, data.tgl_selesai, data.jml_hari, data.alasan, fotoUrl, "Menunggu", ""
    ]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "get_cuti_user") {
    var sheetUser = ss.getSheetByName("DataSatpam");
    var yr = new Date().getFullYear();
    var jatahTahunan = 12; var periode = yr + "-01-01|" + yr + "-12-31";
    if (sheetUser) {
      var rowsUser = sheetUser.getDataRange().getValues();
      for (var i = 1; i < rowsUser.length; i++) {
        if (rowsUser[i][0].toString() === data.username) {
          jatahTahunan = rowsUser[i][6] !== undefined && rowsUser[i][6] !== "" ? parseInt(rowsUser[i][6]) : 12;
          periode = rowsUser[i][8] !== undefined && rowsUser[i][8] !== "" ? rowsUser[i][8] : (yr + "-01-01|" + yr + "-12-31");
          break;
        }
      }
    }

    var sheetCuti = ss.getSheetByName("DataCuti");
    var riwayat = [];
    var terpakaiTahunan = 0;
    
    if (sheetCuti) {
      var rows = sheetCuti.getDataRange().getValues();
      for (var i = rows.length - 1; i > 0; i--) {
        if (rows[i][1].toString() === data.username) {
          riwayat.push({
            id: i, tgl_pengajuan: rows[i][0], jenis: rows[i][3], mulai: rows[i][4], selesai: rows[i][5],
            hari: rows[i][6], alasan: rows[i][7], lampiran: rows[i][8], status: rows[i][9], catatan: rows[i][10]
          });
          if (rows[i][9] === "Disetujui") {
            if (rows[i][3] === "Cuti Tahunan") terpakaiTahunan += parseInt(rows[i][6] || 0);
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({
      status: "success", riwayat: riwayat, jatah_tahunan: jatahTahunan, sisa_tahunan: jatahTahunan - terpakaiTahunan, periode: periode
    })).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "get_cuti_all") {
    var sheetCuti = ss.getSheetByName("DataCuti");
    var list = [];
    if (sheetCuti) {
      var rows = sheetCuti.getDataRange().getValues();
      for (var i = rows.length - 1; i > 0; i--) {
        list.push({
          id: i, tgl_pengajuan: rows[i][0], username: rows[i][1], nama: rows[i][2],
          jenis: rows[i][3], mulai: rows[i][4], selesai: rows[i][5], hari: rows[i][6],
          alasan: rows[i][7], lampiran: rows[i][8], status: rows[i][9], catatan: rows[i][10]
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success", data: list})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "aksi_cuti") {
    var sheet = ss.getSheetByName("DataCuti");
    if(sheet) {
      sheet.getRange(data.id + 1, 10).setValue(data.keputusan); 
      sheet.getRange(data.id + 1, 11).setValue(data.catatan);
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Database Cuti tidak ditemukan"})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "update_jatah") {
    var sheet = ss.getSheetByName("DataSatpam");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.username) {
        sheet.getRange(i + 1, 7).setValue(data.tahunan);
        sheet.getRange(i + 1, 9).setValue(data.periode);
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "User tidak ditemukan!"})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- REKAPITULASI DENGAN INTEGRASI CUTI ---
  else if (action === "get_jadwal_bulanan") {
    var sheet = ss.getSheetByName("DataJadwal");
    if (!sheet) {
      sheet = ss.insertSheet("DataJadwal");
      sheet.appendRow(["BulanTahun", "TIM", "Nama", "1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"]);
      // Baris Contoh agar User paham
      sheet.appendRow(["2024-04", "TIM 1", "DEDDY PURWANTO", "MG", "P", "M", "", "P", "M", "PR", "MR", "PS", "MS", "PG", "MG", "", "P", "M", "P", "M", "PR", "MR", "PS", "MS", "PG", "MG", "", "P", "M", "P", "M", "PR", "MR", "PS"]);
    }
    var rows = sheet.getDataRange().getValues();
    var jadwal = [];
    var blnThn = data.bulanTahun;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] && rows[i][0].toString() === blnThn) {
        var jRow = { tim: rows[i][1], nama: rows[i][2], shifts: {} };
        for(var d=1; d<=31; d++) {
          jRow.shifts[d] = rows[i][d+2] ? rows[i][d+2].toString().trim() : "";
        }
        jadwal.push(jRow);
      }
    }
    
    var presensi = [];
    if(data.bandingkan) {
      var sheetP = ss.getSheetByName("DataPresensi");
      var pRows = sheetP ? sheetP.getDataRange().getValues() : [];
      var mPart = blnThn.split("-")[1];
      var yPart = blnThn.split("-")[0];
      var suffix = "-" + mPart + "-" + yPart;
      
      for(var i=1; i<pRows.length; i++) {
        var tglP = pRows[i][0];
        var tglStr = (tglP instanceof Date) ? Utilities.formatDate(tglP, "GMT+7", "dd-MM-yyyy") : tglP.toString().trim();
        if(tglStr.endsWith(suffix)) {
           presensi.push({
             tgl: tglStr,
             nama: pRows[i][2] ? pRows[i][2].toString() : "",
             username: pRows[i][1] ? pRows[i][1].toString() : "",
             masuk: pRows[i][3] ? true : false
           });
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "success", jadwal: jadwal, presensi: presensi})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "get_rekap_data") {
    try {
      var uname = data.username;
      var isShift = data.isShift === true;
      var tglAwal = new Date(data.tglAwal);
      var tglAkhir = new Date(data.tglAkhir);
      
      var sheetPresensi = ss.getSheetByName("DataPresensi");
      var rowsPresensi = sheetPresensi ? sheetPresensi.getDataRange().getValues() : [];
      
      var sheetCuti = ss.getSheetByName("DataCuti");
      var rowsCuti = sheetCuti ? sheetCuti.getDataRange().getValues() : [];
      
      var namaSatpam = "";
      var peran = "Satpam";
      var sheetSatpam = ss.getSheetByName("DataSatpam");
      if(sheetSatpam) {
          var sData = sheetSatpam.getDataRange().getValues();
          for(var i=1; i<sData.length; i++) {
              if(sData[i][0].toString() === uname) { 
                  namaSatpam = sData[i][1].toString().toUpperCase(); 
                  peran = sData[i][9] ? sData[i][9].toString() : "Satpam";
                  break; 
              }
          }
      }
      
      var sheetJadwal = ss.getSheetByName("DataJadwal");
      var rowsJadwal = sheetJadwal ? sheetJadwal.getDataRange().getValues() : [];
      
      var cutiDisetujui = []; 
      for(var i = 1; i < rowsCuti.length; i++) {
        if(rowsCuti[i][1].toString() === uname && rowsCuti[i][9] === "Disetujui") {
          var cMulai = new Date(rowsCuti[i][4]);
          var cSelesai = new Date(rowsCuti[i][5]);
          var jenis = rowsCuti[i][3];
          for (var d = new Date(cMulai); d <= cSelesai; d.setDate(d.getDate() + 1)) {
            cutiDisetujui.push({ tgl: Utilities.formatDate(d, "GMT+7", "dd-MM-yyyy"), jenis: jenis });
          }
        }
      }

      var hasil = [];
      var sum = { tHari: 0, hHadir: 0, hSanksi: 0, tPot: 0 };
      
      for (var d = new Date(tglAwal); d <= tglAkhir; d.setDate(d.getDate() + 1)) {
        var tglStr = Utilities.formatDate(d, "GMT+7", "dd-MM-yyyy");
        var dayOfWeek = d.getDay();
        var blnThn = Utilities.formatDate(d, "GMT+7", "yyyy-MM");
        var dayOfMonth = d.getDate();
        
        var shiftTerjadwal = null; 
        var adaJadwalBulanIni = false;
        
        for (var j = 1; j < rowsJadwal.length; j++) {
            if(rowsJadwal[j][0] && rowsJadwal[j][0].toString() === blnThn) {
                adaJadwalBulanIni = true;
                if(rowsJadwal[j][2] && rowsJadwal[j][2].toString().toUpperCase() === namaSatpam) {
                    shiftTerjadwal = rowsJadwal[j][dayOfMonth + 2] ? rowsJadwal[j][dayOfMonth + 2].toString().toUpperCase().trim() : "";
                    break;
                }
            }
        }
        
        var rowP = null;
        for(var i = rowsPresensi.length - 1; i > 0; i--) {
          var rTgl = rowsPresensi[i][0];
          var rTglStr = (rTgl instanceof Date) ? Utilities.formatDate(rTgl, "GMT+7", "dd-MM-yyyy") : rTgl.toString().trim();
          if (rTglStr === tglStr && rowsPresensi[i][1].toString() === uname) {
            rowP = rowsPresensi[i];
            break; 
          }
        }
        
        var isCuti = false;
        var jenisCuti = "";
        for(var c=0; c<cutiDisetujui.length; c++) {
          if(cutiDisetujui[c].tgl === tglStr) { isCuti = true; jenisCuti = cutiDisetujui[c].jenis; break; }
        }

        var resRow = { tgl: tglStr, msk: "", plg: "", tlt: 0, cpt: 0, pm: 0, pp: 0, pt: 0, sts: "" };
        
        if (isCuti) {
           resRow.sts = jenisCuti;
           resRow.msk = "-"; resRow.plg = "-";
           if (jenisCuti === "Izin") {
             resRow.pt = 1; sum.tPot += 1;
           } else if (jenisCuti === "Sakit (Surat Dokter)" || jenisCuti === "Cuti Tahunan") {
             resRow.pt = 0;
           } else {
             resRow.pt = 0;
           }
           sum.tHari++;
           if(resRow.pt > 0) sum.hSanksi++;
        }
        else if (rowP) {
           var mTxt = rowP[3] ? formatJam(rowP[3]) : "";
           var pTxt = rowP[4] ? formatJam(rowP[4]) : "";
           resRow.msk = mTxt; resRow.plg = pTxt;
           
           if(mTxt) {
              var jM = parseInt(mTxt.split(".")[0]); var mM = parseInt(mTxt.split(".")[1]);
              var act = jM * 60 + mM;
              
              if (peran === "Satpam") {
                 var isShiftP = false; var isShiftM = false;
                 if (adaJadwalBulanIni && shiftTerjadwal !== null && shiftTerjadwal !== "") {
                     if (shiftTerjadwal.startsWith("P")) isShiftP = true;
                     else if (shiftTerjadwal.startsWith("M")) isShiftM = true;
                     else isShiftP = true; 
                 } else {
                     isShiftP = (jM >= 4 && jM < 15);
                     isShiftM = (jM >= 15 || jM < 4);
                 }
                 
                 if(isShiftP) {
                    var bts = 7*60;
                    if(act > bts) resRow.tlt = act - bts;
                 } else if (isShiftM) {
                    var bts = 19*60; if(jM < 4) act += 24*60;
                    if(act > bts) resRow.tlt = act - bts;
                 }
              } else {
                 var bts = (peran === "Petugas Kebersihan" || peran === "Pengemudi") ? (6 * 60 + 30) : (7 * 60);
                 if (act > bts) resRow.tlt = act - bts;
              }
              
              if(resRow.tlt > 0 && resRow.tlt <= 30) resRow.pm = 0.25;
              else if(resRow.tlt > 30) resRow.pm = 0.5;
           }
           
           if(!pTxt) {
              resRow.plg = "Lupa Absen";
              resRow.pp = 3; 
              resRow.pm = 0; 
              resRow.cpt = 0;
           } else {
              var jK = parseInt(pTxt.split(".")[0]); var mK = parseInt(pTxt.split(".")[1]);
              var actK = jK * 60 + mK;
              
              if (peran === "Satpam") {
                  var msTxt = rowP[3] ? formatJam(rowP[3]) : "07.00"; 
                  var jsM = parseInt(msTxt.split(".")[0]);
                  var isSP = false;
                  if (adaJadwalBulanIni && shiftTerjadwal !== null && shiftTerjadwal !== "") {
                      if (shiftTerjadwal.startsWith("P")) isSP = true;
                  } else {
                      isSP = (jsM >= 4 && jsM < 15);
                  }
                  
                  if(isSP) {
                     var btsK = 15*60; 
                     if(actK < btsK && actK > 7*60) resRow.cpt = btsK - actK; 
                  } else {
                     var btsK = 7*60 + 24*60; if(jK >= 15) actK += 24*60; 
                     if(actK < btsK && actK > 19*60) resRow.cpt = btsK - actK;
                  }
              } else {
                  var btsK = 17*60;
                  if(actK < btsK) resRow.cpt = btsK - actK; 
              }
              
              if(resRow.cpt > 0 && resRow.cpt <= 30) resRow.pp = 0.25;
              else if(resRow.cpt > 30) resRow.pp = 0.5;
           }
           
           resRow.pt = resRow.pm + resRow.pp;
           resRow.sts = "Hadir";
           
           sum.tHari++;
           sum.hHadir++;
           if(resRow.pt > 0) sum.hSanksi++;
           sum.tPot += resRow.pt;
        }
        else {
           if (peran === "Satpam") {
               if (adaJadwalBulanIni) {
                   if (shiftTerjadwal && shiftTerjadwal !== "") {
                       resRow.sts = "Alpha"; resRow.pt = 3; sum.tHari++; sum.tPot += 3;
                   } else {
                       resRow.sts = "Libur"; resRow.msk = "-"; resRow.plg = "-";
                   }
               } else {
                   resRow.sts = "Alpha"; resRow.pt = 3; sum.tHari++; sum.tPot += 3;
               }
           } else {
               if (dayOfWeek === 0 || dayOfWeek === 6) {
                   resRow.sts = "Libur"; resRow.msk = "-"; resRow.plg = "-";
               } else {
                   resRow.sts = "Alpha"; resRow.pt = 3; sum.tHari++; sum.tPot += 3;
               }
           }
        }
        hasil.push(resRow);
      }
      
      return ContentService.createTextOutput(JSON.stringify({status: "success", data: hasil, sum: sum})).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
    }
  }
}
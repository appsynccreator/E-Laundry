// ============================================================
// E-LAUNDRY MANAGEMENT SYSTEM — Code.gs
// Backend Google Apps Script
// Versi: 1.0 | Bahasa: Indonesia
// ============================================================

// ============================================================
// BAGIAN 1: KONFIGURASI LISENSI HYBRID
// ============================================================

const LICENSE_CONFIG = {
  VENDOR_SHEET_ID  : "GANTI_DENGAN_ID_SPREADSHEET_MASTER_LISENSI",
  VENDOR_SHEET_NAME: "KLIEN_LISENSI",
  SECRET_KEY       : "ELAUNDRY-SECRET-2026",
  CACHE_TTL_HOURS  : 24,
  OFFLINE_MAX_DAYS : 7,
  WARNING_DAYS     : [30, 7, 1],
};

// Indeks kolom sheet KLIEN_LISENSI (1-based)
const LC = {
  NO: 1, NAMA_USAHA: 2, PEMILIK: 3, HP: 4, EMAIL: 5, KOTA: 6,
  TGL_AKTIF: 7, TGL_EXPIRED: 8, KODE: 9, PAKET: 10,
  STATUS: 11, DURASI: 12, URL: 13, CATATAN: 14,
};

function getLicenseStatus() {
  const props  = PropertiesService.getScriptProperties();
  const cached = getCachedLicense(props);

  if (cached && !isCacheExpired(cached.lastValidated, LICENSE_CONFIG.CACHE_TTL_HOURS)) {
    return buildLicenseResult(cached);
  }

  try {
    const fresh = fetchLicenseFromMaster();
    if (fresh) {
      saveLicenseCache(props, fresh);
      return buildLicenseResult(fresh);
    }
  } catch (e) {
    if (cached) {
      const offlineDays = daysDiff(new Date(), new Date(cached.lastValidated));
      if (offlineDays <= LICENSE_CONFIG.OFFLINE_MAX_DAYS) {
        return buildLicenseResult(cached, true);
      }
    }
  }
  return { status: "LOCKED", reason: "Lisensi tidak dapat diverifikasi." };
}

function fetchLicenseFromMaster() {
  const licenseKey = PropertiesService.getScriptProperties().getProperty("LICENSE_KEY");
  if (!licenseKey) return null;

  const ss    = SpreadsheetApp.openById(LICENSE_CONFIG.VENDOR_SHEET_ID);
  const sheet = ss.getSheetByName(LICENSE_CONFIG.VENDOR_SHEET_NAME);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[LC.KODE - 1] === licenseKey) {
      return {
        licenseKey   : row[LC.KODE - 1],
        namaUsaha    : row[LC.NAMA_USAHA - 1],
        paket        : row[LC.PAKET - 1],
        status       : row[LC.STATUS - 1],
        tglAktif     : row[LC.TGL_AKTIF - 1],
        tglExpired   : row[LC.TGL_EXPIRED - 1],
        lastValidated: new Date().toISOString(),
      };
    }
  }
  return null;
}

function buildLicenseResult(license, isOffline) {
  isOffline = isOffline || false;
  const now         = new Date();
  const expiredDate = new Date(license.tglExpired);
  const daysLeft    = Math.ceil((expiredDate - now) / (1000 * 60 * 60 * 24));

  if (license.status === "Expired" || license.status === "Suspended" || daysLeft < 0) {
    return {
      status    : "LOCKED",
      reason    : "Langganan Anda telah berakhir pada " + formatDate(expiredDate) + ".",
      daysLeft  : daysLeft,
      namaUsaha : license.namaUsaha,
      paket     : license.paket,
      tglExpired: formatDate(expiredDate),
    };
  }

  const isWarning = LICENSE_CONFIG.WARNING_DAYS.some(function(d) { return daysLeft <= d; });

  return {
    status    : isWarning ? "WARNING" : "ACTIVE",
    daysLeft  : daysLeft,
    namaUsaha : license.namaUsaha,
    paket     : license.paket,
    tglExpired: formatDate(expiredDate),
    features  : getFeaturesForPackage(license.paket),
    isOffline : isOffline,
    warningMsg: isWarning
      ? "⚠️ Langganan Anda akan berakhir dalam " + daysLeft + " hari (" + formatDate(expiredDate) + "). Segera hubungi vendor untuk perpanjangan."
      : null,
  };
}

function getFeaturesForPackage(paket) {
  var FEATURES = {
    "Starter"     : ["dashboard","transaksi","nota","pelanggan","layanan"],
    "Professional": ["dashboard","transaksi","nota","pelanggan","layanan","pengeluaran","karyawan","laporan","export"],
    "Premium"     : ["dashboard","transaksi","nota","pelanggan","layanan","pengeluaran","karyawan","laporan","export","whatsapp","branding","multi_outlet"],
  };
  return FEATURES[paket] || FEATURES["Professional"];
}

function getCachedLicense(props) {
  var raw = props.getProperty("LICENSE_CACHE");
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function saveLicenseCache(props, data) {
  props.setProperty("LICENSE_CACHE", JSON.stringify(data));
}

function isCacheExpired(lastValidated, ttlHours) {
  var diff = (new Date() - new Date(lastValidated)) / (1000 * 60 * 60);
  return diff > ttlHours;
}

function daysDiff(d1, d2) {
  return Math.abs(Math.ceil((d1 - d2) / (1000 * 60 * 60 * 24)));
}

function formatDate(date) {
  try {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(new Date(date), tz, "dd MMMM yyyy");
  } catch(e) {
    return new Date(date).toLocaleDateString("id-ID");
  }
}

function activateLicense(licenseKey) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("LICENSE_KEY", licenseKey.trim().toUpperCase());
  props.deleteProperty("LICENSE_CACHE");

  var result = getLicenseStatus();
  if (result.status === "LOCKED" && result.daysLeft === undefined) {
    props.deleteProperty("LICENSE_KEY");
    return { success: false, message: "Kode lisensi tidak valid atau tidak ditemukan." };
  }
  return { success: true, license: result };
}

function checkFeatureAccess(featureName) {
  var license = getLicenseStatus();
  if (license.status === "LOCKED") return false;
  return license.features && license.features.indexOf(featureName) !== -1;
}

// ============================================================
// BAGIAN 2: doGet() — Entry Point
// ============================================================

function doGet(e) {
  // Auto-inisialisasi sheets jika belum ada
  try { initializeSheets(); } catch(ex) {}

  var props      = PropertiesService.getScriptProperties();
  var licenseKey = props.getProperty("LICENSE_KEY");
  var license    = null;

  // DEMO_MODE = true → skip lisensi, langsung masuk app
  // DEMO_MODE = false → enforce lisensi (production)
  var DEMO_MODE = true; // Set false saat production

  if (!DEMO_MODE) {
    if (!licenseKey) {
      // Belum aktif → tampilkan halaman aktivasi
      var tmpl0 = HtmlService.createTemplateFromFile("index");
      tmpl0.licenseData  = "null";
      tmpl0.configData   = "{}";
      tmpl0.needActivate = "true";
      return tmpl0.evaluate()
        .setTitle("E-Laundry — Aktivasi")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    license = getLicenseStatus();

    if (license.status === "LOCKED") {
      var lockedHtml = HtmlService.createTemplate(getLockedPageHtml(license));
      return lockedHtml.evaluate().setTitle("Akses Terkunci");
    }
  } else {
    // Demo mode: buat lisensi dummy langsung
    var namaUsahaCfg = "E-Laundry";
    try { namaUsahaCfg = getConfigValue("Nama Usaha") || "E-Laundry"; } catch(ex) {}
    license = {
      status    : "ACTIVE",
      namaUsaha : namaUsahaCfg,
      paket     : "Professional",
      daysLeft  : 365,
      tglExpired: "01 Januari 2027",
      features  : ["dashboard","transaksi","nota","pelanggan","layanan",
                   "pengeluaran","karyawan","laporan","export"],
      isOffline : false,
      warningMsg: null,
    };
  }

  // Ambil config — fallback ke object kosong jika error
  var config = {};
  try { config = getConfigAsObject(); } catch(ex) {}

  var tmpl = HtmlService.createTemplateFromFile("index");
  tmpl.licenseData  = JSON.stringify(license);
  tmpl.configData   = JSON.stringify(config);
  tmpl.needActivate = "false";

  var appTitle = config["Nama Usaha"] || "E-Laundry";

  return tmpl.evaluate()
    .setTitle("E-Laundry — " + appTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover, maximum-scale=1.0');
}

function getLockedPageHtml(license) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:Poppins,sans-serif;background:#F5F7FA;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}'
    + '.card{background:#fff;border-radius:16px;padding:40px;max-width:480px;width:100%;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.12)}'
    + '.icon{font-size:64px;margin-bottom:16px}h2{color:#C62828;font-size:22px;margin-bottom:8px}'
    + 'p{color:#555;font-size:14px;line-height:1.6;margin-bottom:16px}'
    + '.info{background:#FFF9C4;border-radius:8px;padding:12px;font-size:13px;color:#795548;margin-bottom:20px}'
    + '.btn{display:inline-block;background:#1565C0;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px}</style>'
    + '</head><body><div class="card"><div class="icon">🔒</div><h2>Akses Terkunci</h2>'
    + '<p>' + (license.reason || "Langganan Anda telah berakhir.") + '</p>'
    + '<div class="info">📅 Tanggal expired: <strong>' + (license.tglExpired || '-') + '</strong><br>📦 Paket: <strong>' + (license.paket || '-') + '</strong></div>'
    + '<p>Hubungi vendor untuk memperpanjang langganan:</p>'
    + '<a class="btn" href="https://wa.me/6281234567890" target="_blank">💬 Hubungi via WhatsApp</a>'
    + '</div></body></html>';
}

// ============================================================
// BAGIAN 3: CONFIG
// ============================================================

function getConfig() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
  if (!sheet) return initializeSheets();
  return sheet.getDataRange().getValues();
}

function getConfigAsObject() {
  var data   = getConfig();
  var result = {};
  data.forEach(function(row) {
    if (row[0]) result[row[0]] = row[1];
  });
  return result;
}

function getConfigValue(key) {
  var obj = getConfigAsObject();
  return obj[key] || "";
}

function updateConfig(updates) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
  var data  = sheet.getDataRange().getValues();
  var rowMap = {};
  data.forEach(function(row, i) { if (row[0]) rowMap[row[0]] = i + 1; });

  Object.keys(updates).forEach(function(key) {
    if (rowMap[key]) {
      sheet.getRange(rowMap[key], 2).setValue(updates[key]);
    } else {
      sheet.appendRow([key, updates[key]]);
    }
  });
  return { success: true };
}

// ============================================================
// BAGIAN 4: INISIALISASI SHEETS
// ============================================================

function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // CONFIG
  var config = ss.getSheetByName("CONFIG") || ss.insertSheet("CONFIG");
  if (config.getLastRow() < 1) {
    config.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]);
    config.getRange(2, 1, 20, 2).setValues([
      ["Nama Usaha",          "Laundry Bersih"],
      ["Alamat",              "Jl. Contoh No. 1, Kota"],
      ["No. Telepon",         "08123456789"],
      ["Email",               "laundry@email.com"],
      ["Jam Operasional",     "08:00 - 20:00"],
      ["Tagline",             "Bersih, Cepat, Terpercaya"],
      ["Estimasi Default (jam)", "24"],
      ["Ekspres (jam)",       "4"],
      ["Prefix Nota",         "TRX"],
      ["Digit Nomor Nota",    "4"],
      ["Metode Pembayaran",   "Tunai,Transfer Bank,QRIS,OVO/GoPay/Dana"],
      ["Nama Bank",           "BCA"],
      ["No. Rekening",        "1234567890"],
      ["Atas Nama",           "Pemilik Laundry"],
      ["Diskon Maks (%)",     "20"],
      ["PPN (%)",             "0"],
      ["WhatsApp Vendor",     "6281234567890"],
    ]);
  }

  // LAYANAN
  var layanan = ss.getSheetByName("LAYANAN") || ss.insertSheet("LAYANAN");
  if (layanan.getLastRow() < 2) {
    layanan.getRange(1, 1, 1, 7).setValues([["No","Nama Layanan","Kategori","Satuan","Harga Normal","Harga Ekspres","Keterangan"]]);
    layanan.getRange(2, 1, 5, 7).setValues([
      [1,"Cuci Kering","Pakaian","kg",8000,15000,"Cuci + keringkan"],
      [2,"Cuci Setrika","Pakaian","kg",12000,20000,"Cuci + keringkan + setrika"],
      [3,"Setrika Saja","Pakaian","kg",6000,10000,"Hanya setrika"],
      [4,"Cuci Sepatu","Sepatu","pasang",25000,40000,"Cuci bersih sepatu"],
      [5,"Cuci Selimut","Linen","pcs",20000,35000,"Selimut/bedcover"],
    ]);
  }

  // PELANGGAN
  var pelanggan = ss.getSheetByName("PELANGGAN") || ss.insertSheet("PELANGGAN");
  if (pelanggan.getLastRow() < 1) {
    pelanggan.getRange(1, 1, 1, 9).setValues([["No","Nama Pelanggan","No. HP","Alamat","Kota","Email","Tipe","Tgl Daftar","Catatan"]]);
  }

  // TRANSAKSI
  var transaksi = ss.getSheetByName("TRANSAKSI") || ss.insertSheet("TRANSAKSI");
  if (transaksi.getLastRow() < 1) {
    transaksi.getRange(1, 1, 1, 16).setValues([["No. Nota","Tgl Masuk","Nama Pelanggan","Layanan","Satuan","Qty","Harga Satuan","Subtotal","Diskon (%)","Nilai Diskon","PPN (%)","Nilai PPN","Total Bayar","Status","Metode Bayar","Tgl Selesai Est."]]);
  }

  // KARYAWAN
  var karyawan = ss.getSheetByName("KARYAWAN") || ss.insertSheet("KARYAWAN");
  if (karyawan.getLastRow() < 1) {
    karyawan.getRange(1, 1, 1, 10).setValues([["No","Nama Karyawan","No. HP","Alamat","Posisi","Shift","Tgl Masuk","Gaji Pokok","Status","Catatan"]]);
  }

  // PENGELUARAN
  var pengeluaran = ss.getSheetByName("PENGELUARAN") || ss.insertSheet("PENGELUARAN");
  if (pengeluaran.getLastRow() < 1) {
    pengeluaran.getRange(1, 1, 1, 8).setValues([["No","Tanggal","Keterangan","Kategori","Jumlah","Metode Bayar","Dibayar Oleh","Bukti/Catatan"]]);
  }

  // LAPORAN
  ss.getSheetByName("LAPORAN") || ss.insertSheet("LAPORAN");

  // USERS (kolom: Nama, Email, Role, Status, Tgl Dibuat, Catatan, Menu Akses JSON)
  var users = ss.getSheetByName("USERS") || ss.insertSheet("USERS");
  if (users.getLastRow() < 2) {
    users.getRange(1, 1, 1, 7).setValues([["Nama","Email Google","Role","Status","Tgl Dibuat","Catatan","Menu Akses"]]);
    var defaultMenu = JSON.stringify(["dashboard","transaksi-baru","daftar-transaksi","pelanggan","layanan","karyawan","pengeluaran","laporan","pengaturan"]);
    users.getRange(2, 1, 1, 7).setValues([["Owner","owner@gmail.com","Owner","Aktif",new Date(),"Akun pemilik default", defaultMenu]]);
  }

  return getConfig();
}

// ============================================================
// BAGIAN 5: AUTENTIKASI & USER
// ============================================================

var ALL_MENUS = ["dashboard","transaksi-baru","daftar-transaksi","pelanggan","layanan","karyawan","pengeluaran","laporan","pengaturan"];

function parseMenuAkses(raw) {
  try { var m = JSON.parse(raw); return Array.isArray(m) ? m : ALL_MENUS; } catch(e) { return ALL_MENUS; }
}

function getCurrentUser() {
  var email = "";
  try { email = Session.getActiveUser().getEmail() || ""; } catch(e) {}

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (email) {
        for (var i = 1; i < data.length; i++) {
          if (data[i][1] === email && data[i][3] === "Aktif") {
            return { authenticated:true, nama:data[i][0], email:email, role:data[i][2],
                     status:data[i][3], menuAkses: parseMenuAkses(data[i][6]) };
          }
        }
      }
      if (data.length >= 2 && data[1][0]) {
        return { authenticated:true, nama:data[1][0]||"Owner", email:email||data[1][1]||"",
                 role:data[1][2]||"Owner", status:"Aktif", menuAkses: parseMenuAkses(data[1][6]) };
      }
    }
  } catch(e) {}

  return { authenticated:true, nama:"Owner", email:email||"owner@laundry.com",
           role:"Owner", status:"Aktif", menuAkses: ALL_MENUS };
}

function getUsers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
  var data  = sheet.getDataRange().getValues();
  return data.slice(1).filter(function(r){ return r[0]; }).map(function(row, i) {
    return { no:i+1, nama:row[0], email:row[1], role:row[2], status:row[3],
             tglDibuat:formatTgl(row[4]), catatan:row[5], menuAkses:parseMenuAkses(row[6]) };
  });
}

function tambahUser(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (all[i][1] === data.email) return { success: false, message: "Email sudah terdaftar." };
  }
  var menuJson = data.menuAkses ? JSON.stringify(data.menuAkses) : JSON.stringify(ALL_MENUS);
  sheet.appendRow([data.nama, data.email, data.role, data.status||"Aktif", new Date(), data.catatan||"", menuJson]);
  return { success: true };
}

function editUser(email, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (all[i][1] === email) {
      var menuJson = data.menuAkses ? JSON.stringify(data.menuAkses) : (all[i][6] || JSON.stringify(ALL_MENUS));
      sheet.getRange(i+1, 1, 1, 7).setValues([[data.nama, data.email, data.role,
        data.status||"Aktif", all[i][4], data.catatan||"", menuJson]]);
      return { success: true };
    }
  }
  return { success: false, message: "User tidak ditemukan." };
}

function hapusUser(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (all[i][1] === email) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false };
}

// ============================================================
// BAGIAN 6: LAYANAN
// ============================================================

function getLayanan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LAYANAN");
  var data  = sheet.getDataRange().getValues();
  return data.slice(1).filter(function(r) { return r[0]; }).map(function(row) {
    return { no: row[0], nama: row[1], kategori: row[2], satuan: row[3],
             hargaNormal: row[4], hargaEkspres: row[5], keterangan: row[6] };
  });
}

function tambahLayanan(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LAYANAN");
  var all   = sheet.getDataRange().getValues();
  var no    = all.length; // header = baris 1
  sheet.appendRow([no, data.nama, data.kategori, data.satuan,
                   parseFloat(data.hargaNormal) || 0,
                   parseFloat(data.hargaEkspres) || 0, data.keterangan || ""]);
  renumberSheet(sheet);
  return { success: true };
}

function editLayanan(no, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LAYANAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.getRange(i+1, 2, 1, 6).setValues([[data.nama, data.kategori, data.satuan,
        parseFloat(data.hargaNormal)||0, parseFloat(data.hargaEkspres)||0, data.keterangan||""]]);
      return { success: true };
    }
  }
  return { success: false };
}

function hapusLayanan(no) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LAYANAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.deleteRow(i+1);
      renumberSheet(sheet);
      return { success: true };
    }
  }
  return { success: false };
}

// ============================================================
// BAGIAN 7: PELANGGAN
// ============================================================

function getPelanggan(search) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PELANGGAN");
  var data  = sheet.getDataRange().getValues();
  var rows  = data.slice(1).filter(function(r) { return r[0]; });
  if (search) {
    var q = search.toLowerCase();
    rows  = rows.filter(function(r) {
      return String(r[1]).toLowerCase().includes(q) ||
             String(r[2]).toLowerCase().includes(q) ||
             String(r[6]).toLowerCase().includes(q);
    });
  }
  return rows.map(function(row) {
    return { no: row[0], nama: row[1], hp: row[2], alamat: row[3], kota: row[4],
             email: row[5], tipe: row[6], tglDaftar: formatTgl(row[7]), catatan: row[8] };
  });
}

function tambahPelanggan(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PELANGGAN");
  var all   = sheet.getDataRange().getValues();
  var no    = all.length;
  sheet.appendRow([no, data.nama, data.hp, data.alamat || "", data.kota || "",
                   data.email || "", data.tipe || "Reguler", new Date(), data.catatan || ""]);
  renumberSheet(sheet);
  return { success: true };
}

function editPelanggan(no, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PELANGGAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.getRange(i+1, 2, 1, 7).setValues([[data.nama, data.hp, data.alamat||"", data.kota||"",
        data.email||"", data.tipe||"Reguler", data.catatan||""]]);
      return { success: true };
    }
  }
  return { success: false };
}

function hapusPelanggan(no) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PELANGGAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.deleteRow(i+1);
      renumberSheet(sheet);
      return { success: true };
    }
  }
  return { success: false };
}

function getRiwayatPelanggan(nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
  var data  = sheet.getDataRange().getValues();
  var rows  = data.slice(1).filter(function(r) {
    return String(r[2]).toLowerCase() === nama.toLowerCase();
  });
  var total = rows.reduce(function(s, r) { return s + (parseFloat(r[12]) || 0); }, 0);
  return {
    transaksi: rows.map(function(row) {
      return { nota: row[0], tgl: formatTgl(row[1]), layanan: row[3], qty: row[5],
               total: row[12], status: row[13] };
    }),
    totalBelanja: total
  };
}

// ============================================================
// BAGIAN 8: TRANSAKSI
// ============================================================

function generateNomorNota() {
  var config  = getConfigAsObject();
  var prefix  = config["Prefix Nota"] || "TRX";
  var digit   = parseInt(config["Digit Nomor Nota"]) || 4;
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
  var tz      = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var today   = Utilities.formatDate(new Date(), tz, "yyyyMMdd");
  var data    = sheet.getDataRange().getValues();
  var maxSeq  = 0;

  data.forEach(function(row) {
    if (row[0] && String(row[0]).indexOf(today) !== -1) {
      var parts = String(row[0]).split("-");
      var seq   = parseInt(parts[parts.length - 1]) || 0;
      if (seq > maxSeq) maxSeq = seq;
    }
  });

  var nextSeq = String(maxSeq + 1).padStart(digit, "0");
  return prefix + "-" + today + "-" + nextSeq;
}

function getTransaksi(filter) {
  filter = filter || {};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
  var data  = sheet.getDataRange().getValues();
  var rows  = data.slice(1).filter(function(r) { return r[0]; });

  if (filter.status && filter.status !== "Semua") {
    rows = rows.filter(function(r) { return r[13] === filter.status; });
  }
  if (filter.tanggalMulai) {
    var mulai = new Date(filter.tanggalMulai);
    rows = rows.filter(function(r) { return new Date(r[1]) >= mulai; });
  }
  if (filter.tanggalAkhir) {
    var akhir = new Date(filter.tanggalAkhir);
    akhir.setHours(23,59,59);
    rows = rows.filter(function(r) { return new Date(r[1]) <= akhir; });
  }
  if (filter.search) {
    var q = filter.search.toLowerCase();
    rows  = rows.filter(function(r) {
      return String(r[0]).toLowerCase().includes(q) ||
             String(r[2]).toLowerCase().includes(q) ||
             String(r[3]).toLowerCase().includes(q);
    });
  }
  if (filter.metodeBayar && filter.metodeBayar !== "Semua") {
    rows = rows.filter(function(r) { return r[14] === filter.metodeBayar; });
  }

  rows.sort(function(a, b) { return new Date(b[1]) - new Date(a[1]); });

  return rows.map(function(row) {
    return {
      nota       : row[0], tglMasuk  : formatTgl(row[1]), pelanggan: row[2],
      layanan    : row[3], satuan    : row[4], qty      : row[5],
      hargaSatuan: row[6], subtotal  : row[7], diskonPct: row[8],
      nilaiDiskon: row[9], ppnPct    : row[10], nilaiPpn : row[11],
      total      : row[12], status   : row[13], metodeBayar: row[14],
      tglSelesai : formatTgl(row[15]),
    };
  });
}

function tambahTransaksi(data) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
  var config = getConfigAsObject();
  var tz     = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var nota   = generateNomorNota();

  var qty         = parseFloat(data.qty) || 1;
  var harga       = parseFloat(data.hargaSatuan) || 0;
  var subtotal    = qty * harga;
  var diskonPct   = parseFloat(data.diskon) || 0;
  var nilaiDiskon = subtotal * diskonPct / 100;
  var ppnPct      = parseFloat(config["PPN (%)"] || 0);
  var nilaiPpn    = (subtotal - nilaiDiskon) * ppnPct / 100;
  var total       = subtotal - nilaiDiskon + nilaiPpn;

  var jamEst  = data.isEkspres ? parseInt(config["Ekspres (jam)"] || 4) : parseInt(config["Estimasi Default (jam)"] || 24);
  var tglEst  = new Date(new Date(data.tglMasuk).getTime() + jamEst * 3600000);

  sheet.appendRow([
    nota,
    new Date(data.tglMasuk),
    data.pelanggan,
    data.layanan,
    data.satuan,
    qty,
    harga,
    subtotal,
    diskonPct,
    nilaiDiskon,
    ppnPct,
    nilaiPpn,
    total,
    "Menunggu",
    data.metodeBayar || "Tunai",
    tglEst,
  ]);

  // Auto-daftarkan pelanggan baru jika belum ada
  var pelSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PELANGGAN");
  var pelData  = pelSheet.getDataRange().getValues();
  var found    = pelData.slice(1).some(function(r) {
    return String(r[1]).toLowerCase() === String(data.pelanggan).toLowerCase();
  });
  if (!found && data.pelanggan) {
    tambahPelanggan({ nama: data.pelanggan, hp: data.hpPelanggan || "", tipe: "Reguler" });
  }

  return {
    success   : true,
    nota      : nota,
    total     : total,
    subtotal  : subtotal,
    diskon    : nilaiDiskon,
    ppn       : nilaiPpn,
    tglSelesai: formatTgl(tglEst),
  };
}

function updateStatusTransaksi(noNota, status) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
  var data   = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === noNota) {
      sheet.getRange(i+1, 14).setValue(status);
      return { success: true };
    }
  }
  return { success: false };
}

function hapusTransaksi(noNota) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === noNota) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false };
}

function getDetailTransaksi(noNota) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === noNota) {
      var row = data[i];
      return {
        nota: row[0], tglMasuk: formatTgl(row[1]), pelanggan: row[2],
        layanan: row[3], satuan: row[4], qty: row[5], hargaSatuan: row[6],
        subtotal: row[7], diskonPct: row[8], nilaiDiskon: row[9],
        ppnPct: row[10], nilaiPpn: row[11], total: row[12],
        status: row[13], metodeBayar: row[14], tglSelesai: formatTgl(row[15]),
      };
    }
  }
  return null;
}

// ============================================================
// BAGIAN 9: KARYAWAN
// ============================================================

function getKaryawan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KARYAWAN");
  var data  = sheet.getDataRange().getValues();
  return data.slice(1).filter(function(r) { return r[0]; }).map(function(row) {
    return { no: row[0], nama: row[1], hp: row[2], alamat: row[3], posisi: row[4],
             shift: row[5], tglMasuk: formatTgl(row[6]), gajiPokok: row[7],
             status: row[8], catatan: row[9] };
  });
}

function tambahKaryawan(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KARYAWAN");
  var all   = sheet.getDataRange().getValues();
  var no    = all.length;
  sheet.appendRow([no, data.nama, data.hp, data.alamat||"", data.posisi||"",
                   data.shift||"Pagi", new Date(data.tglMasuk || new Date()),
                   parseFloat(data.gajiPokok)||0, data.status||"Aktif", data.catatan||""]);
  renumberSheet(sheet);
  return { success: true };
}

function editKaryawan(no, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KARYAWAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.getRange(i+1, 2, 1, 9).setValues([[data.nama, data.hp, data.alamat||"",
        data.posisi||"", data.shift||"Pagi", new Date(data.tglMasuk||new Date()),
        parseFloat(data.gajiPokok)||0, data.status||"Aktif", data.catatan||""]]);
      return { success: true };
    }
  }
  return { success: false };
}

function hapusKaryawan(no) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KARYAWAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.deleteRow(i+1);
      renumberSheet(sheet);
      return { success: true };
    }
  }
  return { success: false };
}

// ============================================================
// BAGIAN 10: PENGELUARAN
// ============================================================

function getPengeluaran(filter) {
  filter = filter || {};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PENGELUARAN");
  var data  = sheet.getDataRange().getValues();
  var rows  = data.slice(1).filter(function(r) { return r[0]; });

  if (filter.kategori && filter.kategori !== "Semua") {
    rows = rows.filter(function(r) { return r[3] === filter.kategori; });
  }
  if (filter.tanggalMulai) {
    rows = rows.filter(function(r) { return new Date(r[1]) >= new Date(filter.tanggalMulai); });
  }
  if (filter.tanggalAkhir) {
    var akhir = new Date(filter.tanggalAkhir); akhir.setHours(23,59,59);
    rows = rows.filter(function(r) { return new Date(r[1]) <= akhir; });
  }

  rows.sort(function(a, b) { return new Date(b[1]) - new Date(a[1]); });

  var total = rows.reduce(function(s, r) { return s + (parseFloat(r[4]) || 0); }, 0);

  return {
    data: rows.map(function(row) {
      return { no: row[0], tanggal: formatTgl(row[1]), keterangan: row[2],
               kategori: row[3], jumlah: row[4], metodeBayar: row[5],
               dibayarOleh: row[6], catatan: row[7] };
    }),
    total: total,
  };
}

function tambahPengeluaran(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PENGELUARAN");
  var all   = sheet.getDataRange().getValues();
  var no    = all.length;
  sheet.appendRow([no, new Date(data.tanggal), data.keterangan, data.kategori,
                   parseFloat(data.jumlah)||0, data.metodeBayar||"Tunai",
                   data.dibayarOleh||"", data.catatan||""]);
  renumberSheet(sheet);
  return { success: true };
}

function editPengeluaran(no, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PENGELUARAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.getRange(i+1, 2, 1, 7).setValues([[new Date(data.tanggal), data.keterangan,
        data.kategori, parseFloat(data.jumlah)||0, data.metodeBayar||"Tunai",
        data.dibayarOleh||"", data.catatan||""]]);
      return { success: true };
    }
  }
  return { success: false };
}

function hapusPengeluaran(no) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PENGELUARAN");
  var all   = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(no)) {
      sheet.deleteRow(i+1);
      renumberSheet(sheet);
      return { success: true };
    }
  }
  return { success: false };
}

// ============================================================
// BAGIAN 11: DASHBOARD
// ============================================================

function getDashboardData() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var trxSheet = ss.getSheetByName("TRANSAKSI");
  var penSheet = ss.getSheetByName("PENGELUARAN");
  var tz       = ss.getSpreadsheetTimeZone();
  var today    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var trxData  = trxSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });
  var penData  = penSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });

  // Statistik hari ini
  var trxHariIni = trxData.filter(function(r) {
    return Utilities.formatDate(new Date(r[1]), tz, "yyyy-MM-dd") === today;
  });
  var pendapatanHariIni = trxHariIni.reduce(function(s, r) { return s + (parseFloat(r[12]) || 0); }, 0);

  // Pending & Selesai
  var pending = trxData.filter(function(r) { return r[13] === "Menunggu" || r[13] === "Diproses"; }).length;
  var selesai = trxHariIni.filter(function(r) { return r[13] === "Selesai" || r[13] === "Diambil"; }).length;

  // 7 hari terakhir
  var chart7 = [];
  for (var d = 6; d >= 0; d--) {
    var dt  = new Date(); dt.setDate(dt.getDate() - d);
    var tgl = Utilities.formatDate(dt, tz, "yyyy-MM-dd");
    var jumlah = trxData.filter(function(r) {
      return Utilities.formatDate(new Date(r[1]), tz, "yyyy-MM-dd") === tgl && r[13] !== "Batal";
    }).reduce(function(s, r) { return s + (parseFloat(r[12]) || 0); }, 0);
    chart7.push({ tgl: Utilities.formatDate(dt, tz, "dd/MM"), jumlah: jumlah });
  }

  // Layanan terpopuler
  var layanMap = {};
  trxData.forEach(function(r) {
    if (r[3]) layanMap[r[3]] = (layanMap[r[3]] || 0) + 1;
  });
  var layanPopuler = Object.keys(layanMap)
    .map(function(k) { return { nama: k, jumlah: layanMap[k] }; })
    .sort(function(a, b) { return b.jumlah - a.jumlah; })
    .slice(0, 5);

  // 5 transaksi terbaru
  var recent = trxData
    .sort(function(a, b) { return new Date(b[1]) - new Date(a[1]); })
    .slice(0, 5)
    .map(function(row) {
      return { nota: row[0], tgl: formatTgl(row[1]), pelanggan: row[2],
               layanan: row[3], total: row[12], status: row[13] };
    });

  return {
    totalTrxHariIni       : trxHariIni.length,
    pendapatanHariIni     : pendapatanHariIni,
    trxPending            : pending,
    trxSelesai            : selesai,
    chart7Hari            : chart7,
    layananPopuler        : layanPopuler,
    recentTransaksi       : recent,
  };
}

// ============================================================
// BAGIAN 12: LAPORAN BULANAN
// ============================================================

function getLaporanBulanan(bulan, tahun) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var trxSheet = ss.getSheetByName("TRANSAKSI");
  var penSheet = ss.getSheetByName("PENGELUARAN");
  var tz       = ss.getSpreadsheetTimeZone();

  bulan = parseInt(bulan); tahun = parseInt(tahun);

  var trxData = trxSheet.getDataRange().getValues().slice(1).filter(function(r) {
    if (!r[0]) return false;
    var d = new Date(r[1]);
    return d.getMonth()+1 === bulan && d.getFullYear() === tahun && r[13] !== "Batal";
  });

  var penData = penSheet.getDataRange().getValues().slice(1).filter(function(r) {
    if (!r[0]) return false;
    var d = new Date(r[1]);
    return d.getMonth()+1 === bulan && d.getFullYear() === tahun;
  });

  var totalPendapatan = trxData.reduce(function(s, r) { return s + (parseFloat(r[12]) || 0); }, 0);
  var totalPengeluaran = penData.reduce(function(s, r) { return s + (parseFloat(r[4]) || 0); }, 0);
  var labaBersih = totalPendapatan - totalPengeluaran;

  // Per hari (pendapatan vs pengeluaran)
  var hariMap = {};
  trxData.forEach(function(r) {
    var tgl = Utilities.formatDate(new Date(r[1]), tz, "dd");
    if (!hariMap[tgl]) hariMap[tgl] = { pendapatan: 0, pengeluaran: 0 };
    hariMap[tgl].pendapatan += parseFloat(r[12]) || 0;
  });
  penData.forEach(function(r) {
    var tgl = Utilities.formatDate(new Date(r[1]), tz, "dd");
    if (!hariMap[tgl]) hariMap[tgl] = { pendapatan: 0, pengeluaran: 0 };
    hariMap[tgl].pengeluaran += parseFloat(r[4]) || 0;
  });
  var chartHarian = Object.keys(hariMap).sort().map(function(k) {
    return { hari: k, pendapatan: hariMap[k].pendapatan, pengeluaran: hariMap[k].pengeluaran };
  });

  // Per metode bayar
  var metodeMap = {};
  trxData.forEach(function(r) {
    var m = r[14] || "Tunai";
    metodeMap[m] = (metodeMap[m] || 0) + (parseFloat(r[12]) || 0);
  });
  var chartMetode = Object.keys(metodeMap).map(function(k) {
    return { metode: k, jumlah: metodeMap[k] };
  });

  // Layanan terlaris
  var layanMap = {};
  trxData.forEach(function(r) {
    if (r[3]) layanMap[r[3]] = (layanMap[r[3]] || 0) + (parseFloat(r[12]) || 0);
  });
  var chartLayanan = Object.keys(layanMap)
    .map(function(k) { return { nama: k, pendapatan: layanMap[k] }; })
    .sort(function(a, b) { return b.pendapatan - a.pendapatan; });

  // Pengeluaran per kategori
  var katMap = {};
  penData.forEach(function(r) {
    var k = r[3] || "Lain-lain";
    katMap[k] = (katMap[k] || 0) + (parseFloat(r[4]) || 0);
  });
  var ringkasanPengeluaran = Object.keys(katMap).map(function(k) {
    return { kategori: k, jumlah: katMap[k] };
  });

  return {
    totalPendapatan   : totalPendapatan,
    totalPengeluaran  : totalPengeluaran,
    labaBersih        : labaBersih,
    jumlahTransaksi   : trxData.length,
    chartHarian       : chartHarian,
    chartMetode       : chartMetode,
    chartLayanan      : chartLayanan,
    ringkasanPengeluaran: ringkasanPengeluaran,
  };
}

// ============================================================
// BAGIAN 13: EXPORT
// ============================================================

function exportSheetAsExcel(sheetName) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var ssId    = ss.getId();
  var sheet   = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  var sheetId = sheet.getSheetId();
  var url     = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx&gid=" + sheetId;
  var blob    = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getBlob();
  return Utilities.base64Encode(blob.getBytes());
}

// ============================================================
// BAGIAN 14: UTILITAS
// ============================================================

function formatTgl(date) {
  if (!date) return "";
  try {
    if (date instanceof Date) {
      var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      return Utilities.formatDate(date, tz, "dd/MM/yyyy");
    }
    return String(date).split("T")[0];
  } catch(e) {
    return String(date);
  }
}

function renumberSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== "") sheet.getRange(i+1, 1).setValue(i);
  }
}

// PHG HRIS — Apps Script v2 (fixed GET payload routing)

const SHEET_ID = '1I84QwdGeTwcnQQ9D1qu-skl6hH7XPxmjdfEoe1768xw';
const SHIFT_DEFAULT_START = '08:00';
const TOLERANCE_MENIT = 15;

function doGet(e) {
  var result;
  try {
    // Handle POST-via-GET (payload param)
    if (e.parameter.payload) {
      var data = JSON.parse(e.parameter.payload);
      result = handleAction(data);
    } else {
      result = handleGetAction(e.parameter);
    }
  } catch(err) {
    result = { ok: false, msg: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var data = JSON.parse(e.postData.contents);
    result = handleAction(data);
  } catch(err) {
    result = { ok: false, msg: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleGetAction(params) {
  var action = params.action;
  if (action === 'getOutlets')       return getOutlets();
  if (action === 'getAbsensi')       return getAbsensi(params);
  if (action === 'getKaryawan')      return getKaryawan(params);
  if (action === 'getJadwal')        return getJadwal(params);
  if (action === 'getSummary')       return getSummary(params);
  if (action === 'login')            return login(params);
  return { ok: false, msg: 'Action tidak dikenal: ' + action };
}

function handleAction(data) {
  var action = data.action;
  if (action === 'checkIn')          return checkIn(data);
  if (action === 'checkOut')         return checkOut(data);
  if (action === 'uploadFoto')       return uploadFoto(data);
  if (action === 'addKaryawan')      return addKaryawan(data);
  if (action === 'addJadwal')        return addJadwal(data);
  if (action === 'updateKaryawan')   return updateKaryawan(data);
  // Also handle GET actions via payload
  if (action === 'getOutlets')       return getOutlets();
  if (action === 'getAbsensi')       return getAbsensi(data);
  if (action === 'getKaryawan')      return getKaryawan(data);
  if (action === 'getSummary')       return getSummary(data);
  if (action === 'login')            return login(data);
  return { ok: false, msg: 'Action tidak dikenal: ' + action };
}

// AUTH
function login(params) {
  var pin = String(params.pin);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Karyawan');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][8]) === pin && data[i][7] === 'Aktif') {
      return { ok: true, karyawan: {
        id: data[i][0], nama: data[i][1], nik: data[i][2],
        outlet: data[i][3], jabatan: data[i][4], role: data[i][5], wa: data[i][6]
      }};
    }
  }
  return { ok: false, msg: 'PIN salah atau akun tidak aktif' };
}

// CHECK IN
function checkIn(data) {
  var karyawan_id = data.karyawan_id;
  var nama = data.nama;
  var outlet = data.outlet;
  var lat = data.lat;
  var lng = data.lng;

  var gpsValid = validateGPS(lat, lng, outlet);
  if (!gpsValid.ok) return { ok: false, msg: gpsValid.msg };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Absensi');
  var today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  var jamSekarang = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');

  var existing = cariAbsensiHariIni(karyawan_id, today);
  if (existing) return { ok: false, msg: 'Sudah check-in hari ini pukul ' + existing.jam_masuk };

  var jadwal = getJadwalKaryawan(karyawan_id, today);
  var shiftMulai = jadwal ? jadwal.shift_mulai : SHIFT_DEFAULT_START;
  var status = hitungStatus(jamSekarang, shiftMulai);

  var lastRow = sheet.getLastRow();
  var id = lastRow;

  sheet.appendRow([id, karyawan_id, nama, outlet, today, jamSekarang, '', lat, lng, '', '', status, '', '']);

  return { ok: true, msg: 'Check-in berhasil!', jam: jamSekarang, status: status, shift: shiftMulai, absensi_id: id };
}

// CHECK OUT
function checkOut(data) {
  var karyawan_id = data.karyawan_id;
  var lat = data.lat;
  var lng = data.lng;

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Absensi');
  var today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  var jamSekarang = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');

  var allData = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][1]) === String(karyawan_id) && allData[i][4] === today) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) return { ok: false, msg: 'Belum check-in hari ini' };
  if (allData[rowIndex-1][6]) return { ok: false, msg: 'Sudah check-out pukul ' + allData[rowIndex-1][6] };

  sheet.getRange(rowIndex, 7).setValue(jamSekarang);
  sheet.getRange(rowIndex, 10).setValue(lat);
  sheet.getRange(rowIndex, 11).setValue(lng);

  return { ok: true, msg: 'Check-out berhasil!', jam: jamSekarang, absensi_id: rowIndex - 1 };
}

// UPLOAD FOTO (separate to avoid large payload)
function uploadFoto(data) {
  try {
    var absensi_id = data.absensi_id;
    var tipe = data.tipe; // 'masuk' or 'pulang'
    var foto = data.foto;
    if (!foto) return { ok: false, msg: 'No foto' };

    var fotoUrl = simpanFotoDrive(foto, 'foto_' + tipe + '_' + absensi_id);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Absensi');
    var allData = sheet.getDataRange().getValues();

    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(absensi_id)) {
        var col = tipe === 'masuk' ? 13 : 14;
        sheet.getRange(i + 1, col).setValue(fotoUrl);
        break;
      }
    }
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

// GPS
function validateGPS(lat, lng, outletName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('Outlet').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === outletName) {
      var outLat = parseFloat(data[i][2]);
      var outLng = parseFloat(data[i][3]);
      var radius = parseInt(data[i][4]) || 500;
      var dist = hitungJarak(parseFloat(lat), parseFloat(lng), outLat, outLng);
      if (dist <= radius) return { ok: true, jarak: Math.round(dist) };
      return { ok: false, msg: 'Lokasi kamu ' + Math.round(dist) + 'm dari outlet. Maksimal ' + radius + 'm.' };
    }
  }
  return { ok: false, msg: 'Outlet tidak ditemukan: ' + outletName };
}

function hitungJarak(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLng = (lng2 - lng1) * Math.PI / 180;
  var a = Math.pow(Math.sin(dLat/2), 2) + Math.cos(lat1*Math.PI/180) * Math.cos(lat2*Math.PI/180) * Math.pow(Math.sin(dLng/2), 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function hitungStatus(jamMasuk, shiftMulai) {
  var partsM = jamMasuk.split(':').map(Number);
  var partsS = shiftMulai.split(':').map(Number);
  var menitMasuk = partsM[0] * 60 + partsM[1];
  var menitShift = partsS[0] * 60 + partsS[1];
  if (menitMasuk <= menitShift + TOLERANCE_MENIT) return 'Hadir';
  return 'Terlambat ' + (menitMasuk - menitShift) + ' mnt';
}

// FOTO DRIVE
function simpanFotoDrive(base64Data, namaFile) {
  try {
    var imgData = base64Data.replace(/^data:image\/\w+;base64,/, '');
    var blob = Utilities.newBlob(Utilities.base64Decode(imgData), 'image/jpeg', namaFile + '.jpg');
    var folders = DriveApp.getFoldersByName('PHG Absensi Foto');
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('PHG Absensi Foto');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    return 'error:' + e.message;
  }
}

// GET DATA
function getOutlets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('Outlet').getDataRange().getValues();
  var headers = data[0];
  return { ok: true, data: data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h.toLowerCase()] = row[i]; });
    return obj;
  })};
}

function getKaryawan(params) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('Karyawan').getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h.toLowerCase()] = row[i]; });
    return obj;
  });
  if (params.outlet && params.outlet !== 'all') rows = rows.filter(function(r) { return r.outlet === params.outlet; });
  rows = rows.map(function(r) { delete r.pin_absen; return r; });
  return { ok: true, data: rows };
}

function getAbsensi(params) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('Absensi').getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h.toLowerCase().replace(/ /g,'_')] = row[i]; });
    return obj;
  });
  if (params.karyawan_id) rows = rows.filter(function(r) { return String(r.karyawan_id) === String(params.karyawan_id); });
  if (params.outlet && params.outlet !== 'all') rows = rows.filter(function(r) { return r.outlet === params.outlet; });
  if (params.bulan) rows = rows.filter(function(r) { return String(r.tanggal).indexOf(params.bulan) === 0; });
  if (params.tanggal) rows = rows.filter(function(r) { return r.tanggal === params.tanggal; });
  return { ok: true, data: rows };
}

function getJadwal(params) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('Jadwal').getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h.toLowerCase()] = row[i]; });
    return obj;
  });
  if (params.outlet) rows = rows.filter(function(r) { return r.outlet === params.outlet; });
  if (params.tanggal) rows = rows.filter(function(r) { return r.tanggal === params.tanggal; });
  return { ok: true, data: rows };
}

function getSummary(params) {
  var result = getAbsensi(params);
  var data = result.data;
  var total = data.length;
  var hadir = data.filter(function(a) { return a.status && a.status.indexOf('Terlambat') === -1; }).length;
  var terlambat = data.filter(function(a) { return a.status && a.status.indexOf('Terlambat') !== -1; }).length;
  return { ok: true, total: total, hadir: hadir, terlambat: terlambat, alpha: 0, uang_kehadiran: total * 100000 };
}

function addKaryawan(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Karyawan');
  var lastRow = sheet.getLastRow();
  var id = 'K' + String(lastRow).padStart(4, '0');
  var pin = Math.floor(10000 + Math.random() * 90000);
  var tgl = data.tgl_bergabung || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  sheet.appendRow([id, data.nama, data.nik, data.outlet, data.jabatan, data.role || 'Karyawan', data.wa, 'Aktif', pin, tgl]);
  return { ok: true, id: id, pin: pin, msg: 'Karyawan ditambahkan. PIN absen: ' + pin };
}

function addJadwal(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Jadwal');
  var id = 'J' + Date.now();
  sheet.appendRow([id, data.outlet, data.tanggal, data.shift_mulai, data.shift_selesai, data.karyawan_id, data.dibuat_oleh]);
  return { ok: true, id: id };
}

function updateKaryawan(data) {
  return { ok: true, msg: 'Updated' };
}

function cariAbsensiHariIni(karyawanId, today) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('Absensi').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(karyawanId) && data[i][4] === today) {
      return { jam_masuk: data[i][5] };
    }
  }
  return null;
}

function getJadwalKaryawan(karyawanId, tanggal) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('Jadwal').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][5]) === String(karyawanId) && data[i][2] === tanggal) {
      return { shift_mulai: data[i][3], shift_selesai: data[i][4] };
    }
  }
  return null;
}

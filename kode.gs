/*************************************************************
 * SISTEM LAPORAN & JOBDESK IAT ELECTRICAL — BACKEND + MIRROR SHEET
 *
 * Gabungan:
 *  - Inti API & builder Gantt (Jobdesk) + Visual Laporan  (versi referensi)
 *  - Sheet tampilan tambahan VIEW ... (dashboard, kalender, dll)
 *
 * CARA PAKAI
 * 1. Spreadsheet -> Extensions -> Apps Script, paste seluruh file ini, simpan.
 * 2. Jalankan setupSistem() satu kali (beri izin akses).
 * 3. Deploy -> New deployment -> Web app (Execute as: Me, Access: Anyone).
 *    Tempel URL /exec ke WEB_APP_URL pada file HTML web.
 * 4. Menu "IAT Electrical" muncul di Spreadsheet untuk refresh tampilan.
 *
 * SHEET DATA : Projects, Categories, Accounts, Jobdesk, Reports,
 *              Jadwal Harian, Arsip PDF
 * SHEET MIRIP WEB : Jobdesk (Gantt), Visual Laporan (komparasi laporan),
 *              VIEW Dashboard, VIEW Kalender, VIEW Timeline, VIEW Komparasi,
 *              VIEW Tugas PIC, VIEW Laporan, VIEW Semua Proyek, VIEW Arsip PDF
 *************************************************************/

/* ===================== KONFIGURASI ===================== */

var SHEETS = {
  projects: 'Projects',
  taskCategories: 'Categories',
  accounts: 'Accounts',
  jobdeskTasks: 'Jobdesk',
  reports: 'Reports',
  jadwalHarian: 'Jadwal Harian',
  arsipPDF: 'Arsip PDF'
};

var WARNA = {
  header: '#1e293b',
  headerText: '#ffffff',
  sub: '#f1f5f9',
  border: '#e2e8f0',
  indigo: '#6366f1',
  hijau: '#10b981',
  kuning: '#f59e0b',
  merah: '#ef4444',
  abu: '#94a3b8',
  weekend: '#f1f5f9',
  hariIni: '#fee2e2',
  kertas: '#ffffff',
  zebra: '#f8fafc',
  teks: '#334155'
};

var HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
var BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
             'Agustus', 'September', 'Oktober', 'November', 'Desember'];

var PIC_COLORS = ['#ef4444', '#f97316', '#f59e0b', '#10b981', '#06b6d4',
                  '#3b82f6', '#6366f1', '#8b5cf6', '#d946ef', '#ec4899'];

// Hanya 3 sheet tampilan. Sheet lain sengaja dihilangkan karena isinya sama
// dengan sheet data atau dengan sheet Jobdesk (Gantt) / Visual Laporan.
var VIEW_SHEETS = ['VIEW Dashboard', 'VIEW Kalender', 'VIEW Tugas PIC'];

// Sheet tampilan lama yang sudah digabung -> dihapus otomatis saat refresh.
var VIEW_SHEETS_LAMA = [
  'VIEW Timeline',      // sama dengan sheet Jobdesk (Gantt)
  'VIEW Komparasi',     // jadi bagian di VIEW Dashboard
  'VIEW Laporan',       // sama dengan sheet data Reports
  'VIEW Semua Proyek',  // sama dengan VIEW Tugas PIC
  'VIEW Arsip PDF'      // sama dengan sheet data Arsip PDF
];

/* ===================== MENU & SETUP ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IAT Electrical')
    .addItem('Refresh Semua Tampilan', 'refreshSemuaView')
    .addItem('Bangun Ulang Jobdesk & Visual Laporan', 'rebuildGanttDanLaporan')
    .addItem('Reset Filter Proyek (tampilkan semua)', 'resetFilterProyek')
    .addSeparator()
    .addItem('Setup / Perbaiki Struktur Sheet', 'setupSistem')
    .addItem('Perbaiki Semua Sheet (data + tampilan)', 'perbaikiSemuaSheet')
    .addToUi();
}

/** Kembalikan filter Dashboard ke semua proyek. */
function resetFilterProyek() {
  _setFilterProyek(SEMUA_PROYEK);
  return refreshSemuaView();
}

/** Membuat semua sheet data (lewat getServerData) lalu semua tampilan. */
function setupSistem() {
  getServerData();
  rebuildGanttDanLaporan();
  refreshSemuaView();
  return 'Setup selesai.';
}

/** Bangun ulang sheet Jobdesk (Gantt) dan Visual Laporan dari data sheet. */
function rebuildGanttDanLaporan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tasks = getSheetData(ss, SHEETS.jobdeskTasks);
  var reports = getSheetData(ss, SHEETS.reports);
  reports.forEach(function (r) {
    if (typeof r.workers === 'string') r.workers = r.workers ? r.workers.split(', ') : [];
    if (!r.workers) r.workers = [];
    if (r.workDone === undefined || r.workDone === null) r.workDone = '';
  });
  buildJobdeskGantt(ss, tasks);
  buildVisualLaporan(ss, tasks, reports);
  SpreadsheetApp.flush();
  return 'Jobdesk & Visual Laporan diperbarui.';
}

/** Tampilkan sheet yang tersembunyi, bangun ulang semuanya dari nol. */
function perbaikiSemuaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function (sh) { if (sh.isSheetHidden()) sh.showSheet(); });
  getServerData();
  var pesan = [];
  try { pesan.push(rebuildGanttDanLaporan()); } catch (e) { pesan.push('Gantt/Laporan: ' + e.message); }
  VIEW_SHEETS.concat(VIEW_SHEETS_LAMA).forEach(function (nama) {
    var sh = ss.getSheetByName(nama);
    if (sh) { try { ss.deleteSheet(sh); } catch (x) {} }
  });
  try { pesan.push(refreshSemuaView()); } catch (e) { pesan.push('View: ' + e.message); }
  try { SpreadsheetApp.getUi().alert(pesan.join('\n')); } catch (x) {}
  return pesan.join('\n');
}

/* ===================== REFRESH SEMUA VIEW ===================== */

/** Hapus sheet tampilan lama yang isinya sudah digabung ke sheet lain. */
function _hapusSheetLama(ss) {
  VIEW_SHEETS_LAMA.forEach(function (nama) {
    var sh = ss.getSheetByName(nama);
    if (sh) { try { ss.deleteSheet(sh); } catch (x) {} }
  });
}

function refreshSemuaView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reports = getSheetData(ss, SHEETS.reports);
  reports.forEach(function (r) {
    if (Array.isArray(r.workers)) r.workers = r.workers.join(', ');
    if (r.workers === undefined || r.workers === null) r.workers = '';
  });
  var d = {
    tasks: getSheetData(ss, SHEETS.jobdeskTasks),
    reports: reports,
    jadwal: getSheetData(ss, SHEETS.jadwalHarian),
    projects: getSheetData(ss, SHEETS.projects)
      .map(function (r) { return r.ProjectName || r.projectName || ''; })
      .filter(Boolean)
  };
  _hapusSheetLama(ss);
  var daftar = [
    ['VIEW Dashboard', _viewDashboard],
    ['VIEW Kalender', _viewKalender],
    ['VIEW Tugas PIC', _viewTugasPIC]
  ];
  var gagal = [];
  daftar.forEach(function (item) {
    try {
      item[1](d);
    } catch (err) {
      gagal.push(item[0] + ': ' + err.message);
      try {
        var sh = _viewSheet(item[0], 8, 20);
        sh.getRange(1, 1).setValue('Gagal membuat tampilan ini: ' + err.message)
          .setFontColor(WARNA.merah).setFontWeight('bold');
      } catch (x) {}
    }
  });
  // Bangun ulang Dashboard sekali lagi supaya tombol navigasi menunjuk ke
  // sheet VIEW yang baru dibuat pada putaran pertama.
  try { _viewDashboard(d); } catch (x) {}
  SpreadsheetApp.flush();
  return gagal.length ? 'Selesai dengan catatan:\n' + gagal.join('\n') : 'Semua tampilan berhasil diperbarui.';
}

/* ===================== INTI API & BUILDER (REFERENSI) ===================== */
function doGet(e) {
  const dataStr = getServerData();
  return ContentService.createTextOutput(dataStr)
      .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payloadStr = e.postData.contents;
    const responseStr = processClientRequest(payloadStr);
    return ContentService.createTextOutput(responseStr)
        .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// TRIGGER OTOMATIS SAAT C4 / F4 DIUBAH MANUAL
// ==========================================
function onEdit(e) {
  if (!e || !e.source) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Filter proyek di VIEW Dashboard (B3) -> perbarui semua tampilan.
  if (sheetName === 'VIEW Dashboard' && e.range.getRow() === 3 && e.range.getColumn() === 2) {
    _setFilterProyek(e.range.getValue());
    try { refreshSemuaView(); } catch (x) {}
    return;
  }



  if (sheetName === "Jobdesk" || sheetName === "Visual Laporan") {
    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (row === 4 && (col === 3 || col === 6)) {
      const ss = e.source;
      const newValue = e.range.getValue();

      const targetSheetName = (sheetName === "Jobdesk") ? "Visual Laporan" : "Jobdesk";
      const targetSheet = ss.getSheetByName(targetSheetName);
      if (targetSheet) {
         targetSheet.getRange(row, col).setValue(newValue);
      }

      const jobdeskTasks = getSheetData(ss, "Jobdesk");
      const reports = getSheetData(ss, "Reports");

      buildJobdeskGantt(ss, jobdeskTasks);
      buildVisualLaporan(ss, jobdeskTasks, reports);

      // Sinkronkan juga tampilan VIEW agar ikut mengikuti perubahan
      try { refreshSemuaView(); } catch (x) {}
    }
  }
}

function getServerData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ["Projects", "Categories", "Accounts", "Jobdesk", "Reports", "Jadwal Harian", "Arsip PDF"];

  requiredSheets.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name === "Accounts") {
        sheet.appendRow(["Username", "Password", "Role"]);
        sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
        sheet.appendRow(["admin", "admin123", "admin"]);
      } else if (name === "Projects") {
        sheet.appendRow(["ProjectName"]);
        sheet.getRange("A1").setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
      } else if (name === "Categories") {
        sheet.appendRow(["CategoryName"]);
        sheet.getRange("A1").setFontWeight("bold").setBackground("#10b981").setFontColor("#ffffff");
      } else if (name === "Reports") {
        sheet.appendRow(["ID", "Date", "Project", "Workers", "WorkDone", "Issues", "Photo 1", "Photo 2", "Photo 3", "Photo 4", "Photo 5"]);
        sheet.getRange("A1:K1").setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
      } else if (name === "Jadwal Harian") {
        sheet.appendRow(["No", "Tanggal", "Hari", "Main Job", "Sub Job", "Batas Waktu", "Keterangan", "Nama Pekerja", "Jam Mulai", "Jam Selesai", "Status", "ID"]);
        sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#2563eb").setFontColor("#ffffff");
      } else if (name === "Arsip PDF") {
        sheet.appendRow(["Tanggal", "Nama File", "Link Drive"]);
        sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#10b981").setFontColor("#ffffff");
      }
    }
  });

  const data = {
    projects: getSheetData(ss, "Projects").map(row => row.ProjectName || row.projectname || row["Project Name"] || row[0]).filter(Boolean),
    taskCategories: getSheetData(ss, "Categories").map(row => row.CategoryName || row.categoryname || row["Category Name"] || row[0]).filter(Boolean),
    accounts: getSheetData(ss, "Accounts"),
    jobdeskTasks: getSheetData(ss, "Jobdesk"),
    reports: getSheetData(ss, "Reports"),
    jadwalHarian: getSheetData(ss, "Jadwal Harian"),
    arsipPDF: getSheetData(ss, "Arsip PDF")
  };

  data.reports.forEach(report => {
    if (typeof report.workers === 'string' && report.workers !== "") {
      report.workers = report.workers.split(', ');
    } else if (!report.workers || report.workers === "") {
      report.workers = [];
    }
  });

  return JSON.stringify(data);
}

function processClientRequest(payloadStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let response = { status: "success" };

  try {
    const payload = JSON.parse(payloadStr);
    const action = payload.action;
    const data = payload.data;

    if (action === "uploadPhoto") {
      var folderName = "Foto_Laporan_IAT";
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var fileBlob = Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType, data.fileName);
      var file = folder.createFile(fileBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      response = { status: "success", fileUrl: file.getUrl() };
      return JSON.stringify(response);
    }

    if (action === "archivePDF") {
      var folderName = "Arsip_PDF_IAT";
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var fileBlob = Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType, data.fileName);
      var file = folder.createFile(fileBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      var arsipSheet = ss.getSheetByName("Arsip PDF");
      const timestamp = new Date().toLocaleString('id-ID');
      arsipSheet.insertRowAfter(1);
      arsipSheet.getRange(2, 1, 1, 3).setValues([[timestamp, data.fileName, file.getUrl()]]);

      response = { status: "success", fileUrl: file.getUrl() };
      return JSON.stringify(response);
    }

    if (action === "syncAll") {
      const projectRows = data.projects.map(p => [p]);
      updateSheetData(ss, "Projects", ["ProjectName"], projectRows);

      if (data.taskCategories) {
        const categoryRows = data.taskCategories.map(c => [c]);
        updateSheetData(ss, "Categories", ["CategoryName"], categoryRows);
      }

      if (data.accounts && data.accounts.length > 0) {
        const accRows = data.accounts.map(a => [a.username || "", a.password || "", a.role || ""]);
        updateSheetData(ss, "Accounts", ["Username", "Password", "Role"], accRows);
      }

      buildJobdeskGantt(ss, data.jobdeskTasks);

      if (data.jadwalHarian) {
        const jhRows = data.jadwalHarian.map((j, idx) => [
          idx + 1, j.tanggal || "", j.hari || "", j.mainJob || "", j.subJob || "",
          j.batasWaktu || "", j.keterangan || "", j.namaPekerja || "",
          j.jamMulai || "", j.jamSelesai || "", j.status || "", j.id || ""
        ]);
        const jhHeaders = ["No", "Tanggal", "Hari", "Main Job", "Sub Job", "Batas Waktu", "Keterangan", "Nama Pekerja", "Jam Mulai", "Jam Selesai", "Status", "ID"];
        updateSheetData(ss, "Jadwal Harian", jhHeaders, jhRows);
        formatJadwalHarian(ss);
      }

      const reportRows = (data.reports || []).map(r => [
        r.id, r.date, r.project, (Array.isArray(r.workers) ? r.workers.join(", ") : (r.workers || "")), r.workDone, r.issues,
        r.photo1 || "", r.photo2 || "", r.photo3 || "", r.photo4 || "", r.photo5 || ""
      ]);
      updateSheetData(ss, "Reports", ["ID", "Date", "Project", "Workers", "WorkDone", "Issues", "Photo 1", "Photo 2", "Photo 3", "Photo 4", "Photo 5"], reportRows);

      buildVisualLaporan(ss, data.jobdeskTasks, data.reports);

      // Perbarui juga sheet tampilan VIEW ... agar isinya mengikuti web
      try { refreshSemuaView(); } catch (x) {}
    }
  } catch (error) {
    response = { status: "error", message: error.toString() };
  }

  return JSON.stringify(response);
}

function formatJadwalHarian(ss) {
  const sheet = ss.getSheetByName("Jadwal Harian");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow > 0 && lastCol > 0) {
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    sheet.getRange(1, 1, 1, lastCol).setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).setVerticalAlignment("middle");
      sheet.getRange(2, 1, lastRow - 1, 3).setHorizontalAlignment("center");
      sheet.getRange(2, 6, lastRow - 1, 1).setHorizontalAlignment("center");
      sheet.getRange(2, 9, lastRow - 1, 3).setHorizontalAlignment("center");
      sheet.getRange(2, 4, lastRow - 1, 2).setHorizontalAlignment("left");
      sheet.getRange(2, 7, lastRow - 1, 2).setHorizontalAlignment("left");
      sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat("yyyy-MM-dd");
    }
    range.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    sheet.setColumnWidth(1, 40); sheet.setColumnWidth(2, 90); sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 200); sheet.setColumnWidth(5, 200); sheet.setColumnWidth(6, 90);
    sheet.setColumnWidth(7, 250); sheet.setColumnWidth(8, 150); sheet.setColumnWidth(9, 90);
    sheet.setColumnWidth(10, 90); sheet.setColumnWidth(11, 120); sheet.hideColumns(12, 1);
    sheet.setFrozenRows(1);
  }
}

// ==============================
// HELPER BERSAMA GANTT / VISUAL LAPORAN
// (dipakai buildJobdeskGantt & buildVisualLaporan — jangan diduplikasi lagi)
// ==============================

/** Baca "Project Start Date" (C4) & "Display Week" (F4) dari sheet tampilan. */
function _ganttPengaturan(sheet) {
  var out = { manualStartDate: null, displayWeek: 1 };
  if (!sheet) return out;
  var dataArr = sheet.getDataRange().getValues();
  if (dataArr.length >= 4) {
    var sd = dataArr[3][2];
    if (sd) out.manualStartDate = parseDateLocal(sd);
    var dw = dataArr[3][5];
    if (dw && !isNaN(parseInt(dw))) out.displayWeek = parseInt(dw);
  }
  return out;
}

/** Susun hirarki Proyek > Tugas > Sub-tugas beserta tanggal min/max. */
function _ganttHirarki(jobdeskTasks) {
  var hierarchy = {}, minD = null, maxD = null;
  (jobdeskTasks || []).forEach(function (t) {
    var pK = (t.projectName || 'Lain-lain').trim().toUpperCase();
    var tK = (t.name || 'Tugas').trim().toUpperCase();
    if (!hierarchy[pK]) hierarchy[pK] = { name: t.projectName || 'Lain-lain', tasks: {}, minDate: null };
    if (!hierarchy[pK].tasks[tK]) hierarchy[pK].tasks[tK] = { name: t.name, items: [], minDate: null };
    hierarchy[pK].tasks[tK].items.push(t);

    var dStart = parseDateLocal(t.startDate);
    var dEnd = parseDateLocal(t.endDate);
    if (dStart) {
      if (!minD || dStart < minD) minD = new Date(dStart.getTime());
      if (!hierarchy[pK].minDate || dStart < hierarchy[pK].minDate) hierarchy[pK].minDate = new Date(dStart.getTime());
      if (!hierarchy[pK].tasks[tK].minDate || dStart < hierarchy[pK].tasks[tK].minDate) hierarchy[pK].tasks[tK].minDate = new Date(dStart.getTime());
    }
    if (dEnd && (!maxD || dEnd > maxD)) maxD = new Date(dEnd.getTime());
  });
  return { hierarchy: hierarchy, minD: minD, maxD: maxD };
}

/** Hitung tanggal awal proyek dan deretan tanggal kolom kalender. */
function _ganttKalender(manualStartDate, minD, maxD, displayWeek) {
  var baseStartDate = manualStartDate ? new Date(manualStartDate.getTime())
    : (minD ? new Date(minD.getTime()) : new Date());
  baseStartDate.setHours(12, 0, 0, 0);
  if (!maxD) maxD = new Date(baseStartDate.getTime());

  var startDay = baseStartDate.getDay();
  var diffM = baseStartDate.getDate() - startDay + (startDay === 0 ? -6 : 1);
  var week1Monday = new Date(baseStartDate.getFullYear(), baseStartDate.getMonth(), diffM, 12, 0, 0);

  var startCal = new Date(week1Monday.getTime());
  startCal.setDate(startCal.getDate() + ((displayWeek - 1) * 7));

  var endCal = new Date(startCal.getTime());
  endCal.setDate(endCal.getDate() + 55);
  if (maxD > endCal) {
    var endDay = maxD.getDay();
    var diffS = maxD.getDate() + (endDay === 0 ? 0 : 7 - endDay);
    endCal = new Date(maxD.getFullYear(), maxD.getMonth(), diffS, 12, 0, 0);
  }

  var dateArr = [];
  var cx = new Date(startCal.getTime());
  while (cx <= endCal) {
    dateArr.push(new Date(cx.getFullYear(), cx.getMonth(), cx.getDate(), 12, 0, 0));
    cx.setDate(cx.getDate() + 1);
  }
  return { baseStartDate: baseStartDate, dateArr: dateArr };
}

/** 7 baris kepala (judul, info proyek, minggu, tanggal, hari, header tabel). */
function _ganttBarisKepala(totalCols, dateArr, baseStartDate, displayWeek, opsi) {
  var today = new Date(); today.setHours(12, 0, 0, 0);
  var v = [], b = [], f = [], c = [];
  function baris(bg, fw, fc) {
    var row = {
      v: Array(totalCols).fill(''), b: Array(totalCols).fill(bg),
      f: Array(totalCols).fill(fw), c: Array(totalCols).fill(fc)
    };
    v.push(row.v); b.push(row.b); f.push(row.f); c.push(row.c);
    return row;
  }

  var r1 = baris('#ffffff', 'bold', '#1e293b'); r1.v[0] = opsi.judul1 || '';
  var r2 = baris('#ffffff', 'bold', opsi.warnaJudul2 || '#1e293b'); r2.v[0] = opsi.judul2 || '';
  baris('#ffffff', 'normal', '#000000');

  var r4 = baris('#ffffff', 'bold', '#1e293b');
  r4.v[0] = 'Project Start Date'; r4.v[2] = baseStartDate;
  r4.v[4] = 'Display Week'; r4.v[5] = displayWeek;
  for (var w = 0; w < dateArr.length / 7; w++) {
    r4.v[10 + w * 7] = 'Week ' + (displayWeek + w);
    r4.b[10 + w * 7] = '#f8fafc';
  }

  var r5 = baris('#ffffff', 'bold', '#1e293b');
  r5.v[0] = 'Project Lead'; r5.v[2] = 'Team IAT';
  for (var w2 = 0; w2 < dateArr.length / 7; w2++) {
    r5.v[10 + w2 * 7] = dateArr[w2 * 7] || '';
    r5.c[10 + w2 * 7] = '#475569';
  }

  var r6 = baris('#f1f5f9', 'bold', '#334155');
  dateArr.forEach(function (d, i) {
    var isT = (d.getTime() === today.getTime());
    r6.v[10 + i] = d.getDate();
    r6.b[10 + i] = isT ? '#ef4444' : '#e2e8f0';
    r6.c[10 + i] = isT ? '#ffffff' : '#1e293b';
  });

  var r7 = baris(opsi.warnaHeader, 'bold', '#ffffff');
  (opsi.headers || []).forEach(function (h, i) { r7.v[i] = h; });
  var dayNames = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
  dateArr.forEach(function (d, i) {
    var isT = (d.getTime() === today.getTime());
    r7.v[10 + i] = dayNames[d.getDay()];
    r7.b[10 + i] = isT ? '#ef4444' : opsi.warnaHari;
  });
  r7.v[10 + dateArr.length] = '_ID';
  r7.v[10 + dateArr.length + 1] = '_Status';
  r7.v[10 + dateArr.length + 2] = '_Meta';

  return { v: v, b: b, f: f, c: c };
}

// ==============================
// GANTT CHART BUILDER (JOBDESK)
// ==============================
function buildJobdeskGantt(ss, jobdeskTasks) {
  if (!jobdeskTasks || jobdeskTasks.length === 0) return;
  let jdSheet = ss.getSheetByName("Jobdesk");

  const cfg = _ganttPengaturan(jdSheet);
  const manualStartDate = cfg.manualStartDate;
  const displayWeek = cfg.displayWeek;

  const hier = _ganttHirarki(jobdeskTasks);
  const hierarchy = hier.hierarchy;
  const kal = _ganttKalender(manualStartDate, hier.minD, hier.maxD, displayWeek);
  const baseStartDate = kal.baseStartDate;
  const dateArr = kal.dateArr;

  let today = new Date(); today.setHours(12,0,0,0);
  let totalCols = 10 + dateArr.length + 3;

  const kepala = _ganttBarisKepala(totalCols, dateArr, baseStartDate, displayWeek, {
    judul1: "IAT Project Electrical Instaler",
    judul2: "INDOSPRING TBK",
    warnaHeader: "#4f46e5",
    warnaHari: "#4338ca",
    headers: ["NO", "TASK/PROYEK", "PIC", "START", "END", "HARI", "%", "", "", ""]
  });
  let valM = kepala.v, bgM = kepala.b, fwM = kepala.f, fcM = kepala.c;

  const createRow = (val, bg, fw, fc) => {
      return { v: Array(totalCols).fill(val), b: Array(totalCols).fill(bg), f: Array(totalCols).fill(fw), c: Array(totalCols).fill(fc) };
  };

  const getC = _warnaPIC; // satu sumber warna PIC (PIC_COLORS)

  const sortedProjects = Object.values(hierarchy).sort((a, b) => (a.minDate || 0) - (b.minDate || 0));
  let pC = 1;
  sortedProjects.forEach(proj => {
    let allP = [];
    for (let tK in proj.tasks) allP = allP.concat(proj.tasks[tK].items);
    let pMinD = allP.map(x=>parseDateLocal(x.startDate)).filter(Boolean).sort((a,b)=>a-b)[0] || null;
    let pMaxD = allP.map(x=>parseDateLocal(x.endDate)).filter(Boolean).sort((a,b)=>b-a)[0] || null;
    let pAvg = Math.round(allP.reduce((s,x)=>s+(parseInt(x.progress)||0),0)/allP.length) || 0;
    let pDays = getDaysDiff(pMinD, pMaxD);

    let rP = createRow("", "#e2e8f0", "bold", "#1e293b");
    rP.v[0] = String(pC); rP.v[1] = proj.name.toUpperCase();
    rP.v[3] = pMinD || ""; rP.v[4] = pMaxD || "";
    rP.v[5] = pDays || 0; rP.v[6] = pAvg+"%";
    dateArr.forEach((d, i) => {
        rP.f[10+i] = "normal"; rP.c[10+i] = "#000000";
        let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
        if(pMinD && pMaxD && d>=pMinD && d<=pMaxD) rP.b[10+i] = "#cbd5e1"; else rP.b[10+i] = isT ? "#fee2e2" : (w?"#f1f5f9":"#e2e8f0");
    });
    valM.push(rP.v); bgM.push(rP.b); fwM.push(rP.f); fcM.push(rP.c);

    const sortedTasks = Object.values(proj.tasks).sort((a, b) => (a.minDate || 0) - (b.minDate || 0));
    let tC = 1;
    sortedTasks.forEach(tGrp => {
        const items = tGrp.items;
        const isStand = items.length === 1 && !items[0].subTask;
        let tMinD = items.map(x=>parseDateLocal(x.startDate)).filter(Boolean).sort((a,b)=>a-b)[0] || null;
        let tMaxD = items.map(x=>parseDateLocal(x.endDate)).filter(Boolean).sort((a,b)=>b-a)[0] || null;
        let tAvg = Math.round(items.reduce((s,x)=>s+(parseInt(x.progress)||0),0)/items.length) || 0;
        let tDays = getDaysDiff(tMinD, tMaxD);

        if (isStand) {
            const tk = items[0];
            let sD = parseDateLocal(tk.startDate); let eD = parseDateLocal(tk.endDate);
            let metaData = JSON.stringify({projectName: proj.name, category: tk.category || "", taskName: tk.name, subTask: ""});
            let tr = createRow("", "#ffffff", "normal", "#334155");
            tr.v[0] = `${pC}.${tC}`; tr.v[1] = tGrp.name; tr.v[2] = tk.pic;
            tr.v[3] = sD || ""; tr.v[4] = eD || "";
            tr.v[5] = tDays || 0; tr.v[6] = tk.progress+"%";
            dateArr.forEach((d, i) => {
                tr.f[10+i] = "normal"; tr.c[10+i] = "#000000";
                let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
                if(sD && eD && d>=sD && d<=eD) tr.b[10+i] = getC(tk.pic); else tr.b[10+i] = isT ? "#fee2e2" : (w?"#f8fafc":"#ffffff");
            });
            tr.v[10 + dateArr.length] = tk.id; tr.v[10 + dateArr.length + 1] = tk.status; tr.v[10 + dateArr.length + 2] = metaData;
            valM.push(tr.v); bgM.push(tr.b); fwM.push(tr.f); fcM.push(tr.c);
        } else {
            let tr = createRow("", "#f8fafc", "bold", "#334155");
            tr.v[0] = `${pC}.${tC}`; tr.v[1] = tGrp.name;
            tr.v[3] = tMinD || ""; tr.v[4] = tMaxD || "";
            tr.v[5] = tDays || 0; tr.v[6] = tAvg+"%";
            dateArr.forEach((d, i) => {
                tr.f[10+i] = "normal"; tr.c[10+i] = "#000000";
                let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
                if(tMinD && tMaxD && d>=tMinD && d<=tMaxD) tr.b[10+i] = "#94a3b8"; else tr.b[10+i] = isT ? "#fee2e2" : (w?"#f1f5f9":"#f8fafc");
            });
            valM.push(tr.v); bgM.push(tr.b); fwM.push(tr.f); fcM.push(tr.c);

            items.sort((a, b) => { return (parseDateLocal(a.startDate)||0) - (parseDateLocal(b.startDate)||0); });
            let sC = 1;
            items.forEach(tk => {
                let sD = parseDateLocal(tk.startDate); let eD = parseDateLocal(tk.endDate);
                let metaData = JSON.stringify({projectName: proj.name, category: tk.category || "", taskName: tk.name, subTask: tk.subTask});
                let sr = createRow("", "#ffffff", "normal", "#475569");
                sr.v[0] = `${pC}.${tC}.${sC}`; sr.v[1] = (tk.subTask ? "↳ " + tk.subTask : "↳ Eksekusi"); sr.v[2] = tk.pic;
                sr.v[3] = sD || ""; sr.v[4] = eD || "";
                sr.v[5] = getDaysDiff(sD, eD) || 0; sr.v[6] = tk.progress+"%";
                dateArr.forEach((d, i) => {
                    sr.f[10+i] = "normal"; sr.c[10+i] = "#000000";
                    let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
                    if(sD && eD && d>=sD && d<=eD) sr.b[10+i] = getC(tk.pic); else sr.b[10+i] = isT ? "#fee2e2" : (w?"#f8fafc":"#ffffff");
                });
                sr.v[10 + dateArr.length] = tk.id; sr.v[10 + dateArr.length + 1] = tk.status; sr.v[10 + dateArr.length + 2] = metaData;
                valM.push(sr.v); bgM.push(sr.b); fwM.push(sr.f); fcM.push(sr.c);
                sC++;
            });
        }
        tC++;
    });
    pC++;
  });

  if(!jdSheet) jdSheet = ss.insertSheet("Jobdesk");
  let maxCols = jdSheet.getMaxColumns();
  if (maxCols < totalCols) { jdSheet.insertColumnsAfter(maxCols, totalCols - maxCols); }
  let maxRows = jdSheet.getMaxRows();
  if (maxRows < valM.length) { jdSheet.insertRowsAfter(maxRows, valM.length - maxRows); }

  // Bongkar merge lama & unhide kolom: clear() tidak membongkar merge sehingga
  // merge ulang bisa gagal dan membuat sheet tidak bisa dibuka.
  try { jdSheet.getRange(1, 1, jdSheet.getMaxRows(), jdSheet.getMaxColumns()).breakApart(); } catch (x) {}
  try { jdSheet.showColumns(1, jdSheet.getMaxColumns()); } catch (x) {}
  jdSheet.setFrozenRows(0); jdSheet.setFrozenColumns(0);
  jdSheet.clear();
  jdSheet.clearFormats();

  if(valM.length > 0) {
      let r = jdSheet.getRange(1, 1, valM.length, totalCols);
      if (valM.length > 6) jdSheet.getRange(7, 1, valM.length - 6, 1).setNumberFormat("@");

      r.setValues(valM);
      r.setBackgrounds(bgM);
      r.setFontWeights(fwM);
      r.setFontColors(fcM);
      r.setVerticalAlignment("middle");

      if (valM.length > 6) {
         jdSheet.getRange(7, 1, valM.length-6, 7).setHorizontalAlignment("center");
         jdSheet.getRange(7, 2, valM.length-6, 1).setHorizontalAlignment("left");
      }
      if (valM.length > 3) jdSheet.getRange(4, 11, valM.length-3, dateArr.length).setHorizontalAlignment("center");

      if (valM.length > 7) {
         jdSheet.getRange(8, 4, valM.length-7, 2).setNumberFormat("yyyy-MM-dd");
         jdSheet.getRange(8, 6, valM.length-7, 1).setNumberFormat("0");
      }

      jdSheet.getRange(4, 3).setNumberFormat("yyyy-MM-dd");
      jdSheet.getRange("A1:B1").mergeAcross().setFontSize(16);
      jdSheet.getRange("A2:B2").mergeAcross().setFontSize(12);
      jdSheet.getRange("A4:B4").mergeAcross();
      jdSheet.getRange("A5:B5").mergeAcross();

      for(let w = 0; w < Math.ceil(dateArr.length / 7); w++) {
          let startCol = 11 + w * 7;
          let colsToMerge = Math.min(7, dateArr.length - w * 7);
          if (colsToMerge > 1) {
              jdSheet.getRange(4, startCol, 1, colsToMerge).mergeAcross();
              jdSheet.getRange(5, startCol, 1, colsToMerge).mergeAcross().setNumberFormat("dd/MM/yy");
          } else {
              jdSheet.getRange(5, startCol, 1, 1).setNumberFormat("dd/MM/yy");
          }
      }

      if (valM.length > 5) jdSheet.getRange(6, 1, valM.length-5, totalCols - 3).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

      jdSheet.setColumnWidth(1, 45); jdSheet.setColumnWidth(2, 260); jdSheet.setColumnWidth(3, 70);
      jdSheet.setColumnWidth(4, 75); jdSheet.setColumnWidth(5, 75); jdSheet.setColumnWidth(6, 45);
      jdSheet.setColumnWidth(7, 45); jdSheet.setColumnWidth(8, 30); jdSheet.setColumnWidth(9, 30);
      jdSheet.setColumnWidth(10, 30);
      try { jdSheet.setColumnWidths(11, dateArr.length, 25); } catch(e){}

      try { jdSheet.hideColumns(11 + dateArr.length, 3); } catch(e){}
      try { jdSheet.setFrozenRows(7); jdSheet.setFrozenColumns(2); } catch(e){}
  }
}

// ==============================
// VISUAL LAPORAN BUILDER
// ==============================
function buildVisualLaporan(ss, jobdeskTasks, reports) {
  if (!jobdeskTasks || jobdeskTasks.length === 0) return;
  let vlSheet = ss.getSheetByName("Visual Laporan");
  reports = reports || [];

  const cfg = _ganttPengaturan(vlSheet);
  const manualStartDate = cfg.manualStartDate;
  const displayWeek = cfg.displayWeek;

  const hier = _ganttHirarki(jobdeskTasks);
  const hierarchy = hier.hierarchy;
  const kal = _ganttKalender(manualStartDate, hier.minD, hier.maxD, displayWeek);
  const baseStartDate = kal.baseStartDate;
  const dateArr = kal.dateArr;

  let today = new Date(); today.setHours(12,0,0,0);
  let totalCols = 10 + dateArr.length + 3;

  const kepala = _ganttBarisKepala(totalCols, dateArr, baseStartDate, displayWeek, {
    judul1: "MATRIKS KOMPARASI & VISUAL LAPORAN",
    judul2: "Arahkan Kursor pada simbol (\u2713) untuk melihat detail laporan yang diselesaikan di tanggal tersebut.",
    warnaJudul2: "#10b981",
    warnaHeader: "#059669",
    warnaHari: "#10b981",
    headers: ["NO", "TASK/PROYEK", "PIC", "START", "END", "HARI", "LAPORAN", "", "", ""]
  });
  let valM = kepala.v, bgM = kepala.b, fwM = kepala.f, fcM = kepala.c;
  let notesM = valM.map(function () { return Array(totalCols).fill(""); });

  const createRow = (val, bg, fw, fc) => {
      return { v: Array(totalCols).fill(val), b: Array(totalCols).fill(bg), f: Array(totalCols).fill(fw), c: Array(totalCols).fill(fc), n: Array(totalCols).fill("") };
  };

  const checkReport = (tk, projName, d) => {
      let noteText = ""; let isReported = false;
      let reps = (reports || []).filter(r => r.project === projName && isSameDayAppScript(parseDateLocal(r.date), d));
      reps.forEach((rep, idx) => {
         let match = false;
         let wd = String(rep.workDone || "");
         if (tk.subTask) { if (wd.indexOf(tk.subTask) !== -1 || wd.indexOf(tk.name) !== -1) match = true; }
         else { if (tk.name && wd.indexOf(tk.name) !== -1) match = true; }

         if(match) {
             isReported = true;
             let wkr = Array.isArray(rep.workers) ? rep.workers.join(', ') : rep.workers;
             noteText += `[LAPORAN ${idx+1}]\nPEKERJA: ${wkr}\nPEKERJAAN SELESAI:\n${wd}\n${rep.issues ? '\nKENDALA: '+rep.issues : ''}\n\n`;
         }
      });
      return { isReported, noteText: noteText.trim() };
  };

  const sortedProjects = Object.values(hierarchy).sort((a, b) => (a.minDate || 0) - (b.minDate || 0));
  let pC = 1;
  sortedProjects.forEach(proj => {
    let allP = []; for (let tK in proj.tasks) allP = allP.concat(proj.tasks[tK].items);
    let pMinD = allP.map(x=>parseDateLocal(x.startDate)).filter(Boolean).sort((a,b)=>a-b)[0] || null;
    let pMaxD = allP.map(x=>parseDateLocal(x.endDate)).filter(Boolean).sort((a,b)=>b-a)[0] || null;
    let pDays = getDaysDiff(pMinD, pMaxD);
    let pRepCount = reports.filter(r => r.project === proj.name).length;

    let rP = createRow("", "#e2e8f0", "bold", "#1e293b");
    rP.v[0] = String(pC); rP.v[1] = proj.name.toUpperCase();
    rP.v[3] = pMinD || ""; rP.v[4] = pMaxD || "";
    rP.v[5] = pDays || 0; rP.v[6] = pRepCount + " Lap";
    dateArr.forEach((d, i) => {
        rP.f[10+i] = "normal"; rP.c[10+i] = "#000000";
        let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
        rP.b[10+i] = isT ? "#fee2e2" : (w?"#f1f5f9":"#e2e8f0");
    });
    valM.push(rP.v); bgM.push(rP.b); fwM.push(rP.f); fcM.push(rP.c); notesM.push(rP.n);

    const sortedTasks = Object.values(proj.tasks).sort((a, b) => (a.minDate || 0) - (b.minDate || 0));
    let tC = 1;
    sortedTasks.forEach(tGrp => {
        const items = tGrp.items;
        const isStand = items.length === 1 && !items[0].subTask;
        let tMinD = items.map(x=>parseDateLocal(x.startDate)).filter(Boolean).sort((a,b)=>a-b)[0] || null;
        let tMaxD = items.map(x=>parseDateLocal(x.endDate)).filter(Boolean).sort((a,b)=>b-a)[0] || null;
        let tDays = getDaysDiff(tMinD, tMaxD);

        if (isStand) {
            const tk = items[0];
            let sD = parseDateLocal(tk.startDate); let eD = parseDateLocal(tk.endDate);
            let metaData = JSON.stringify({projectName: proj.name, category: tk.category || "", taskName: tk.name, subTask: ""});
            let tr = createRow("", "#ffffff", "normal", "#334155");
            tr.v[0] = `${pC}.${tC}`; tr.v[1] = tGrp.name; tr.v[2] = tk.pic;
            tr.v[3] = sD || ""; tr.v[4] = eD || "";
            tr.v[5] = tDays || 0; tr.v[6] = "";
            dateArr.forEach((d, i) => {
                tr.f[10+i] = "normal"; tr.c[10+i] = "#000000";
                let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
                let isP = sD && eD && d>=sD && d<=eD;
                let {isReported, noteText} = checkReport(tk, proj.name, d);

                if(isReported) {
                    tr.v[10+i] = "✓"; tr.b[10+i] = "#10b981"; tr.c[10+i] = "#ffffff"; tr.f[10+i] = "bold"; tr.n[10+i] = noteText;
                } else if(isP) {
                    if (d < today) { tr.v[10+i] = "!"; tr.b[10+i] = "#fecdd3"; tr.c[10+i] = "#e11d48"; tr.f[10+i] = "bold"; tr.n[10+i]="MISSED: Target jadwal berjalan tapi tidak ada laporan"; }
                    else { tr.b[10+i] = "#e2e8f0"; }
                } else { tr.b[10+i] = isT ? "#fee2e2" : (w?"#f8fafc":"#ffffff"); }
            });
            tr.v[10 + dateArr.length] = tk.id; tr.v[10 + dateArr.length + 1] = tk.status; tr.v[10 + dateArr.length + 2] = metaData;
            valM.push(tr.v); bgM.push(tr.b); fwM.push(tr.f); fcM.push(tr.c); notesM.push(tr.n);
        } else {
            let tr = createRow("", "#f8fafc", "bold", "#334155");
            tr.v[0] = `${pC}.${tC}`; tr.v[1] = tGrp.name;
            tr.v[3] = tMinD || ""; tr.v[4] = tMaxD || "";
            tr.v[5] = tDays || 0; tr.v[6] = "";
            dateArr.forEach((d, i) => {
                tr.f[10+i] = "normal"; tr.c[10+i] = "#000000";
                let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
                tr.b[10+i] = isT ? "#fee2e2" : (w?"#f1f5f9":"#f8fafc");
            });
            valM.push(tr.v); bgM.push(tr.b); fwM.push(tr.f); fcM.push(tr.c); notesM.push(tr.n);

            items.sort((a, b) => { return (parseDateLocal(a.startDate)||0) - (parseDateLocal(b.startDate)||0); });
            let sC = 1;
            items.forEach(tk => {
                let sD = parseDateLocal(tk.startDate); let eD = parseDateLocal(tk.endDate);
                let metaData = JSON.stringify({projectName: proj.name, category: tk.category || "", taskName: tk.name, subTask: tk.subTask});
                let sr = createRow("", "#ffffff", "normal", "#475569");
                sr.v[0] = `${pC}.${tC}.${sC}`; sr.v[1] = (tk.subTask ? "↳ " + tk.subTask : "↳ Eksekusi"); sr.v[2] = tk.pic;
                sr.v[3] = sD || ""; sr.v[4] = eD || "";
                sr.v[5] = getDaysDiff(sD, eD) || 0; sr.v[6] = "";
                dateArr.forEach((d, i) => {
                    sr.f[10+i] = "normal"; sr.c[10+i] = "#000000";
                    let w = (d.getDay()===0||d.getDay()===6); let isT = (d.getTime() === today.getTime());
                    let isP = sD && eD && d>=sD && d<=eD;
                    let {isReported, noteText} = checkReport(tk, proj.name, d);

                    if(isReported) {
                        sr.v[10+i] = "✓"; sr.b[10+i] = "#10b981"; sr.c[10+i] = "#ffffff"; sr.f[10+i] = "bold"; sr.n[10+i] = noteText;
                    } else if(isP) {
                        if (d < today) { sr.v[10+i] = "!"; sr.b[10+i] = "#fecdd3"; sr.c[10+i] = "#e11d48"; sr.f[10+i] = "bold"; sr.n[10+i]="MISSED: Target jadwal berjalan tapi tidak ada laporan"; }
                        else { sr.b[10+i] = "#e2e8f0"; }
                    } else { sr.b[10+i] = isT ? "#fee2e2" : (w?"#f8fafc":"#ffffff"); }
                });
                sr.v[10 + dateArr.length] = tk.id; sr.v[10 + dateArr.length + 1] = tk.status; sr.v[10 + dateArr.length + 2] = metaData;
                valM.push(sr.v); bgM.push(sr.b); fwM.push(sr.f); fcM.push(sr.c); notesM.push(sr.n);
                sC++;
            });
        }
        tC++;
    });
    pC++;
  });

  if(!vlSheet) vlSheet = ss.insertSheet("Visual Laporan");
  let maxCols = vlSheet.getMaxColumns();
  if (maxCols < totalCols) { vlSheet.insertColumnsAfter(maxCols, totalCols - maxCols); }
  let maxRows = vlSheet.getMaxRows();
  if (maxRows < valM.length) { vlSheet.insertRowsAfter(maxRows, valM.length - maxRows); }

  try { vlSheet.getRange(1, 1, vlSheet.getMaxRows(), vlSheet.getMaxColumns()).breakApart(); } catch (x) {}
  try { vlSheet.showColumns(1, vlSheet.getMaxColumns()); } catch (x) {}
  vlSheet.setFrozenRows(0); vlSheet.setFrozenColumns(0);
  vlSheet.clear(); vlSheet.clearNotes(); vlSheet.clearFormats();

  if(valM.length > 0) {
      let r = vlSheet.getRange(1, 1, valM.length, totalCols);
      if (valM.length > 6) vlSheet.getRange(7, 1, valM.length - 6, 1).setNumberFormat("@");

      r.setValues(valM); r.setBackgrounds(bgM); r.setFontWeights(fwM); r.setFontColors(fcM);
      r.setNotes(notesM);
      r.setVerticalAlignment("middle");

      if (valM.length > 6) {
         vlSheet.getRange(7, 1, valM.length-6, 7).setHorizontalAlignment("center");
         vlSheet.getRange(7, 2, valM.length-6, 1).setHorizontalAlignment("left");
      }
      if (valM.length > 3) vlSheet.getRange(4, 11, valM.length-3, dateArr.length).setHorizontalAlignment("center");
      if (valM.length > 7) {
         vlSheet.getRange(8, 4, valM.length-7, 2).setNumberFormat("yyyy-MM-dd");
         vlSheet.getRange(8, 6, valM.length-7, 1).setNumberFormat("0");
      }

      vlSheet.getRange(4, 3).setNumberFormat("yyyy-MM-dd");
      vlSheet.getRange("A1:B1").mergeAcross().setFontSize(16);
      vlSheet.getRange("A2:B2").mergeAcross().setFontSize(12).setFontStyle("italic");
      vlSheet.getRange("A4:B4").mergeAcross(); vlSheet.getRange("A5:B5").mergeAcross();

      for(let w = 0; w < Math.ceil(dateArr.length / 7); w++) {
          let startCol = 11 + w * 7;
          let colsToMerge = Math.min(7, dateArr.length - w * 7);
          if (colsToMerge > 1) {
              vlSheet.getRange(4, startCol, 1, colsToMerge).mergeAcross();
              vlSheet.getRange(5, startCol, 1, colsToMerge).mergeAcross().setNumberFormat("dd/MM/yy");
          } else { vlSheet.getRange(5, startCol, 1, 1).setNumberFormat("dd/MM/yy"); }
      }

      if (valM.length > 5) vlSheet.getRange(6, 1, valM.length-5, totalCols - 3).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

      vlSheet.setColumnWidth(1, 45); vlSheet.setColumnWidth(2, 260); vlSheet.setColumnWidth(3, 70);
      vlSheet.setColumnWidth(4, 75); vlSheet.setColumnWidth(5, 75); vlSheet.setColumnWidth(6, 45);
      vlSheet.setColumnWidth(7, 75); vlSheet.setColumnWidth(8, 30); vlSheet.setColumnWidth(9, 30); vlSheet.setColumnWidth(10, 30);
      try { vlSheet.setColumnWidths(11, dateArr.length, 25); } catch(e){}

      try { vlSheet.hideColumns(11 + dateArr.length, 3); } catch(e){}
      try { vlSheet.setFrozenRows(7); vlSheet.setFrozenColumns(2); } catch(e){}
  }
}

// ==============================
// HELPER (PARSERS)
// ==============================
function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  if (sheetName === "Accounts" || sheetName === "Arsip PDF") {
      const headers = values[0];
      const result = [];
      for (let i = 1; i < values.length; i++) {
          let row = values[i]; let obj = {};
          for (let j = 0; j < headers.length; j++) {
              let key = headers[j];
              if (key === "Username") key = "username"; else if (key === "Password") key = "password"; else if (key === "Role") key = "role";
              if (key === "Tanggal") key = "tanggal"; else if (key === "Nama File") key = "filename"; else if (key === "Link Drive") key = "url";
              obj[key] = row[j] !== undefined && row[j] !== null ? row[j].toString() : "";
          }
          if(obj.username || obj.filename) result.push(obj);
      }
      return result;
  }

  if (sheetName === "Jadwal Harian") {
      const headers = values[0]; const result = [];
      for (let i = 1; i < values.length; i++) {
          let row = values[i]; let obj = {};
          for (let j = 0; j < headers.length; j++) {
              let key = headers[j];
              if (key === "Tanggal") key = "tanggal"; else if (key === "Hari") key = "hari"; else if (key === "Main Job") key = "mainJob";
              else if (key === "Sub Job") key = "subJob"; else if (key === "Batas Waktu") key = "batasWaktu"; else if (key === "Keterangan") key = "keterangan";
              else if (key === "Nama Pekerja") key = "namaPekerja"; else if (key === "Jam Mulai") key = "jamMulai"; else if (key === "Jam Selesai") key = "jamSelesai";
              else if (key === "Status") key = "status"; else if (key === "ID") key = "id"; else if (key === "No") continue;

              if (key === "tanggal") {
                  let d = parseDateLocal(row[j]);
                  obj[key] = d ? d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0') + "-" + String(d.getDate()).padStart(2,'0') : row[j];
              } else { obj[key] = row[j]; }
          }
          result.push(obj);
      }
      return result;
  }

  let headerRowIdx = 0;
  if (sheetName === "Jobdesk") {
      headerRowIdx = values.findIndex(r => r[0] === "NO" && r[1] === "TASK/PROYEK");
      if (headerRowIdx === -1) headerRowIdx = 6;
  }

  const headers = values[headerRowIdx];
  if(!headers) return [];

  const result = [];
  const idIdx = headers.indexOf("_ID"); const metaIdx = headers.indexOf("_Meta");

  for (let i = headerRowIdx + 1; i < values.length; i++) {
    const row = values[i];
    if (sheetName === "Jobdesk" && idIdx !== -1 && !row[idIdx]) continue;

    const obj = {};
    if (sheetName === "Jobdesk" && metaIdx !== -1 && row[metaIdx]) {
       try {
           const meta = JSON.parse(row[metaIdx]);
           obj.projectName = meta.projectName;
           obj.category = meta.category || "";
           obj.name = meta.taskName;
           obj.subTask = meta.subTask;
       } catch(e) {}
    }

    for (let j = 0; j < headers.length; j++) {
      let key = headers[j];
      if (!key) continue;
      if (key === "ProjectName" || key === "Project Name") key = "ProjectName";
      if (key === "CategoryName" || key === "Category Name") key = "CategoryName";

      if (sheetName === "Projects" || sheetName === "Categories") {
          obj[key] = row[j];
      }

      if (sheetName === "Reports") {
          if (key === "ID") key = "id"; if (key === "Date") key = "date"; if (key === "Project") key = "project"; if (key === "Workers") key = "workers";
          if (key === "WorkDone") key = "workDone"; if (key === "Issues") key = "issues"; if (key === "Photo 1") key = "photo1";
          if (key === "Photo 2") key = "photo2"; if (key === "Photo 3") key = "photo3"; if (key === "Photo 4") key = "photo4"; if (key === "Photo 5") key = "photo5";
          obj[key] = row[j];
      }

      if (sheetName === "Jobdesk") {
          if (key === "_ID") key = "id"; else if (key === "PIC") key = "pic"; else if (key === "START") key = "startDate";
          else if (key === "END") key = "endDate"; else if (key === "_Status") key = "status";
          else if (key === "%") { key = "progress"; row[j] = parseInt(String(row[j]).replace('%','')) || 0; } else { continue; }

          if((key === "startDate" || key === "endDate")) {
              let parsedD = parseDateLocal(row[j]);
              if (parsedD) obj[key] = parsedD.getFullYear() + "-" + String(parsedD.getMonth()+1).padStart(2,'0') + "-" + String(parsedD.getDate()).padStart(2,'0');
          } else { obj[key] = row[j]; }
      }
    }

    if (sheetName === "Jobdesk") {
       if(obj.progress !== undefined) obj.progress = parseInt(obj.progress) || 0;
       result.push(obj);
    } else { result.push(obj); }
  }
  return result;
}

function updateSheetData(ss, sheetName, headers, rows) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  // pastikan ukuran cukup & merge lama dibongkar agar penulisan tidak gagal
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  if (sheet.getMaxRows() < rows.length + 1) sheet.insertRowsAfter(sheet.getMaxRows(), rows.length + 1 - sheet.getMaxRows());
  try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart(); } catch (x) {}
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (x) {}
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
}

function parseDateLocal(dateInput) {
  if(!dateInput) return null;
  if (dateInput instanceof Date) return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate(), 12, 0, 0, 0);
  let dStr = dateInput.toString().trim().split('T')[0];
  if (dStr.includes('/')) {
    let pts = dStr.split('/');
    if(pts.length !== 3) return null;
    if(pts[2].length === 2) pts[2] = "20" + pts[2];
    return new Date(parseInt(pts[2]), parseInt(pts[1]) - 1, parseInt(pts[0]), 12, 0, 0, 0);
  }
  if (dStr.includes('-')) {
    let parts = dStr.split('-');
    if(parts.length === 3) {
        if (parts[0].length === 4) return new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]), 12, 0, 0, 0);
        else return new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]), 12, 0, 0, 0);
    }
  }
  let d = new Date(dateInput);
  if(!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0, 0);
  return null;
}

function getDaysDiff(d1, d2) {
  if(!d1 || !d2) return 0;
  let start = new Date(d1); start.setHours(0,0,0,0);
  let end = new Date(d2); end.setHours(0,0,0,0);
  let diff = Math.round((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  return diff > 0 ? diff : 0;
}

function isSameDayAppScript(d1, d2) {
  if(!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
}

/* ===================== SHEET TAMPILAN "VIEW ..." ===================== */
function _ukuranMinimal(sh, kolom, baris) {
  var maxK = sh.getMaxColumns();
  if (kolom > maxK) sh.insertColumnsAfter(maxK, kolom - maxK);
  var maxB = sh.getMaxRows();
  if (baris > maxB) sh.insertRowsAfter(maxB, baris - maxB);
}

function _viewSheet(nama, kolom, baris) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nama);
  if (!sh) sh = ss.insertSheet(nama);
  if (sh.isSheetHidden()) sh.showSheet();
  _ukuranMinimal(sh, Math.max(kolom || 12, 12), Math.max(baris || 60, 60));
  // buka semua merge lama: sh.clear() tidak membongkar merge, sehingga
  // merge ulang pada refresh berikutnya bisa gagal.
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (x) {}
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.clear();
  try { sh.clearConditionalFormatRules(); } catch (x) {}
  sh.setHiddenGridlines(true);
  // Tampilan dasar yang konsisten untuk semua sheet VIEW
  try {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())
      .setFontFamily('Inter').setFontSize(10).setFontColor(WARNA.teks)
      .setBackground(WARNA.kertas);
  } catch (x) {}
  return sh;
}

function _judul(sh, teks, lebar) {
  _ukuranMinimal(sh, lebar, 2);
  var rng = sh.getRange(1, 1, 1, lebar);
  try { rng.breakApart(); } catch (x) {}
  try { rng.merge(); } catch (x) {}
  rng.setValue(teks)
    .setFontSize(16).setFontWeight('bold')
    .setBackground(WARNA.header).setFontColor(WARNA.headerText)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(1, 44);
  sh.setRowHeight(2, 10);
}

function _headerTabel(sh, baris, kolomHeader) {
  _ukuranMinimal(sh, kolomHeader.length, baris + 1);
  sh.getRange(baris, 1, 1, kolomHeader.length).setValues([kolomHeader])
    .setFontWeight('bold').setFontSize(10)
    .setBackground(WARNA.sub).setFontColor(WARNA.teks)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, WARNA.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(baris, 28);
  sh.setFrozenRows(baris);
}

function _isiTabel(sh, baris, data, lebarKolom) {
  if (!data.length) {
    _ukuranMinimal(sh, 1, baris);
    sh.getRange(baris, 1).setValue('Belum ada data.').setFontColor(WARNA.abu).setFontStyle('italic');
    return;
  }
  _ukuranMinimal(sh, lebarKolom, baris + data.length);
  // samakan panjang tiap baris agar setValues tidak error
  var rapi = data.map(function (r) {
    var out = r.slice(0, lebarKolom);
    while (out.length < lebarKolom) out.push('');
    return out;
  });
  var rng = sh.getRange(baris, 1, rapi.length, lebarKolom);
  rng.setValues(rapi)
    .setBorder(true, true, true, true, true, true, WARNA.border, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');
  // baris zebra agar tabel mudah dibaca
  rng.setBackgrounds(rapi.map(function (r, i) {
    var w = i % 2 ? WARNA.zebra : WARNA.kertas;
    return r.map(function () { return w; });
  }));
}

// Satu-satunya parser tanggal: hasil selalu dinormalkan ke tengah hari
// agar aman terhadap pergeseran zona waktu.
function _tanggal(v) {
  return parseDateLocal(v);
}

function _fmt(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yy') : '-';
}

function _hariIni() {
  var n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate(), 12, 0, 0, 0);
}

function _statusLabel(t) {
  var p = Number(t.progress) || 0;
  if (p >= 100 || t.status === 'completed') return 'SELESAI';
  if (p > 0 || t.status === 'progress') return 'PROGRES';
  var e = _tanggal(t.endDate);
  if (e && e < _hariIni()) return 'TERLAMBAT';
  return 'PENDING';
}

/** Warna konsisten per PIC (satu-satunya sumber: PIC_COLORS). */
function _warnaPIC(nama) {
  if (!nama) return WARNA.abu;
  var hash = 0;
  for (var i = 0; i < nama.length; i++) hash = nama.charCodeAt(i) + ((hash << 5) - hash);
  return PIC_COLORS[Math.abs(hash) % PIC_COLORS.length];
}

function _warnaiStatus(sh, barisAwal, kolom, jumlah) {
  if (!jumlah) return;
  var rng = sh.getRange(barisAwal, kolom, jumlah, 1);
  var nilai = rng.getValues();
  var bg = nilai.map(function (r) {
    var teks = String(r[0]);
    return [teks === 'SELESAI' ? WARNA.hijau
      : (teks === 'PROGRES' || teks === 'BERJALAN') ? WARNA.kuning
      : (teks === 'TERLAMBAT' || teks === 'PERLU PERHATIAN') ? WARNA.merah : WARNA.abu];
  });
  rng.setBackgrounds(bg).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
}


/* --------- 1. VIEW Dashboard --------- */

/* --------- Navigasi & filter (dipakai VIEW Dashboard) --------- */

var KUNCI_FILTER = 'IAT_FILTER_PROYEK';
var SEMUA_PROYEK = 'SEMUA PROYEK';

/** Ambil filter proyek yang tersimpan (bertahan walau sheet dibangun ulang). */
function _filterProyek() {
  try {
    return PropertiesService.getDocumentProperties().getProperty(KUNCI_FILTER) || SEMUA_PROYEK;
  } catch (x) { return SEMUA_PROYEK; }
}

function _setFilterProyek(nilai) {
  try {
    PropertiesService.getDocumentProperties()
      .setProperty(KUNCI_FILTER, nilai || SEMUA_PROYEK);
  } catch (x) {}
}

/**
 * Baris tombol lompat ke sheet lain (link internal #gid=...), supaya detail
 * cukup dilihat lewat Dashboard tanpa menambah sheet baru.
 */
function _barisNavigasi(sh, baris, tujuan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var urlDasar = ss.getUrl().replace(/#.*$/, '');
  var kol = 1;
  tujuan.forEach(function (t) {
    var target = ss.getSheetByName(t[1]);
    var rng = sh.getRange(baris, kol, 1, 2);
    try { rng.breakApart(); } catch (x) {}
    try { rng.merge(); } catch (x) {}
    if (target) {
      // Pakai rich text link (bukan formula HYPERLINK) supaya tidak #ERROR!
      // pada spreadsheet dengan locale yang memakai ';' sebagai pemisah argumen.
      var label = t[0];
      var rt = SpreadsheetApp.newRichTextValue()
        .setText(label)
        .setLinkUrl(0, label.length, urlDasar + '#gid=' + target.getSheetId())
        .build();
      rng.clearContent();
      rng.setRichTextValue(rt);
      rng.setFontColor(WARNA.headerText).setBackground(WARNA.indigo);
    } else {
      rng.setValue(t[0] + ' (belum ada)')
        .setFontColor(WARNA.abu).setBackground(WARNA.sub);
    }
    rng.setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    kol += 2;
  });
  sh.setRowHeight(baris, 26);
}


/** Dropdown filter proyek di Dashboard. */
function _barisFilter(sh, baris, daftarProyek, terpilih) {
  sh.getRange(baris, 1).setValue('Filter Proyek')
    .setFontWeight('bold').setFontColor(WARNA.teks);
  var opsi = [SEMUA_PROYEK].concat(daftarProyek);
  var sel = sh.getRange(baris, 2, 1, 3);
  try { sel.breakApart(); } catch (x) {}
  try { sel.merge(); } catch (x) {}
  sel.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(opsi, true)
      .setAllowInvalid(false).build()
  );
  sel.setValue(terpilih).setFontWeight('bold')
    .setBackground('#e0e7ff').setFontColor('#1e293b')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(baris, 5, 1, 4).merge()
    .setValue('Pilih proyek untuk memfilter semua tabel di bawah. Tampilan langsung diperbarui.')
    .setFontStyle('italic').setFontColor(WARNA.abu).setVerticalAlignment('middle');
  sh.setRowHeight(baris, 26);
}

function _viewDashboard(d) {
  var sh = _viewSheet('VIEW Dashboard', 12, 24 + d.tasks.length + d.projects.length);
  var filter = _filterProyek();
  var daftarProyek = d.projects.slice();
  d.tasks.forEach(function (t) {
    var p = t.projectName || '';
    if (p && daftarProyek.indexOf(p) === -1) daftarProyek.push(p);
  });
  daftarProyek.sort();
  if (filter !== SEMUA_PROYEK && daftarProyek.indexOf(filter) === -1) filter = SEMUA_PROYEK;

  var cocok = function (nama) { return filter === SEMUA_PROYEK || nama === filter; };
  var tasks = d.tasks.filter(function (t) { return cocok(t.projectName || '(Tanpa Proyek)'); });
  var reports = d.reports.filter(function (r) { return cocok(r.project); });

  _judul(sh, '📊 DASHBOARD — Sistem Laporan & Jobdesk IAT Electrical', 8);

  _barisNavigasi(sh, 2, [
    ['📈 Jobdesk (Gantt)', 'Jobdesk'],
    ['🧾 Visual Laporan', 'Visual Laporan'],
    ['👷 Tugas PIC', 'VIEW Tugas PIC'],
    ['📅 Kalender', 'VIEW Kalender']
  ]);
  _barisFilter(sh, 3, daftarProyek, filter);
  sh.setRowHeight(4, 10);

  var total = tasks.length;
  var selesai = tasks.filter(function (t) { return _statusLabel(t) === 'SELESAI'; }).length;
  var progres = tasks.filter(function (t) { return _statusLabel(t) === 'PROGRES'; }).length;
  var telat = tasks.filter(function (t) { return _statusLabel(t) === 'TERLAMBAT'; }).length;
  var rata = total ? Math.round(tasks.reduce(function (a, t) { return a + (Number(t.progress) || 0); }, 0) / total) : 0;

  var kartu = [
    ['Total Tugas', total, WARNA.indigo],
    ['Selesai', selesai, WARNA.hijau],
    ['Sedang Berjalan', progres, WARNA.kuning],
    ['Terlambat', telat, WARNA.merah],
    ['Rata-rata Progres', rata + '%', WARNA.indigo],
    ['Total Laporan', reports.length, WARNA.abu]
  ];
  kartu.forEach(function (k, i) {
    var kol = 1 + i * 2;
    sh.getRange(5, kol, 1, 2).merge().setValue(k[0])
      .setFontSize(10).setFontColor('#ffffff').setBackground(k[2]).setHorizontalAlignment('center');
    sh.getRange(6, kol, 1, 2).merge().setValue(k[1])
      .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#ffffff').setBorder(true, true, true, true, false, false, k[2], SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.setRowHeight(6, 44);
  });

  // Ringkasan per proyek
  sh.getRange(8, 1).setValue('RINGKASAN PER PROYEK' + (filter === SEMUA_PROYEK ? '' : ' — ' + filter))
    .setFontWeight('bold').setFontSize(12);
  var head = ['Proyek', 'Total Tugas', 'Selesai', 'Berjalan', 'Terlambat', 'Progres (%)', 'Status', 'Jumlah Laporan'];
  _headerTabel(sh, 9, head);

  var proyek = {};
  tasks.forEach(function (t) {
    var p = t.projectName || '(Tanpa Proyek)';
    proyek[p] = proyek[p] || { total: 0, s: 0, g: 0, l: 0, sum: 0 };
    var lab = _statusLabel(t);
    proyek[p].total++;
    proyek[p].sum += Number(t.progress) || 0;
    if (lab === 'SELESAI') proyek[p].s++;
    else if (lab === 'PROGRES') proyek[p].g++;
    else if (lab === 'TERLAMBAT') proyek[p].l++;
  });
  daftarProyek.forEach(function (p) {
    if (cocok(p) && !proyek[p]) proyek[p] = { total: 0, s: 0, g: 0, l: 0, sum: 0 };
  });

  var baris = Object.keys(proyek).map(function (p) {
    var o = proyek[p];
    var pct = o.total ? Math.round(o.sum / o.total) : 0;
    var lap = reports.filter(function (r) { return r.project === p; }).length;
    return [p, o.total, o.s, o.g, o.l, pct, pct >= 100 ? 'SELESAI' : (o.l ? 'PERLU PERHATIAN' : 'BERJALAN'), lap];
  });
  _isiTabel(sh, 10, baris, head.length);
  if (baris.length) {
    sh.getRange(10, 6, baris.length, 1).setNumberFormat('0"%"').setHorizontalAlignment('center');
    sh.getRange(10, 2, baris.length, 4).setHorizontalAlignment('center');
    _warnaiStatus(sh, 10, 7, baris.length);
  }
  sh.setColumnWidth(1, 220);
  sh.autoResizeColumns(2, 7);
  sh.setFrozenRows(3);

  // Komparasi Aktual vs Plan digabung di sini (dulu sheet "VIEW Komparasi").
  var akhir = _bagianKomparasi(sh, { tasks: tasks }, 10 + baris.length + 2);
  _barisNavigasi(sh, akhir + 2, [
    ['📈 Jobdesk (Gantt)', 'Jobdesk'],
    ['🧾 Visual Laporan', 'Visual Laporan'],
    ['👷 Tugas PIC', 'VIEW Tugas PIC'],
    ['📅 Kalender', 'VIEW Kalender']
  ]);
}


/* --------- 2. VIEW Kalender --------- */

function _viewKalender(d) {
  var sh = _viewSheet('VIEW Kalender', 10, 60 + d.jadwal.length);
  var kini = _hariIni();
  _judul(sh, '📅 KALENDER JOBDESK — ' + BULAN[kini.getMonth()] + ' ' + kini.getFullYear(), 7);
  _barisNavigasi(sh, 2, [
    ['📊 Dashboard', 'VIEW Dashboard'],
    ['📈 Jobdesk (Gantt)', 'Jobdesk'],
    ['👷 Tugas PIC', 'VIEW Tugas PIC']
  ]);

  _headerTabel(sh, 3, HARI);
  sh.getRange(3, 1, 1, 7).setHorizontalAlignment('center');

  var awal = new Date(kini.getFullYear(), kini.getMonth(), 1);
  var jmlHari = new Date(kini.getFullYear(), kini.getMonth() + 1, 0).getDate();
  var offset = awal.getDay();
  var grid = [];
  var sel = [];
  for (var i = 0; i < offset; i++) sel.push('');
  for (var tgl = 1; tgl <= jmlHari; tgl++) {
    var hari = new Date(kini.getFullYear(), kini.getMonth(), tgl);
    var aktif = d.tasks.filter(function (t) {
      var s = _tanggal(t.startDate), e = _tanggal(t.endDate);
      return s && e && hari >= s && hari <= e;
    });
    var isi = tgl + (aktif.length ? '\n• ' + aktif.slice(0, 4).map(function (t) {
      return (t.name || '') + (t.pic ? ' (' + t.pic + ')' : '');
    }).join('\n• ') : '');
    if (aktif.length > 4) isi += '\n+' + (aktif.length - 4) + ' lainnya';
    sel.push(isi);
    if (sel.length === 7) { grid.push(sel); sel = []; }
  }
  if (sel.length) { while (sel.length < 7) sel.push(''); grid.push(sel); }

  if (grid.length) {
    var rng = sh.getRange(4, 1, grid.length, 7);
    rng.setValues(grid).setWrap(true).setVerticalAlignment('top').setFontSize(9)
      .setBorder(true, true, true, true, true, true, WARNA.border, SpreadsheetApp.BorderStyle.SOLID);
    var bg = grid.map(function (row) {
      return row.map(function (v, c) {
        if (!v) return '#ffffff';
        var no = parseInt(String(v).split('\n')[0], 10);
        if (no === kini.getDate()) return '#e0e7ff';
        return (c === 0 || c === 6) ? WARNA.weekend : '#ffffff';
      });
    });
    rng.setBackgrounds(bg);
    for (var r = 0; r < grid.length; r++) sh.setRowHeight(4 + r, 80);
  }

  for (var c2 = 1; c2 <= 7; c2++) sh.setColumnWidth(c2, 170);

  // Detail jadwal per tanggal tidak diulang di sini karena sudah ada
  // pada sheet data "Jadwal Harian".
  var catatan = 4 + grid.length + 1;
  sh.getRange(catatan, 1, 1, 7).merge()
    .setValue('Detail jam & keterangan per tanggal ada di sheet "Jadwal Harian".')
    .setFontStyle('italic').setFontColor(WARNA.abu);
}

/* --------- Bagian KOMPARASI (di dalam VIEW Dashboard) --------- */

function _bagianKomparasi(sh, d, barisMulai) {
  sh.getRange(barisMulai, 1).setValue('KOMPARASI AKTUAL vs PLAN')
    .setFontWeight('bold').setFontSize(12);
  var head = ['Proyek', 'Tugas', 'PIC', 'Rencana Selesai', 'Plan (%)', 'Aktual (%)', 'Deviasi (%)', 'Keterangan'];
  var barisHead = barisMulai + 1;
  _ukuranMinimal(sh, head.length, barisHead + 1);
  sh.getRange(barisHead, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground(WARNA.sub).setFontColor(WARNA.teks)
    .setBorder(true, true, true, true, true, true, WARNA.border, SpreadsheetApp.BorderStyle.SOLID);

  var kini = _hariIni();
  var baris = (d.tasks || []).map(function (t) {
    var st = _tanggal(t.startDate), e = _tanggal(t.endDate);
    var plan = 0;
    if (st && e) {
      var totalMs = e.getTime() - st.getTime();
      plan = kini <= st ? 0 : (kini >= e ? 100 : Math.round(((kini - st) / (totalMs || 1)) * 100));
    }
    var aktual = Number(t.progress) || 0;
    var dev = aktual - plan;
    return [
      t.projectName || '-',
      (t.name || '') + (t.subTask ? ' — ' + t.subTask : ''),
      t.pic || '-', _fmt(e), plan, aktual, dev,
      dev >= 0 ? 'SESUAI / LEBIH CEPAT' : (dev < -20 ? 'TERTINGGAL JAUH' : 'SEDIKIT TERTINGGAL')
    ];
  });
  var isi = barisHead + 1;
  _isiTabel(sh, isi, baris, head.length);
  if (baris.length) {
    sh.getRange(isi, 5, baris.length, 3).setHorizontalAlignment('center');
    sh.getRange(isi, 7, baris.length, 2)
      .setBackgrounds(baris.map(function (r) {
        var dev = r[6];
        var w = dev >= 0 ? WARNA.hijau : (dev < -20 ? WARNA.merah : WARNA.kuning);
        return [w, w];
      }))
      .setFontColor('#ffffff').setFontWeight('bold');
  }
  return isi + baris.length;
}

/* --------- 3. VIEW Tugas PIC --------- */

function _viewTugasPIC(d) {
  var sh = _viewSheet('VIEW Tugas PIC', 9, 22 + d.tasks.length * 2);
  var filter = _filterProyek();
  var tasks = filter === SEMUA_PROYEK ? d.tasks : d.tasks.filter(function (t) {
    return (t.projectName || '(Tanpa Proyek)') === filter;
  });
  _judul(sh, '👷 TUGAS PER PIC / PEKERJA — ' +
    (filter === SEMUA_PROYEK ? 'semua proyek' : filter), 9);
  _barisNavigasi(sh, 2, [
    ['📊 Dashboard', 'VIEW Dashboard'],
    ['📈 Jobdesk (Gantt)', 'Jobdesk'],
    ['🧾 Visual Laporan', 'Visual Laporan'],
    ['📅 Kalender', 'VIEW Kalender']
  ]);

  var perPic = {};
  tasks.forEach(function (t) {
    var pic = t.pic || '(Belum ditugaskan)';
    (perPic[pic] = perPic[pic] || []).push(t);
  });

  var head = ['PIC', 'Proyek', 'Kategori', 'Tugas', 'Sub Tugas', 'Mulai', 'Selesai', 'Progres (%)', 'Status'];
  _headerTabel(sh, 3, head);

  var baris = [];
  Object.keys(perPic).sort().forEach(function (pic) {
    perPic[pic].forEach(function (t) {
      baris.push([pic, t.projectName || '-', t.category || '-',
        t.name || '-', t.subTask || '-',
        _fmt(_tanggal(t.startDate)), _fmt(_tanggal(t.endDate)),
        Number(t.progress) || 0, _statusLabel(t)]);
    });
  });
  _isiTabel(sh, 4, baris, head.length);
  if (baris.length) {
    sh.getRange(4, 8, baris.length, 1).setHorizontalAlignment('center');
    _warnaiStatus(sh, 4, 9, baris.length);
  }

  // Ringkasan beban kerja
  var mulai = 4 + baris.length + 2;
  sh.getRange(mulai - 1, 1).setValue('RINGKASAN BEBAN KERJA').setFontWeight('bold').setFontSize(12);
  var head2 = ['PIC', 'Total Tugas', 'Selesai', 'Sisa', 'Rata-rata Progres (%)'];
  _ukuranMinimal(sh, head2.length, mulai + 1);
  sh.getRange(mulai, 1, 1, head2.length).setValues([head2]).setFontWeight('bold')
    .setBackground(WARNA.sub).setFontColor('#334155');
  var ring = Object.keys(perPic).sort().map(function (pic) {
    var arr = perPic[pic];
    var s = arr.filter(function (t) { return _statusLabel(t) === 'SELESAI'; }).length;
    var avg = Math.round(arr.reduce(function (a, t) { return a + (Number(t.progress) || 0); }, 0) / arr.length);
    return [pic, arr.length, s, arr.length - s, avg];
  });
  _isiTabel(sh, mulai + 1, ring, head2.length);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 160); sh.setColumnWidth(4, 240);
}

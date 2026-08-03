/**
 * BACKEND UNIVERSAL DATA SERVICE & ROUTER
 * Aplikasi Poin Pelanggaran & Prestasi MTs Darun Najah Gading
 *
 * Cocok untuk frontend GitHub Pages:
 * - GET  ?action=get_initial_data
 * - POST form-encoded untuk save_transaction dan delete_transaction
 */

function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : "";

    switch (action) {
      case "get_initial_data":
        return responseSuccess(getInitialData());
      default:
        return responseError("Aksi tidak dikenali");
    }
  } catch (err) {
    return responseError(err.toString());
  }
}

function doPost(e) {
  try {
    var payload = parseRequestPayload(e);
    var action = payload.action;

    switch (action) {
      case "save_transaction":
        return responseSuccess(saveTransaction(payload));
      case "delete_transaction":
        return responseSuccess(deleteTransaction(payload));
      default:
        return responseError("Aksi tidak dikenali");
    }
  } catch (err) {
    return responseError(err.toString());
  }
}

function parseRequestPayload(e) {
  if (e && e.parameter) {
    var normalized = {};
    var keys = Object.keys(e.parameter);

    keys.forEach(function(key) {
      var value = e.parameter[key];
      if (Array.isArray(value)) {
        normalized[key] = value.length > 1 ? value : value[0];
      } else {
        normalized[key] = value;
      }
    });

    if (keys.length > 0) {
      return normalized;
    }
  }

  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      var contents = e.postData.contents;
      if (typeof contents === 'string' && contents.indexOf('=') !== -1) {
        var formParams = {};
        contents.split('&').forEach(function(part) {
          if (!part) return;
          var pair = part.split('=');
          var key = decodeURIComponent(pair[0]);
          var value = pair.length > 1 ? decodeURIComponent(pair.slice(1).join('=')) : '';
          formParams[key] = value;
        });
        return formParams;
      }
      return {};
    }
  }

  return {};
}

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty("SPREADSHEET_ID");

  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  throw new Error('SPREADSHEET_ID belum diisi di Script Properties, dan script ini tidak terhubung ke spreadsheet aktif.');
}

function requireSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" tidak ditemukan.');
  }
  return sheet;
}

// --- SETUP DATABASE AWAL & INTEGRASI DATA MASTER LENGKAP ---
function setupDatabase() {
  var ss = getSpreadsheet_();

  // 1. Sheet Guru (15 Guru Resmi)
  var sheetGuru = ss.getSheetByName("Guru") || ss.insertSheet("Guru");
  sheetGuru.clearContents();
  sheetGuru.appendRow(["id_guru", "nama_guru"]);

  var sampleGuru = [
    [1, "Zainuri, S.Ag. M.PdI"],
    [2, "Indah Estu W, S.Pd"],
    [3, "Masnukhin, S.PdI"],
    [4, "Anik Mariatul K, S.Pd"],
    [5, "Binti Mu'minatin, S.Pd"],
    [6, "Lukluil Maknun, S.Pd.I"],
    [7, "Agit Firdaus, S.Si"],
    [8, "Munikah, S.Pd"],
    [9, "Ana Maftuha, S.PdI"],
    [10, "Duwi Rahayu, S.Pd"],
    [11, "Ahmad Khusairi, S.Pd"],
    [12, "Nur Tahjud Yaumil, S.Pd"],
    [13, "Farid Surya Efendi"],
    [14, "M. Wahyu Eka S, S.IP"],
    [15, "Hasan"]
  ];
  sheetGuru.getRange(2, 1, sampleGuru.length, 2).setValues(sampleGuru);

  // 2. Sheet Siswa
  var sheetSiswa = ss.getSheetByName("Siswa") || ss.insertSheet("Siswa");
  sheetSiswa.clearContents();
  sheetSiswa.appendRow(["id_siswa", "nisn", "nama_siswa", "kelas"]);

  var sampleSiswa = [
    [1, "007001", "Aina Mufidah Kusuma Putri", "VII"],
    [2, "007002", "Amelia Azzahra", "VII"],
    [3, "007003", "Andyn Nur Rosita", "VII"],
    [4, "007004", "Aqil Naufal Zahir", "VII"],
    [5, "007005", "Asva Januar Pradana", "VII"],
    [6, "007006", "Ayla Ayu Puspita", "VII"],
    [7, "007007", "Azzahra Meliana", "VII"],
    [8, "007008", "Dara Putri Sandika", "VII"],
    [9, "007009", "Dimas Rahmada Setiawan", "VII"],
    [10, "007010", "Inayatul Fadliyah", "VII"],
    [11, "007011", "Lilik Soqifatul", "VII"],
    [12, "007012", "Mikayla Febrina Azzahra", "VII"],
    [13, "007013", "Muhammad Abizar Adillah", "VII"],
    [14, "007014", "Muhammad Alvino Puji Setiawan", "VII"],
    [15, "007015", "Muhammad Reza Pratama", "VII"],
    [16, "007016", "Mutiara Zaskia Putri", "VII"],
    [17, "007017", "Nagita Rafa Nuraini", "VII"],
    [18, "007018", "Nazifa Naura Iskandar", "VII"],
    [19, "007019", "Nur Lailatul Aprilia", "VII"],
    [20, "007020", "Salisa Putri Aulia", "VII"],
    [21, "007021", "Siti Nor Kaamila", "VII"],
    [22, "007022", "Tegar Adam Firmansyah", "VII"],
    [23, "007023", "Gus Solahudin Al Faruq", "VII"],
    [24, "008001", "Mochammad Hendrik Irwansyah", "VIII"],
    [25, "008002", "Ahmad Rendy Adi Pratama", "VIII"],
    [26, "008003", "Ayu Diva Nur Khoriyah", "VIII"],
    [27, "008004", "Bagus Pratama", "VIII"],
    [28, "008005", "Eka Risma Andini", "VIII"],
    [29, "008006", "Iza Fakhorotun Nisa'", "VIII"],
    [30, "008007", "Kevin Anggara Saputra", "VIII"],
    [31, "008008", "M. Awan Prasetyo", "VIII"],
    [32, "008009", "Mar'atus solikha", "VIII"],
    [33, "008010", "Moh. Dimas Rifaldi", "VIII"],
    [34, "008011", "Mohammad Faris Adi Saputra", "VIII"],
    [35, "008012", "Muhamad Syarif Fadhil", "VIII"],
    [36, "008013", "Muhammad Azam Akbar", "VIII"],
    [37, "008014", "Muhammad Farel Fikri Wahyudin", "VIII"],
    [38, "008015", "Mukhammad Aat Ainur S", "VIII"],
    [39, "008016", "Natasya Zahra Nazaya", "VIII"],
    [40, "008017", "Nikeyla Azzahra", "VIII"],
    [41, "008018", "Raka Aditya", "VIII"],
    [42, "008019", "Reynata Jhehan Adistin", "VIII"],
    [43, "008020", "Rizki Ariya Bayu Siswanto", "VIII"],
    [44, "008021", "Umi Anggraini", "VIII"],
    [45, "008022", "Zahira Regina Putri", "VIII"],
    [46, "008023", "Zahra Puspita Sari", "VIII"],
    [47, "008024", "Dimas Rico Sanjaya", "VIII"],
    [48, "009001", "Adel Amelia Tatuil", "IX"],
    [49, "009002", "Ahmad Kelvin Fairus", "IX"],
    [50, "009003", "Ahmad Winas Yusril", "IX"],
    [51, "009004", "Akhmad Fajar Romadhoni", "IX"],
    [52, "009005", "Aliyu Nisya Mustofa Nggoni", "IX"],
    [53, "009006", "Amabel Damara Dwi F.", "IX"],
    [54, "009007", "Eka Putri Rahayu", "IX"],
    [55, "009008", "Fitri Dwi Rahmadania", "IX"],
    [56, "009009", "Hesti Dwi kurniawati", "IX"],
    [57, "009010", "M Aditya Irmawan", "IX"],
    [58, "009011", "Azzam al Fajri", "IX"],
    [59, "009012", "Moh Fajar Prasetyo", "IX"],
    [60, "009013", "Mohammad Farel Dwi Cahyo", "IX"],
    [61, "009014", "Muhammad Annafi Hidayatulloh", "IX"],
    [62, "009015", "Muhammad Arbi Hamdi", "IX"],
    [63, "009016", "Muhammad Hamdan Hidayatulloh", "IX"],
    [64, "009017", "Nur Afrida", "IX"],
    [65, "009018", "Rahmat Hadi Prayitno", "IX"],
    [66, "009019", "Salfa Anggraini", "IX"],
    [67, "009020", "Septya Vika Feriska", "IX"],
    [68, "009021", "Zulia Zahra Syafiyah", "IX"],
    [69, "009022", "ROSDA MARCH MUNI MAYA MAPIASE", "IX"],
    [70, "009023", "Muhammad Revi Aditya", "IX"],
    [71, "009024", "Ahmad Maulana", "IX"]
  ];
  sheetSiswa.getRange(2, 1, sampleSiswa.length, 4).setValues(sampleSiswa);

  // 3. Sheet Master Poin
  var sheetPoin = ss.getSheetByName("Master_Poin") || ss.insertSheet("Master_Poin");
  sheetPoin.clearContents();
  sheetPoin.appendRow(["id_poin", "kategori", "nama_kejadian", "bobot_poin"]);

  var samplePoin = [
    ["P-A01", "Pelanggaran", "[Ringan] Terlambat datang ke madrasah tanpa alasan yang diterima", 5],
    ["P-A02", "Pelanggaran", "[Ringan] Tidak memakai seragam sesuai ketentuan", 5],
    ["P-A03", "Pelanggaran", "[Ringan] Tidak memakai atribut lengkap", 5],
    ["P-A04", "Pelanggaran", "[Ringan] Memakai sepatu yang tidak sesuai ketentuan", 5],
    ["P-A05", "Pelanggaran", "[Ringan] Tidak membawa perlengkapan belajar", 5],
    ["P-A06", "Pelanggaran", "[Ringan] Tidak melaksanakan piket kelas", 5],
    ["P-A07", "Pelanggaran", "[Ringan] Membuang sampah sembarangan", 5],
    ["P-A08", "Pelanggaran", "[Ringan] Makan atau minum saat pembelajaran tanpa izin", 5],
    ["P-A09", "Pelanggaran", "[Ringan] Keluar kelas saat pergantian jam tanpa keperluan jelas", 5],
    ["P-A10", "Pelanggaran", "[Ringan] Membuat kegaduhan atau mengganggu ketertiban kelas", 5],
    ["P-A11", "Pelanggaran", "[Ringan] Tidak membawa perlengkapan salat sesuai ketentuan", 5],
    ["P-A12", "Pelanggaran", "[Ringan] Berpenampilan tidak rapi sesuai aturan madrasah", 5],
    ["P-B01", "Pelanggaran", "[Sedang] Tidak mengikuti pembelajaran tanpa izin", 10],
    ["P-B02", "Pelanggaran", "[Sedang] Tidak mengikuti salat berjamaah atau kegiatan keagamaan wajib tanpa alasan", 10],
    ["P-B03", "Pelanggaran", "[Sedang] Berada di kantin atau tempat lain saat jam pelajaran tanpa izin", 10],
    ["P-B04", "Pelanggaran", "[Sedang] Menggunakan kata-kata kasar atau tidak pantas", 10],
    ["P-B05", "Pelanggaran", "[Sedang] Mengganggu teman secara berulang", 10],
    ["P-B06", "Pelanggaran", "[Sedang] Tidak mengerjakan tugas secara sengaja dan berulang", 10],
    ["P-B07", "Pelanggaran", "[Sedang] Membawa atau menggunakan HP tanpa izin madrasah", 15],
    ["P-B08", "Pelanggaran", "[Sedang] Menggunakan HP pada waktu yang tidak diperbolehkan", 15],
    ["P-B09", "Pelanggaran", "[Sedang] Membolos satu atau beberapa jam pelajaran", 15],
    ["P-B10", "Pelanggaran", "[Sedang] Keluar dari lingkungan madrasah tanpa izin", 15],
    ["P-B11", "Pelanggaran", "[Sedang] Rambut atau penampilan tidak sesuai ketentuan setelah peringatan", 15],
    ["P-B12", "Pelanggaran", "[Sedang] Merusak fasilitas madrasah karena kelalaian", 15],
    ["P-B13", "Pelanggaran", "[Sedang] Berbohong kepada guru atau tenaga kependidikan", 15],
    ["P-B14", "Pelanggaran", "[Sedang] Tidak masuk madrasah tanpa keterangan", 20],
    ["P-B15", "Pelanggaran", "[Sedang] Menyontek atau membantu teman menyontek", 20],
    ["P-B16", "Pelanggaran", "[Sedang] Bersikap tidak sopan terhadap guru, tenaga kependidikan, atau warga madrasah", 20],
    ["P-C01", "Pelanggaran", "[Berat] Membolos satu hari penuh tanpa keterangan", 25],
    ["P-C02", "Pelanggaran", "[Berat] Memalsukan tanda tangan surat izin atau dokumen madrasah", 25],
    ["P-C03", "Pelanggaran", "[Berat] Mengambil atau menggunakan barang milik orang lain tanpa izin", 25],
    ["P-C04", "Pelanggaran", "[Berat] Melakukan perundungan verbal atau sosial", 30],
    ["P-C05", "Pelanggaran", "[Berat] Mengunggah konten yang menghina atau mencemarkan nama baik", 30],
    ["P-C06", "Pelanggaran", "[Berat] Merusak fasilitas madrasah dengan sengaja", 30],
    ["P-C07", "Pelanggaran", "[Berat] Berkelahi atau melakukan kekerasan fisik", 40],
    ["P-C08", "Pelanggaran", "[Berat] Melakukan perundungan fisik atau ancaman kekerasan", 40],
    ["P-C09", "Pelanggaran", "[Berat] Melakukan pelecehan atau tindakan tidak senonoh", 40],
    ["P-C10", "Pelanggaran", "[Berat] Mencuri barang atau uang", 40],
    ["P-C11", "Pelanggaran", "[Berat] Membawa, menyimpan, atau menyebarkan konten pornografi", 40],
    ["P-C12", "Pelanggaran", "[Berat] Membawa atau menggunakan rokok/vape di lingkungan madrasah", 50],
    ["P-C13", "Pelanggaran", "[Berat] Melakukan tindakan yang mencemarkan nama baik madrasah", 50],
    ["P-C14", "Pelanggaran", "[Berat] Membawa benda berbahaya yang dapat melukai orang lain", 50],
    ["P-C15", "Pelanggaran", "[Berat] Membawa, menggunakan, atau mengedarkan minuman keras/narkotika", 50],
    ["P-C16", "Pelanggaran", "[Berat] Melakukan kekerasan berat yang menyebabkan cedera", 50],
    ["R-01", "Prestasi", "[Pengurangan] Tidak melakukan pelanggaran selama 1 bulan berturut-turut", 5],
    ["R-02", "Prestasi", "[Pengurangan] Menunjukkan perubahan sikap dan kedisiplinan secara konsisten selama 1 bulan", 5],
    ["R-03", "Prestasi", "[Pengurangan] Melaksanakan tugas pembinaan yang diberikan kesiswaan dengan baik", 5],
    ["R-04", "Prestasi", "[Pengurangan] Aktif menjaga kebersihan dan ketertiban lingkungan madrasah secara sukarela", 5],
    ["R-05", "Prestasi", "[Pengurangan] Membantu kegiatan sosial atau kepedulian di lingkungan madrasah", 5],
    ["R-06", "Prestasi", "[Pengurangan] Menjadi petugas kegiatan madrasah dan melaksanakan tugas dengan baik", 5],
    ["R-07", "Prestasi", "[Pengurangan] Menunjukkan kejujuran atau tanggung jawab yang patut diapresiasi", 5],
    ["R-08", "Prestasi", "[Pengurangan] Tidak mengulangi jenis pelanggaran yang sama selama 2 bulan", 10],
    ["R-09", "Prestasi", "[Pengurangan] Menunjukkan perubahan perilaku yang signifikan berdasarkan evaluasi wali kelas/kesiswaan", 10],
    ["R-10", "Prestasi", "[Pengurangan] Aktif dan konsisten dalam kegiatan keagamaan atau pembiasaan positif madrasah", 10],
    ["R-11", "Prestasi", "[Pengurangan] Aktif dalam kegiatan OSIS, ekstrakurikuler, atau kegiatan resmi madrasah dengan disiplin", 10],
    ["R-12", "Prestasi", "[Pengurangan] Menjadi teladan dalam kedisiplinan dan ketertiban selama satu semester", 15],
    ["R-13", "Prestasi", "[Pengurangan/Prestasi] Meraih prestasi tingkat madrasah", 5],
    ["R-14", "Prestasi", "[Pengurangan/Prestasi] Meraih prestasi tingkat kecamatan", 10],
    ["R-15", "Prestasi", "[Pengurangan/Prestasi] Meraih prestasi tingkat kabupaten/kota", 15],
    ["R-16", "Prestasi", "[Pengurangan/Prestasi] Meraih prestasi tingkat provinsi", 20],
    ["R-17", "Prestasi", "[Pengurangan/Prestasi] Meraih prestasi tingkat nasional atau lebih tinggi", 25]
  ];
  sheetPoin.getRange(2, 1, samplePoin.length, 4).setValues(samplePoin);

  // 4. Sheet Transaksi Poin
  var sheetTrx = ss.getSheetByName("Transaksi_Poin") || ss.insertSheet("Transaksi_Poin");
  if (sheetTrx.getLastRow() === 0) {
    sheetTrx.appendRow(["id_transaksi", "tanggal", "id_siswa", "nama_siswa", "kelas", "kategori", "nama_kejadian", "bobot_poin", "keterangan", "nama_guru_input", "created_at"]);
  }

  // 5. Sheet Logs
  var sheetLogs = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
  if (sheetLogs.getLastRow() === 0) {
    sheetLogs.appendRow(["timestamp", "nama_guru", "aksi", "detail"]);
  }

  return "Database Berhasil Diperbarui! 17 Butir Pengurangan Poin / Prestasi Telah Terpasang.";
}

// --- SERVICE FUNCTIONS ---
function getInitialData() {
  var ss = getSpreadsheet_();

  var dataGuru = sheetToObjects(requireSheet_(ss, "Guru"));
  var dataSiswa = sheetToObjects(requireSheet_(ss, "Siswa")).map(function(s) {
    return {
      id_siswa: s.id_siswa,
      nisn: String(s.nisn || ""),
      nama_siswa: String(s.nama_siswa || ""),
      kelas: String(s.kelas || "")
    };
  });

  var rawPoin = sheetToObjects(requireSheet_(ss, "Master_Poin"));
  var masterPoin = { Pelanggaran: [], Prestasi: [] };
  rawPoin.forEach(function(p) {
    if (!masterPoin[p.kategori]) masterPoin[p.kategori] = [];
    masterPoin[p.kategori].push(p);
  });

  var rawTrx = sheetToObjects(requireSheet_(ss, "Transaksi_Poin"));
  var transaksi = rawTrx.map(function(t) {
    return {
      id: t.id_transaksi,
      tanggal: String(t.tanggal).slice(0, 10),
      id_siswa: parseInt(t.id_siswa, 10),
      nama_siswa: t.nama_siswa,
      kelas: String(t.kelas),
      kategori: t.kategori,
      kejadian: t.nama_kejadian,
      bobot: parseInt(t.bobot_poin, 10),
      keterangan: t.keterangan,
      guru: t.nama_guru_input
    };
  }).reverse();

  return {
    guru: dataGuru,
    siswa: dataSiswa,
    masterPoin: masterPoin,
    transaksi: transaksi
  };
}

function saveTransaction(payload) {
  var ss = getSpreadsheet_();
  var sheetTrx = requireSheet_(ss, "Transaksi_Poin");

  var idTrx = "TRX-" + new Date().getTime();
  var now = new Date();

  sheetTrx.appendRow([
    idTrx,
    payload.tanggal,
    payload.id_siswa,
    payload.nama_siswa,
    payload.kelas,
    payload.kategori,
    payload.kejadian,
    payload.bobot,
    payload.keterangan,
    payload.guru,
    now
  ]);

  writeLog(payload.guru, "Input Poin", "Input " + payload.kategori + " (" + payload.bobot + ") untuk " + payload.nama_siswa);

  return {
    id: idTrx,
    tanggal: payload.tanggal,
    id_siswa: payload.id_siswa,
    nama_siswa: payload.nama_siswa,
    kelas: payload.kelas,
    kategori: payload.kategori,
    kejadian: payload.kejadian,
    bobot: payload.bobot,
    keterangan: payload.keterangan,
    guru: payload.guru
  };
}

function deleteTransaction(payload) {
  var ss = getSpreadsheet_();
  var sheetTrx = requireSheet_(ss, "Transaksi_Poin");
  var data = sheetTrx.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === payload.id_transaksi) {
      sheetTrx.deleteRow(i + 1);
      writeLog(payload.guru, "Hapus Poin", "Menghapus transaksi ID: " + payload.id_transaksi);
      return true;
    }
  }

  throw new Error("ID Transaksi tidak ditemukan.");
}

// --- HELPER FUNCTIONS ---
function writeLog(guru, aksi, detail) {
  var ss = getSpreadsheet_();
  var sheetLogs = requireSheet_(ss, "Logs");
  sheetLogs.appendRow([new Date(), guru, aksi, detail]);
}

function sheetToObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  var headers = values[0];
  var results = [];

  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
    }
    results.push(obj);
  }
  return results;
}

function responseSuccess(data) {
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    data: data
  })).setMimeType(ContentService.MimeType.JSON);
}

function responseError(message) {
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: message
  })).setMimeType(ContentService.MimeType.JSON);
}

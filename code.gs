function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu Raport')
    .addItem('Update Data & Pelajaran', 'updateNamaMapel')
    .addItem('Update Nilai & Deskripsi', 'updateNilaiDanDeskripsi')
    .addItem('Update Catatan Guru', 'updateCatatan')
    .addToUi();
}

function updateNamaMapel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMapel = ss.getSheetByName("Mapel");
  var sheetNilai = ss.getSheetByName("Nilai");
  var sheetRaport = ss.getSheetByName("Raport");

  if (!sheetMapel || !sheetNilai || !sheetRaport) {
    SpreadsheetApp.getUi().alert("Sheet 'Mapel', 'Nilai', atau 'Raport' tidak ditemukan!");
    return;
  }

  var targetMapelCols = ["D", "I", "N", "S", "X", "AC", "AH", "AM", "AR", "AW", "BB", "BG", "BL", "BQ", "BV"];
  var targetEkstraCols = ["CB", "CD", "CF", "CH", "CJ"];
  var mapelData = sheetMapel.getRange("B2:B16").getValues();
  var ekstraData = sheetMapel.getRange("B18:B22").getValues();

  sheetNilai.showColumns(1, sheetNilai.getMaxColumns());
  for (var i = 0; i < mapelData.length; i++) {
    var colIndex = sheetNilai.getRange(targetMapelCols[i] + "1").getColumn();
    if (mapelData[i][0]) {
      sheetNilai.getRange(1, colIndex).setValue(mapelData[i][0]);
    } else {
      sheetNilai.hideColumns(colIndex, 5); // Sembunyikan jika kosong
    }
  }
  for (var j = 0; j < ekstraData.length; j++) {
    var colIndexEkstra = sheetNilai.getRange(targetEkstraCols[j] + "1").getColumn();
    if (ekstraData[j][0]) {
      sheetNilai.getRange(1, colIndexEkstra).setValue(ekstraData[j][0]);
    } else {
      sheetNilai.hideColumns(colIndexEkstra, 2);
    }
  }
  var dataNama = sheetNilai.getRange("A2:C" + sheetNilai.getLastRow()).getValues();
  sheetRaport.getRange("A2:C" + (dataNama.length + 1)).setValues(dataNama);
  SpreadsheetApp.getUi().alert("Nama Mapel, Ekstrakurikuler, dan data siswa berhasil diperbarui!");
}

function updateNilaiDanDeskripsi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNilai = ss.getSheetByName("Nilai");
  const sheetDeskripsi = ss.getSheetByName("Deskripsi");
  const sheetRaport = ss.getSheetByName("Raport");

  const jumlahMapel = 15;
  const lingkupPerMapel = 5;
  const kolomAwalNilai = 4; // kolom D
  const kolomDeskripsiMulai = 4; // kolom D di sheet Deskripsi

  const dataNilai = sheetNilai.getDataRange().getValues().slice(1); // tanpa header
  const deskripsiList = sheetDeskripsi.getRange(2, kolomDeskripsiMulai, 75, 1).getValues().flat();

  // Kolom tempat nilai akhir & deskripsi ditulis di Raport: [D, G, J, M, ..., AT]
  const targetKolomAwal = ["D", "G", "J", "M", "P", "S", "V", "Y", "AB", "AE", "AH", "AK", "AN", "AQ", "AT"];

  const hasilNilai = [];
  const hasilDeskripsiTinggi = [];
  const hasilDeskripsiRendah = [];

  dataNilai.forEach((row) => {
    const barisNilai = [];

    for (let i = 0; i < jumlahMapel; i++) {
      const startIndex = kolomAwalNilai - 1 + (i * lingkupPerMapel);
      const nilaiMapel = row.slice(startIndex, startIndex + lingkupPerMapel);
      const nilaiValid = nilaiMapel.filter(n => !isNaN(n));
      const rata2 = nilaiValid.length ? Math.round(nilaiValid.reduce((a, b) => a + b) / nilaiValid.length) : "";

      barisNilai.push(rata2);
    }

    hasilNilai.push(barisNilai);
  });

  for (let i = 0; i < jumlahMapel; i++) {
    const kolomIndex = kolomAwalNilai - 1 + (i * lingkupPerMapel);
    const deskripsiMapel = deskripsiList.slice(i * lingkupPerMapel, (i + 1) * lingkupPerMapel);
    const tinggi = [];
    const rendah = [];

    dataNilai.forEach((row) => {
      const nilaiMapel = row.slice(kolomIndex, kolomIndex + lingkupPerMapel);
      const nilaiValid = nilaiMapel.map(n => Number(n)).filter(n => !isNaN(n));

      const nilaiMax = Math.max(...nilaiValid);
      const nilaiMin = Math.min(...nilaiValid);

      const deskripsiTinggi = nilaiMapel
        .map((n, idx) => n === nilaiMax ? deskripsiMapel[idx] : null)
        .filter(Boolean)
        .map(d => d.charAt(0).toLowerCase() + d.slice(1));
      
      const deskripsiRendah = nilaiMapel
        .map((n, idx) => n === nilaiMin ? deskripsiMapel[idx] : null)
        .filter(Boolean)
        .map(d => d.charAt(0).toLowerCase() + d.slice(1));

      tinggi.push([`Capaian baik pada ${[...new Set(deskripsiTinggi)].join(", ")}`]);
      rendah.push([`Perlu penguatan pada ${[...new Set(deskripsiRendah)].join(", ")}`]);
    });

    hasilDeskripsiTinggi.push(tinggi);
    hasilDeskripsiRendah.push(rendah);
  }

  // Menuliskan ke Sheet Raport
  hasilNilai.forEach((barisNilai, rowIdx) => {
    barisNilai.forEach((nilai, mapelIdx) => {
      const col = targetKolomAwal[mapelIdx];
      const colIndex = sheetRaport.getRange(col + "1").getColumn();

      sheetRaport.getRange(rowIdx + 2, colIndex).setValue(nilai);
      sheetRaport.getRange(rowIdx + 2, colIndex + 1).setValue(hasilDeskripsiTinggi[mapelIdx][rowIdx][0]);
      sheetRaport.getRange(rowIdx + 2, colIndex + 2).setValue(hasilDeskripsiRendah[mapelIdx][rowIdx][0]);
    });
  });

  SpreadsheetApp.getUi().alert("Nilai akhir dan deskripsi berhasil diperbarui!");
}

function updateCatatan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNilai = ss.getSheetByName("Nilai");
  var sheetDeskripsi = ss.getSheetByName("Deskripsi");
  var sheetRaport = ss.getSheetByName("Raport");
  if (!sheetNilai || !sheetDeskripsi || !sheetRaport) {
    SpreadsheetApp.getUi().alert("Sheet 'Nilai', 'Deskripsi', atau 'Raport' tidak ditemukan!");
    return;
  }

  var nilaiRange = sheetNilai.getRange(2, 4, sheetNilai.getLastRow() - 1, 76).getValues();
  var catatanDeskripsi = sheetDeskripsi.getRange("B84:B86").getValues().flat();
  var catatanArray = [];
  nilaiRange.forEach(row => {
    var validNilai = row.filter(Number.isFinite);
    var rataRata = validNilai.length > 0 ? validNilai.reduce((a, b) => a + b, 0) / validNilai.length : 0;
    var catatan = "";
    if (rataRata >= 90) {
      catatan = catatanDeskripsi[0]; 
    } else if (rataRata >= 80) {
      catatan = catatanDeskripsi[1]; 
    } else if (rataRata >= 65) {
      catatan = catatanDeskripsi[2];
    }
    catatanArray.push([catatan]);
  });

  sheetNilai.getRange(2, 95, catatanArray.length, 1).setValues(catatanArray);
  sheetRaport.getRange(2, 65, catatanArray.length, 1).setValues(catatanArray);
  SpreadsheetApp.getUi().alert("Catatan guru berhasil diperbarui!");
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/1317/1317755.png')
    .setTitle('Raport')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getIdentitas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Identitas');
  const data = sheet.getRange('A2:F2').getValues()[0];
  return {
    sekolah: data[0],
    alamat: data[1],
    kelas: data[2],
    fase: data[3],
    semester: data[4],
    tahun: data[5]
  };
}

function getStudentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Raport');
  if (!sheet) return { nama: 'Tidak ditemukan', nis: '-' };
  const data = sheet.getRange('B2:C2').getValues();
  return {
    nama: data[0][0] || 'Tidak ada data',
    nis: data[0][1] || '-',
  };
}

function getLegalitasData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Identitas');
  const data = sheet.getRange('A5:E5').getValues()[0];
  return {
    tempatTanggal: data[0],
    namaGuru: data[1],
    nipGuru: data[2],
    kepalaSekolah: data[3],
    nipKepalaSekolah: data[4],
  };
}

function getStudentList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Raport');
  if (!sheet) return [];
  const data = sheet.getRange('B2:C' + sheet.getLastRow()).getValues();
  return data.map(([nama, nis]) => ({ nama, nis }));
}

function getCapaianKompetensi(namaSiswa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetData = ss.getSheetByName('Raport');
  const sheetMapel = ss.getSheetByName('Mapel');
  const muatanPelajaran = sheetMapel.getRange('B2:B16').getValues()
    .flat()
    .filter(mapel => mapel); 
  const data = sheetData.getRange(2, 2, sheetData.getLastRow() - 1, 48).getValues();
  const results = [];
  let nomorUrut = 1;
  const studentRow = data.find(row => row[0] === namaSiswa);
  if (!studentRow) return [];
  muatanPelajaran.forEach((mapel, i) => {
    results.push({
      id: nomorUrut++,
      mapel: mapel,
      nilai: studentRow[i * 3 + 2] || 0, 
      good: studentRow[i * 3 + 3] || 'Tidak Ada Mapel', 
      less: studentRow[i * 3 + 4] || 'Tidak Ada Mapel'
    });
  });
  return results;
}

function getEkstraData(namaSiswa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMapel = ss.getSheetByName('Mapel');
  const sheetRaport = ss.getSheetByName('Raport');
  if (!sheetMapel || !sheetRaport) return [];
  const namaEkstra = sheetMapel.getRange('B18:B22').getValues().flat();
  const dataRaport = sheetRaport.getRange(2, 2, sheetRaport.getLastRow() - 1, 59).getValues();
  const studentRow = dataRaport.find(row => row[0] === namaSiswa);
  if (!studentRow) return [];
  const ekstraData = studentRow.slice(48, 59);
  const results = [];
  let nomorUrut = 1;
  namaEkstra.forEach((ekstra, index) => {
    if (ekstra) {
      const predikat = ekstraData[index * 2] || '-';
      const keterangan = ekstraData[index * 2 + 1] || '-';
      results.push({
        no: nomorUrut++,
        ekstra,
        predikat,
        keterangan
      });
    }
  });
  return results;
}

function getPresensiData(namaSiswa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMapel = ss.getSheetByName('Mapel');
  const sheetRaport = ss.getSheetByName('Raport');
  if (!sheetMapel || !sheetRaport) return [];
  const jenisKetidakhadiran = sheetMapel.getRange('B24:B26').getValues().flat().filter(item => item);
  const data = sheetRaport.getRange(2, 2, sheetRaport.getLastRow() - 1, 62).getValues();
  const studentRow = data.find(row => row[0] === namaSiswa);
  if (!studentRow) return [];
  const jumlahPresensi = [60, 61, 62].map(index => studentRow[index - 1] || 0);
  return jenisKetidakhadiran.map((jenis, i) => ({
    no: i + 1,
    ketidakhadiran: jenis,
    keterangan: `${jumlahPresensi[i]} Hari`
  }));
}

function getCatatanSiswa(namaSiswa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRaport = ss.getSheetByName('Raport');
  if (!sheetRaport) return "Catatan tidak tersedia.";
  const dataRaport = sheetRaport.getRange(2, 2, sheetRaport.getLastRow() - 1, 65).getValues();
  const studentRow = dataRaport.find(row => row[0] === namaSiswa);
  if (!studentRow) return "Catatan tidak tersedia.";
  return studentRow[63] ? studentRow[63] : "Catatan tidak tersedia.";
}

function getKeputusanSiswa(namaSiswa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRaport = ss.getSheetByName('Raport');
  if (!sheetRaport) return { naik: "-", tinggal: "-" };
  const dataRaport = sheetRaport.getRange(2, 2, sheetRaport.getLastRow() - 1, 69).getValues();
  const studentRow = dataRaport.find(row => row[0] === namaSiswa);
  if (!studentRow) return { naik: "-", tinggal: "-" };
  return {
    naik: studentRow[66] || "-",
    tinggal: studentRow[67] || "-"
  };
  // Fungsi ini untuk menampilkan halaman login (login.html)
let sessionNama = '';

function doGet() {
  const nama = sessionNama;
  if (!nama) {
    return HtmlService.createHtmlOutputFromFile('login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('Login Rapor');
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/1317/1317755.png')
    .setTitle('Raport')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setSessionNama(nama) {
  sessionNama = nama;
}

function getSessionNama() {
  return sessionNama;
}


}

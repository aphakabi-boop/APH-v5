function doGet(e) {
  const page = e.parameter.page;
  
  if (page === 'rekap') {
    return HtmlService.createTemplateFromFile('Rekap.html')
      .evaluate()
      .setTitle('Rekap Stok & Produksi APH')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('APH')
    .evaluate()
    .setTitle('Form Data Stok APH Kegiatan P4')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function forceAuth() {
  DriveApp.getRootFolder();
}


// ========================================================================= //
// FUNGSI REKAP MULTI-TANGGAL: DARI DataAPH1 KE Data_Stok (23 KOLOM)
// ========================================================================= //
function rekapKeDataStok() {
  const SPREADSHEET_ID = "1l64-CR_QVUkQJXmVCtmhFHD0xmk0lORYq80Go15gTXI";
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const sheetSumber = ss.getSheetByName('DataAPH1');
  const sheetRekap = ss.getSheetByName('Data_Stok');
  
  const data = sheetSumber.getDataRange().getValues();
  if (data.length < 5) return;
  
  // Fungsi bantu: Format tanggal & Parse Volume
  function formatPeriode(val) {
    if (!val || val === "") return "-";
    let tgl = new Date(val);
    if (isNaN(tgl.getTime())) return "-";
    let b = ("0" + (tgl.getMonth() + 1)).slice(-2);
    let t = tgl.getFullYear();
    return b + "-" + t;
  }

  function parseVolume(val) {
    if (!val) return 0;
    let num = parseFloat(String(val).replace(",", "."));
    return isNaN(num) ? 0 : num;
  }

  // Objek untuk mengelompokkan data berdasarkan Jenis APH secara independen
  const mapEks = new Map();
  const mapLphp = new Map();

  function getMapItem(map, prov, poktan, program, jenis, sat) {
    let j = String(jenis || "").trim();
    if (!j || j === '-' || j === '') return null;
    let s = String(sat || "-").trim();
    let key = `${prov}|${poktan}|${program}|${j}|${s}`;
    
    if (!map.has(key)) {
      map.set(key, {
        prov: prov, poktan: poktan, program: program, jenis: j, sat: s,
        tglProd: '-', tglPerb: '-', tglPem: '-',
        stok: 0, prod: 0, perb: 0, pem: 0
      });
    }
    return map.get(key);
  }

  // --- 1. EKSTRAKSI INDEPENDEN (MEMECAH BARIS BERDASARKAN AKTIVITAS) ---
  data.forEach((row, index) => {
    if (index < 4) return; 
    
    let prov = String(row[1] || "-").trim();
    let poktan = String(row[5] || "-").trim();
    let program = String(row[6] || "P4").trim();
    
    if (!prov || prov === "-" || prov.toUpperCase() === "PROVINSI" || prov === "0") return;

    // ----- A. DATA EKSPLORASI -----
    // 1. Stok Awal Eksplorasi (J, L, K)
    let eksStok = getMapItem(mapEks, prov, poktan, program, row[9], row[11]);
    if (eksStok) eksStok.stok += parseVolume(row[10]);

    // 2. Produksi Eksplorasi (S, U, T, V)
    let eksProd = getMapItem(mapEks, prov, poktan, program, row[18], row[20]);
    if (eksProd) { 
        eksProd.prod += parseVolume(row[19]); 
        let tgl = formatPeriode(row[21]);
        if (tgl !== "-") eksProd.tglProd = tgl;
    }

    // 3. Perbanyakan Eksplorasi (AD, AF, AE, AG)
    let eksPerb = getMapItem(mapEks, prov, poktan, program, row[29], row[31]);
    if (eksPerb) { 
        eksPerb.perb += parseVolume(row[30]); 
        let tgl = formatPeriode(row[32]);
        if (tgl !== "-") eksPerb.tglPerb = tgl;
    }

    // 4. Pemanfaatan Eksplorasi (AO, AQ, AP, AU)
    let eksPem = getMapItem(mapEks, prov, poktan, program, row[40], row[42]);
    if (eksPem) { 
        eksPem.pem += parseVolume(row[41]); 
        let tgl = formatPeriode(row[46]);
        if (tgl !== "-") eksPem.tglPem = tgl;
    }

    // ----- B. DATA ISOLAT / LPHP -----
    // 1. Stok Awal LPHP (N, P, O)
    let lphpStok = getMapItem(mapLphp, prov, poktan, program, row[13], row[15]);
    if (lphpStok) lphpStok.stok += parseVolume(row[14]);

    // 2. Produksi LPHP (X, Z, Y, AA)
    let lphpProd = getMapItem(mapLphp, prov, poktan, program, row[23], row[25]);
    if (lphpProd) { 
        lphpProd.prod += parseVolume(row[24]); 
        let tgl = formatPeriode(row[26]);
        if (tgl !== "-") lphpProd.tglProd = tgl;
    }

    // 3. Perbanyakan LPHP (AI, AK, AJ, AL)
    let lphpPerb = getMapItem(mapLphp, prov, poktan, program, row[34], row[36]);
    if (lphpPerb) { 
        lphpPerb.perb += parseVolume(row[35]); 
        let tgl = formatPeriode(row[37]);
        if (tgl !== "-") lphpPerb.tglPerb = tgl;
    }

    // 4. Pemanfaatan LPHP (AW, AY, AX, BC)
    let lphpPem = getMapItem(mapLphp, prov, poktan, program, row[48], row[50]);
    if (lphpPem) { 
        lphpPem.pem += parseVolume(row[49]); 
        let tgl = formatPeriode(row[54]);
        if (tgl !== "-") lphpPem.tglPem = tgl;
    }
  });

  // --- 2. PENYUSUNAN & PENGGABUNGAN DATA (ALIGNMENT) ---
  const groupData = new Map();
  function getGroup(prov, poktan, prog) {
    let key = `${prov}|${poktan}|${prog}`;
    if (!groupData.has(key)) groupData.set(key, { eks: [], lphp: [] });
    return groupData.get(key);
  }

  // Masukkan data yang sudah dinormalisasi ke grup wilayah
  for (let item of mapEks.values()) getGroup(item.prov, item.poktan, item.program).eks.push(item);
  for (let item of mapLphp.values()) getGroup(item.prov, item.poktan, item.program).lphp.push(item);

  let hasilAkhir = [];
  
  for (let [key, group] of groupData.entries()) {
    let [prov, poktan, prog] = key.split('|');
    let maxLen = Math.max(group.eks.length, group.lphp.length);
    
    // Rangkai baris sejajar untuk Data_Stok
    for (let i = 0; i < maxLen; i++) {
      let e = group.eks[i] || { tglProd:'-', tglPerb:'-', tglPem:'-', jenis:'-', sat:'-', stok:0, prod:0, perb:0, pem:0 };
      let l = group.lphp[i] || { tglProd:'-', tglPerb:'-', tglPem:'-', jenis:'-', sat:'-', stok:0, prod:0, perb:0, pem:0 };

      let sisaE = parseFloat((e.stok + e.prod + e.perb - e.pem).toFixed(4));
      let sisaL = parseFloat((l.stok + l.prod + l.perb - l.pem).toFixed(4));

      hasilAkhir.push([
        e.tglProd, e.tglPerb, e.tglPem, l.tglProd, l.tglPerb, l.tglPem, prov, poktan, prog,
        e.jenis, e.sat, e.stok, e.prod, e.perb, e.pem, sisaE,
        l.jenis, l.sat, l.stok, l.prod, l.perb, l.pem, sisaL
      ]);
    }
  }
  
  // --- 3. EKSEKUSI TULIS KE SPREADSHEET ---
  if (sheetRekap.getLastRow() > 1) {
    sheetRekap.getRange(2, 1, sheetRekap.getLastRow() - 1, 23).clearContent();
  }
  
  if (hasilAkhir.length > 0) {
    hasilAkhir.sort((a, b) => {
      if (a[6] !== b[6]) return a[6].localeCompare(b[6]);
      return a[7].localeCompare(b[7]);
    });
    
    sheetRekap.getRange(2, 1, hasilAkhir.length, 23).setValues(hasilAkhir);
  }
  
  SpreadsheetApp.flush();
}

function parseVolume(val) {
  if (!val) return 0;
  let num = parseFloat(String(val).replace(",", "."));
  return isNaN(num) ? 0 : num;
}

// ========================================================================= //
// 2. FUNGSI UTAMA UNTUK MENAMPILKAN DATA TABEL & FILTER (23 KOLOM)
// ========================================================================= //
function getRekapDataServer(bulanMulai, bulanAkhir, provinsiFilter, poktanFilter, programFilter) {
  const SPREADSHEET_ID = "1l64-CR_QVUkQJXmVCtmhFHD0xmk0lORYq80Go15gTXI";
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Stok');
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  
  data.shift(); // Hapus header
  let hasil = [];
  
  function toBulanIndex(str) {
    if (!str || str === 'ALL' || str === '') return null;
    let parts = String(str).split('-');
    if (parts.length === 2) return parseInt(parts[1] + parts[0]); // Format MM-YYYY -> YYYYMM
    return 0;
  }
  
  let idxMulai = toBulanIndex(bulanMulai);
  let idxAkhir = toBulanIndex(bulanAkhir);
  
  const fProv = (provinsiFilter && provinsiFilter !== 'ALL' && provinsiFilter !== '') ? String(provinsiFilter).trim().toLowerCase() : null;
  const fPoktan = (poktanFilter && poktanFilter !== 'ALL' && poktanFilter !== '') ? String(poktanFilter).trim().toLowerCase() : null;
  const fProg = (programFilter && programFilter !== 'ALL' && programFilter !== '' && !String(programFilter).includes('--')) ? String(programFilter).trim().toLowerCase() : null;

  data.forEach(row => {
    // Cek apakah ada minimal salah satu dari Kolom A-F yang masuk dalam rentang filter bulan
    let matchBulan = true;
    if (idxMulai !== null || idxAkhir !== null) {
      let adaYangMasuk = false;
      for (let i = 0; i <= 5; i++) {
        let rawB = row[i];
        let bStr = "";
        if (rawB instanceof Date) {
          bStr = Utilities.formatDate(rawB, Session.getScriptTimeZone(), "MM-yyyy");
        } else {
          bStr = String(rawB || '').trim();
        }
        let idxB = toBulanIndex(bStr);
        if (idxB !== null) {
          let pass = true;
          if (idxMulai !== null && idxB < idxMulai) pass = false;
          if (idxAkhir !== null && idxB > idxAkhir) pass = false;
          if (pass) {
            adaYangMasuk = true;
            break;
          }
        }
      }
      matchBulan = adaYangMasuk;
    }

    // AMBIL DATA BERDASARKAN POSISI KOLOM YANG TEPAT:
    const provData = String(row[6] || '').trim().toLowerCase();   // Kolom G: Provinsi
    const poktanData = String(row[7] || '').trim().toLowerCase(); // Kolom H: Kelompok Tani
    const programData = String(row[8] || '').trim().toLowerCase(); // Kolom I: Program
    
    const matchProv = (!fProv || provData === fProv);
    const matchPoktan = (!fPoktan || poktanData === fPoktan);
    const matchProgram = (!fProg || programData === fProg);

    if (matchBulan && matchProv && matchPoktan && matchProgram) {
      // Format tanggal Date object di seluruh baris agar aman dikirim ke web
      let safeRow = row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd-MM-yyyy") : cell);
      hasil.push(safeRow);
    }
  });
  
  return hasil;
}

// ========================================================================= //
// FUNGSI UNTUK MENGISI DROPDOWN FILTER DI WEB SECARA OTOMATIS (TERMASUK PROGRAM)
// ========================================================================= //

// ========================================================================= //
// 1. FUNGSI UNTUK MENGAMBIL OPSI FILTER (DINAMIS & AMAN DARI DATE OBJECT)
// ========================================================================= //
function getFilterOptions() {
  const SPREADSHEET_ID = "1l64-CR_QVUkQJXmVCtmhFHD0xmk0lORYq80Go15gTXI";
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Stok');
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { bulan: [], allProvinsi: [], allPoktan: [], programMap: {} };
  
  data.shift(); // Hapus baris header
  
  let bulanSet = new Set();
  let allProvinsi = new Set();
  let allPoktan = new Set();
  let programMap = {}; 

data.forEach(row => {
    // 1. AMBIL BULAN DARI KOLOM A SAMPAI F (Index 0 sampai 5)
    for (let i = 0; i <= 5; i++) {
      let rawBulan = row[i];
      if (rawBulan instanceof Date) {
        let bStr = Utilities.formatDate(rawBulan, Session.getScriptTimeZone(), "MM-yyyy");
        if (bStr) bulanSet.add(bStr);
      } else if (rawBulan) {
        let bStr = String(rawBulan).trim();
        if (bStr && bStr !== '-' && bStr.length >= 6) bulanSet.add(bStr);
      }
    }

    // 2. KOREKSI INDEKS KOLOM SESUAI SPREADSHEET ASLI ANDA:
    let prov = row[6] ? String(row[6]).trim() : '';       // Kolom G: Provinsi
    let poktan = row[7] ? String(row[7]).trim() : '';     // Kolom H: Kelompok Tani
    let program = row[8] ? String(row[8]).trim() : 'Umum'; // Kolom I: Program

    // 3. MASUKKAN KE DROPDOWN
    if (prov && prov !== '-' && prov !== '') allProvinsi.add(prov);
    if (poktan && poktan !== '-' && poktan !== '') allPoktan.add(poktan);

    if (program && program !== '-' && program !== '' && prov) {
      if (!programMap[program]) programMap[program] = {};
      if (!programMap[program][prov]) programMap[program][prov] = new Set();
      if (poktan) programMap[program][prov].add(poktan);
    }
  });

  // Konversi Set ke Array terurut
  let bulanArr = Array.from(bulanSet).sort();
  let provArr = Array.from(allProvinsi).sort();
  let poktanArr = Array.from(allPoktan).sort();

  let formattedProgramMap = {};
  for (let prog in programMap) {
    formattedProgramMap[prog] = {};
    for (let prov in programMap[prog]) {
      formattedProgramMap[prog][prov] = Array.from(programMap[prog][prov]).sort();
    }
  }

  return {
    bulan: bulanArr,
    allProvinsi: provArr,
    allPoktan: poktanArr,
    programMap: formattedProgramMap
  };
}


// ===============================
// FOLDER EVIDEN DRIVE
// ===============================
function getOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function createEvidenceFolder_(provinsi, kabupaten, penerima) {
  const ROOT_FOLDER_ID = "1TLPkeYdPhuQTH6YaFxYZTsUjSc2NZYVZ";
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);

  const now = new Date();
  const tahun = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
  const tanggal = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const folderTahun = getOrCreateFolder_(root, tahun);
  const folderProv = getOrCreateFolder_(folderTahun, provinsi);
  const folderKab = getOrCreateFolder_(folderProv, kabupaten);

  const safePenerima = penerima.replace(/[\\/:*?"<>|]/g, "-");
  const folderPen = getOrCreateFolder_(folderKab, safePenerima);
  const folderLaporan = folderPen.createFolder(`APH_${tanggal}`);

  return folderLaporan;
}

// ===============================
// SUBMIT FORM (google.script.run)
// ===============================
function submitForm(formObject) {
  try {
    const SPREADSHEET_ID = "1l64-CR_QVUkQJXmVCtmhFHD0xmk0lORYq80Go15gTXI";
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("DataAPH1");

    if (!sheet) throw new Error("Sheet DataAPH1 tidak ditemukan");

    const provinsi  = formObject.provinsi || "";
    const kabupaten = formObject.kabupaten || "";
    const kecamatan = formObject.kecamatan || "";
    const desa      = formObject.desa || "";
    const penerima  = formObject.penerima || "";

    const A = v => Array.isArray(v) ? v : (v ? [v] : []);

    const fotoBase64 = A(formObject["fotoBase64[]"]);
    let folderUrl = "";

    if (fotoBase64.length > 0 && fotoBase64[0]) {
      const folder = createEvidenceFolder_(provinsi, kabupaten, penerima);

      fotoBase64.forEach((data, i) => {
        const matches = data.match(/^data:(image\/\w+);base64,/);
        if (!matches) return;

        const mimeType = matches[1];
        const ext = mimeType.split("/")[1];

        const bytes = Utilities.base64Decode(data.replace(/^data:image\/\w+;base64,/, ""));
        const blob = Utilities.newBlob(bytes, mimeType, `eviden_${i + 1}.${ext}`);

        folder.createFile(blob);
      });

      folderUrl = folder.getUrl();
    }

    const lastRow = sheet.getLastRow();
    let nextNumber = 1;

    if (lastRow > 1) {
      const lastValue = sheet.getRange(lastRow, 1).getValue();
      nextNumber = Number(lastValue) ? Number(lastValue) + 1 : lastRow;
    }

    const programName = formObject.jenis_program || "P4"; // Tangkap Jenis Program (Kolom G)

    const stoExplorasi = A(formObject["stok-explorasi[]"]);
    const stoVolExplorasi = A(formObject["stok-explorasi_volume[]"]);
    const stoSatExplorasi = A(formObject["stok-explorasi_satuan[]"]);

    const stoIsolat = A(formObject["stok-isolat[]"]);
    const stoVolIsolat = A(formObject["stok-isolat_volume[]"]);
    const stoSatIsolat = A(formObject["stok-isolat_satuan[]"]);

    const prodExplorasi = A(formObject["produksi-explorasi[]"]);
    const prodVolExplorasi = A(formObject["produksi-explorasi_volume[]"]);
    const prodSatExplorasi = A(formObject["produksi-explorasi_satuan[]"]);
    const prodTglExplorasi = A(formObject["produksi-explorasi_tanggal[]"]);

    const prodIsolat = A(formObject["produksi-isolat[]"]);
    const prodVolIsolat = A(formObject["produksi-isolat_volume[]"]);
    const prodSatIsolat = A(formObject["produksi-isolat_satuan[]"]);
    const prodTglIsolat = A(formObject["produksi-isolat_tanggal[]"]);

    const perbanyakanExplorasi = A(formObject["perbanyakan-explorasi[]"]);
    const perbVolExplorasi = A(formObject["perbanyakan-explorasi_volume[]"]);
    const perbSatExplorasi = A(formObject["perbanyakan-explorasi_satuan[]"]);
    const perbTglExplorasi = A(formObject["perbanyakan-explorasi_tanggal[]"]);

    const perbanyakanIsolat = A(formObject["perbanyakan-isolat[]"]);
    const perbVolIsolat = A(formObject["perbanyakan-isolat_volume[]"]);
    const perbSatIsolat = A(formObject["perbanyakan-isolat_satuan[]"]);
    const perbTglIsolat = A(formObject["perbanyakan-isolat_tanggal[]"]);

    const manfaatExplorasi = A(formObject["pemanfaatan-explorasi[]"]);
    const manfaatVolExplorasi = A(formObject["pemanfaatan-explorasi_volume[]"]);
    const manfaatSatExplorasi = A(formObject["pemanfaatan-explorasi_satuan[]"]);
    const manfaatKomExplorasi = A(formObject["pemanfaatan-explorasi_komoditas[]"]);
    const manfaatLuasExplorasi = A(formObject["pemanfaatan-explorasi_luasan[]"]);
    const manfaatLokExplorasi = A(formObject["pemanfaatan-explorasi_lokasi[]"]);
    const manfaatLokLainExplorasi = A(formObject["pemanfaatan-explorasi_lokasi_lain[]"]);
    const manfaatTglExplorasi = A(formObject["pemanfaatan-explorasi_tanggal[]"]);

    const manfaatIsolat = A(formObject["pemanfaatan-isolat[]"]);
    const manfaatVolIsolat = A(formObject["pemanfaatan-isolat_volume[]"]);
    const manfaatSatIsolat = A(formObject["pemanfaatan-isolat_satuan[]"]);
    const manfaatKomIsolat = A(formObject["pemanfaatan-isolat_komoditas[]"]);
    const manfaatLuasIsolat = A(formObject["pemanfaatan-isolat_luasan[]"]);
    const manfaatLokIsolat = A(formObject["pemanfaatan-isolat_lokasi[]"]);
    const manfaatLokLainIsolat = A(formObject["pemanfaatan-isolat_lokasi_lain[]"]);
    const manfaatTglIsolat = A(formObject["pemanfaatan-isolat_tanggal[]"]);

    const optJenis = A(formObject["jenis_opt[]"]);
    const optJenisLain = A(formObject["jenis_opt_lain[]"]);
    const optIntensitas = A(formObject["intensitas_serangan[]"]);

    const sisExplorasi = A(formObject["sisa-explorasi[]"]);
    const sisVolExplorasi = A(formObject["sisa-explorasi_volume[]"]);
    const sisSatExplorasi = A(formObject["sisa-explorasi_satuan[]"]);

    const sisIsolat = A(formObject["sisa-isolat[]"]);
    const sisVolIsolat = A(formObject["sisa-isolat_volume[]"]);
    const sisSatIsolat = A(formObject["sisa-isolat_satuan[]"]);

    const maxRows = Math.max(
      A(formObject["stok-explorasi[]"]).length,
      A(formObject["stok-isolat[]"]).length,
      A(formObject["produksi-explorasi[]"]).length,
      A(formObject["produksi-isolat[]"]).length,
      A(formObject["perbanyakan-explorasi[]"]).length,
      A(formObject["perbanyakan-isolat[]"]).length,
      A(formObject["pemanfaatan-explorasi[]"]).length,
      A(formObject["pemanfaatan-isolat[]"]).length,
      A(formObject["jenis_opt[]"]).length,
      A(formObject["sisa-explorasi[]"]).length,
      A(formObject["sisa-isolat[]"]).length
    );

    for (let i = 0; i < maxRows; i++) {
      const lokasiExpl = manfaatLokLainExplorasi[i]?.trim() || manfaatLokExplorasi[i] || "";
      const lokasiIso = manfaatLokLainIsolat[i]?.trim() || manfaatLokIsolat[i] || "";

      let finalOpt = optJenis[i] || "";
      if (finalOpt === "__other__") finalOpt = optJenisLain[i] || "";

      let vStoEks = stoVolExplorasi[i] ? parseVolume(stoVolExplorasi[i]) : "";
      let vStoIso = stoVolIsolat[i] ? parseVolume(stoVolIsolat[i]) : "";
      
      let vProdEks = prodVolExplorasi[i] ? parseVolume(prodVolExplorasi[i]) : "";
      let vProdIso = prodVolIsolat[i] ? parseVolume(prodVolIsolat[i]) : "";
      
      let vPerbEks = perbVolExplorasi[i] ? parseVolume(perbVolExplorasi[i]) : "";
      let vPerbIso = perbVolIsolat[i] ? parseVolume(perbVolIsolat[i]) : "";

      let vManEks = manfaatVolExplorasi[i] ? parseVolume(manfaatVolExplorasi[i]) : "";
      let vManIso = manfaatVolIsolat[i] ? parseVolume(manfaatVolIsolat[i]) : "";
      
      let vSisEks = sisVolExplorasi[i] ? parseVolume(sisVolExplorasi[i]) : "";
      let vSisIso = sisVolIsolat[i] ? parseVolume(sisVolIsolat[i]) : "";

      sheet.appendRow([
        nextNumber++, // A: No Urut
        provinsi, kabupaten, kecamatan, desa, penerima, // B - F: Data Umum
        programName, // G: Jenis Program

        "Stok APH",
        stoExplorasi[i] ? "Hasil Eksplorasi" : "",
        stoExplorasi[i] || "", vStoEks, stoSatExplorasi[i] || "",
        stoIsolat[i] ? "Isolat LPHP" : "",
        stoIsolat[i] || "", vStoIso, stoSatIsolat[i] || "",

        "Produksi APH",
        prodExplorasi[i] ? "Hasil Eksplorasi" : "",
        prodExplorasi[i] || "", vProdEks, prodSatExplorasi[i] || "", prodTglExplorasi[i] || "",
        prodIsolat[i] ? "Isolat LPHP" : "",
        prodIsolat[i] || "", vProdIso, prodSatIsolat[i] || "", prodTglIsolat[i] || "",

        "Perbanyakan APH", // AB
        perbanyakanExplorasi[i] ? "Hasil Eksplorasi" : "", // AC
        perbanyakanExplorasi[i] || "", vPerbEks, perbSatExplorasi[i] || "", perbTglExplorasi[i] || "", // AD, AE, AF, AG
        perbanyakanIsolat[i] ? "Isolat LPHP" : "", // AH
        perbanyakanIsolat[i] || "", vPerbIso, perbSatIsolat[i] || "", perbTglIsolat[i] || "", // AI, AJ, AK, AL

        "Pemanfaatan APH",
        manfaatExplorasi[i] ? "Hasil Eksplorasi" : "",
        manfaatExplorasi[i] || "", vManEks, manfaatSatExplorasi[i] || "",
        manfaatKomExplorasi[i] || "", manfaatLuasExplorasi[i] || "", lokasiExpl, manfaatTglExplorasi[i] || "", 
        manfaatIsolat[i] ? "Isolat LPHP" : "",
        manfaatIsolat[i] || "", vManIso, manfaatSatIsolat[i] || "",
        manfaatKomIsolat[i] || "", manfaatLuasIsolat[i] || "", lokasiIso, manfaatTglIsolat[i] || "",

        // == KATEGORI SERANGAN OPT ==
        finalOpt, // BD: Jenis OPT
        optIntensitas[i] || "", // BE: Intensitas Serangan

        (sisExplorasi[i] || sisIsolat[i]) ? "Sisa APH" : "",
        sisExplorasi[i] ? "Hasil Eksplorasi" : "",
        sisExplorasi[i] || "", vSisEks, sisSatExplorasi[i] || "",
        sisIsolat[i] ? "Isolat LPHP" : "",
        sisIsolat[i] || "", vSisIso, sisSatIsolat[i] || "",

        folderUrl,
        new Date()
      ]);
    }

    // --- PANGGIL MESIN REKAP OTOMATIS ---
    SpreadsheetApp.flush(); 
    rekapKeDataStok(); 

    // --- KIRIM NOTIFIKASI TELEGRAM ---
    let pesanNotif = "🚨 <b>LAPORAN APH BARU MASUK!</b>\n\n" +
                     "<b>Provinsi:</b> " + provinsi + "\n" +
                     "<b>Kab/Kota:</b> " + kabupaten + "\n" +
                     "<b>Kelompok Tani:</b> " + penerima + "\n\n" +
                     "<i>Data telah berhasil masuk ke Database dan Dashboard. Silakan cek rekap!</i> ✅";
                     
    kirimNotifTelegram(pesanNotif);

    return "✅ Data & foto eviden berhasil disimpan";

  } catch (err) {
    return "❌ ERROR: " + err.message;
  }
}

function parseVolume(volume) {
  if (volume === undefined || volume === null || volume === "") return 0;
  
  if (volume instanceof Date || Object.prototype.toString.call(volume) === '[object Date]') {
    return 0; 
  }
  
  let str = String(volume).trim();
  str = str.replace(/\s/g, ""); 
  str = str.replace(/[^0-9,.]/g, ""); 
  if (!str) return 0;

  if (str.includes(",") && str.includes(".")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (str.includes(",")) {
    str = str.replace(",", ".");
  } else if (str.includes(".")) {
    const bagian = str.split(".");
    if (bagian.length === 2 && bagian[1].length <= 2) {
      // biarkan sebagai desimal
    } else if (bagian.length === 2 && bagian[1].length === 3) {
      str = str.replace(".", "");
    } else if (bagian.length > 2) {
      str = str.replace(/\./g, "");
    }
  }

  const angka = Number(str);
  return Number.isFinite(angka) ? angka : 0;
}

// ==========================================
// FITUR NOTIFIKASI TELEGRAM OTOMATIS
// ==========================================
function kirimNotifTelegram(pesan) {
  const BOT_TOKEN = "8739477560:AAHb7Zku7B8AlB-d6egsRcCWgnaLweWdwxM"; 
  const CHAT_ID = "321489779"; 

  const url = "https://api.telegram.org/bot" + BOT_TOKEN + "/sendMessage";

  const payload = {
    "chat_id": CHAT_ID,
    "text": pesan,
    "parse_mode": "HTML",
    "disable_web_page_preview": true
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("Gagal mengirim Telegram: " + e.message);
  }
}
// ==========================================
// FILE: PlotSiswa.js
// FUNGSI: Modul Relasi Siswa ke Kelas (Rombel)
// ==========================================
function getInitPlotSiswa() {
  try {
    const sheetKelas = DB.getSheet('mst_kelas');
    const dataKelasRaw = sheetKelas.getDataRange().getDisplayValues();
    const headKelas = dataKelasRaw.length > 0 ? dataKelasRaw[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_')) : [];
    
    let kelasList = [];
    for(let i=1; i<dataKelasRaw.length; i++) {
      let obj = {};
      headKelas.forEach((h, idx) => obj[h] = dataKelasRaw[i][idx]);
      if((obj.deleted_at === "" || !obj.deleted_at) && (obj.is_active === 'TRUE' || obj.is_active === true)) {
        kelasList.push({ id_kelas: obj.id_kelas, nama_kelas: obj.nama_kelas, tingkat: obj.tingkat });
      }
    }
    const sheetSiswa = DB.getSheet('mst_siswa');
    const dataSiswaRaw = sheetSiswa.getDataRange().getDisplayValues();
    const headSiswa = dataSiswaRaw.length > 0 ? dataSiswaRaw[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_')) : [];
    
    let siswaList = [];
    for(let i=1; i<dataSiswaRaw.length; i++) {
      let obj = {};
      headSiswa.forEach((h, idx) => obj[h] = dataSiswaRaw[i][idx]);
      if((obj.deleted_at === "" || !obj.deleted_at) && (obj.is_active === 'TRUE' || obj.is_active === true)) {
        siswaList.push({ id_siswa: obj.id_siswa, nis: obj.nis, nama_lengkap: obj.nama_lengkap, jk: obj.jk });
      }
    }
    return { status: "success", kelas: kelasList, siswa: siswaList };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function getSiswaByKelas(tahun_ajaran, id_kelas) {
  try {
    const sheet = DB.getSheet('trx_kelas_siswa');
    let allData = sheet.getDataRange().getDisplayValues();
    
    // PERBAIKAN: Menyesuaikan dengan Header Spreadsheet Anda
    if(allData.length === 0 || allData[0].length < 4) {
      sheet.getRange(1, 1, 1, 4).setValues([['id_kelas_siswa', 'id_tapel', 'id_kelas', 'id_siswa']]);
      allData = sheet.getDataRange().getDisplayValues();
    }
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxTapel = headers.indexOf('id_tapel');
    const idxKelas = headers.indexOf('id_kelas');
    const idxSiswa = headers.indexOf('id_siswa');
    let plottedInThis = [];
    let plottedInOther = [];
    if(idxTapel > -1 && idxKelas > -1 && idxSiswa > -1) {
      for(let i=1; i<allData.length; i++) {
        if(allData[i][idxTapel].toString().trim() === tahun_ajaran.toString().trim()) {
          if(allData[i][idxKelas].toString().trim() === id_kelas.toString().trim()) {
            plottedInThis.push(allData[i][idxSiswa].toString().trim());
          } else {
            plottedInOther.push({ 
              id_siswa: allData[i][idxSiswa].toString().trim(), 
              id_kelas: allData[i][idxKelas].toString().trim() 
            });
          }
        }
      }
    }
    
    return { status: "success", inThisClass: plottedInThis, inOtherClass: plottedInOther };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function savePlottingSiswa(tahun_ajaran, id_kelas, arr_id_siswa) {
  try {
    const sheet = DB.getSheet('trx_kelas_siswa');
    let allData = sheet.getDataRange().getValues();
    
    // PERBAIKAN: Menyesuaikan dengan Header Spreadsheet Anda
    if (allData.length === 0 || allData[0].length < 4) {
      sheet.getRange(1, 1, 1, 4).setValues([['id_kelas_siswa', 'id_tapel', 'id_kelas', 'id_siswa']]);
      allData = sheet.getDataRange().getValues();
    }
    let headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    
    let idxId = headers.indexOf('id_kelas_siswa');
    let idxTapel = headers.indexOf('id_tapel');
    let idxKelas = headers.indexOf('id_kelas');
    let idxSiswa = headers.indexOf('id_siswa');
    
    if (idxId === -1 || idxTapel === -1 || idxKelas === -1 || idxSiswa === -1) {
      sheet.getRange(1, 1, 1, 4).setValues([['id_kelas_siswa', 'id_tapel', 'id_kelas', 'id_siswa']]);
      headers = ['id_kelas_siswa', 'id_tapel', 'id_kelas', 'id_siswa'];
      idxId = 0; idxTapel = 1; idxKelas = 2; idxSiswa = 3;
    }
    
    // HAPUS DATA LAMA (Metode Wipe baris per baris dari bawah)
    for (let i = allData.length - 1; i > 0; i--) {
      if (allData[i][idxTapel].toString().trim() === tahun_ajaran.toString().trim() && 
          allData[i][idxKelas].toString().trim() === id_kelas.toString().trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    // INSERT DATA BARU
    if (arr_id_siswa && arr_id_siswa.length > 0) {
      let newRows = [];
      arr_id_siswa.forEach(id_siswa => {
        let row = new Array(headers.length).fill("");
        row[idxId] = Utilities.getUuid();
        row[idxTapel] = tahun_ajaran;
        row[idxKelas] = id_kelas;
        row[idxSiswa] = id_siswa;
        newRows.push(row);
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    }
    
    return { status: "success", message: "Plotting rombel berhasil disimpan!" };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function getRekapPlotting(tahun_ajaran) {
  try {
    const sheetKelas = DB.getSheet('mst_kelas');
    const dataKelasRaw = sheetKelas.getDataRange().getDisplayValues();
    const headK = dataKelasRaw[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let mapKelas = {};
    for(let i=1; i<dataKelasRaw.length; i++) {
      if((dataKelasRaw[i][headK.indexOf('deleted_at')] === "" || !dataKelasRaw[i][headK.indexOf('deleted_at')]) && (dataKelasRaw[i][headK.indexOf('is_active')] === 'TRUE' || dataKelasRaw[i][headK.indexOf('is_active')] === true)) {
         mapKelas[dataKelasRaw[i][headK.indexOf('id_kelas')]] = dataKelasRaw[i][headK.indexOf('nama_kelas')];
      }
    }
    const sheetSiswa = DB.getSheet('mst_siswa');
    const dataSiswaRaw = sheetSiswa.getDataRange().getDisplayValues();
    const headS = dataSiswaRaw[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let allSiswa = [];
    let mapSiswa = {};
    for(let i=1; i<dataSiswaRaw.length; i++) {
      if((dataSiswaRaw[i][headS.indexOf('deleted_at')] === "" || !dataSiswaRaw[i][headS.indexOf('deleted_at')]) && (dataSiswaRaw[i][headS.indexOf('is_active')] === 'TRUE' || dataSiswaRaw[i][headS.indexOf('is_active')] === true)) {
         let s = {
           id_siswa: dataSiswaRaw[i][headS.indexOf('id_siswa')],
           nis: dataSiswaRaw[i][headS.indexOf('nis')],
           nama_lengkap: dataSiswaRaw[i][headS.indexOf('nama_lengkap')],
           jk: dataSiswaRaw[i][headS.indexOf('jk')]
         };
         allSiswa.push(s);
         mapSiswa[s.id_siswa] = s;
      }
    }
    const sheetTrx = DB.getSheet('trx_kelas_siswa');
    let allTrx = sheetTrx.getDataRange().getDisplayValues();
    
    // PERBAIKAN: Menyesuaikan dengan Header Spreadsheet Anda
    if(allTrx.length === 0 || allTrx[0].length < 4) {
      sheetTrx.getRange(1, 1, 1, 4).setValues([['id_kelas_siswa', 'id_tapel', 'id_kelas', 'id_siswa']]);
      allTrx = sheetTrx.getDataRange().getDisplayValues();
    }
    let plottedIds = new Set();
    let rekapKelas = {}; 
    if(allTrx.length > 1) {
      const headT = allTrx[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
      const idxTapel = headT.indexOf('id_tapel');
      const idxKelas = headT.indexOf('id_kelas');
      const idxSiswa = headT.indexOf('id_siswa');
      if(idxTapel > -1 && idxKelas > -1 && idxSiswa > -1) {
        for(let i=1; i<allTrx.length; i++) {
          if(allTrx[i][idxTapel].toString().trim() === tahun_ajaran.toString().trim()) {
             let idK = allTrx[i][idxKelas].toString().trim();
             let idS = allTrx[i][idxSiswa].toString().trim();
             plottedIds.add(idS);
             if(!rekapKelas[idK]) rekapKelas[idK] = [];
             if(mapSiswa[idS]) rekapKelas[idK].push(mapSiswa[idS]);
          }
        }
      }
    }
    let belumPlot = allSiswa.filter(s => !plottedIds.has(s.id_siswa));
    
    let rekapArray = [];
    for (let idK in mapKelas) { 
       rekapArray.push({
         id_kelas: idK,
         nama_kelas: mapKelas[idK],
         siswa: rekapKelas[idK] || []
       });
    }
    rekapArray.sort((a,b) => a.nama_kelas.localeCompare(b.nama_kelas));
    return { status: 'success', rekapKelas: rekapArray, belumPlot: belumPlot };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}
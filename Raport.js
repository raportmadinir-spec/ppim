// ==========================================
// FILE: Raport.js
// FUNGSI: Mengelola Data Raport, Absensi, & Catatan
// ==========================================

// --- FUNGSI BACKEND ABSENSI WALI KELAS ---
function getSiswaDanAbsen(id_tapel, id_kelas, jenis_ujian) {
  try {
    const sheetSiswa = DB.getSheet('mst_siswa');
    const dataSiswa = sheetSiswa.getDataRange().getDisplayValues();
    const headS = dataSiswa[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let mapSiswa = {};
    for(let i=1; i<dataSiswa.length; i++) {
      mapSiswa[dataSiswa[i][headS.indexOf('id_siswa')]] = dataSiswa[i][headS.indexOf('nama_lengkap')];
    }

    const sheetPlot = DB.getSheet('trx_kelas_siswa');
    let dataPlot = sheetPlot.getDataRange().getDisplayValues();
    let headP = dataPlot[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let siswaDiKelas = [];
    for(let i=1; i<dataPlot.length; i++) {
      if(dataPlot[i][headP.indexOf('id_tapel')] == id_tapel && dataPlot[i][headP.indexOf('id_kelas')] == id_kelas) {
        siswaDiKelas.push(dataPlot[i][headP.indexOf('id_siswa')]);
      }
    }
    if(siswaDiKelas.length === 0) return {status:"error", message:"Belum ada santri."};

    const sheetAbsen = DB.getSheet('trx_absen');
    let allAbsen = sheetAbsen.getDataRange().getDisplayValues();
    if(allAbsen.length === 0 || allAbsen[0].length < 9) {
      sheetAbsen.getRange(1, 1, 1, 9).setValues([['id_absen', 'id_tapel', 'id_kelas', 'jenis_ujian', 'id_siswa', 'sakit', 'izin', 'alpha', 'catatan']]);
      allAbsen = sheetAbsen.getDataRange().getDisplayValues();
    }
    const headA = allAbsen[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let mapAbsen = {};
    for(let i=1; i<allAbsen.length; i++) {
      if(allAbsen[i][headA.indexOf('id_tapel')] == id_tapel && allAbsen[i][headA.indexOf('id_kelas')] == id_kelas && allAbsen[i][headA.indexOf('jenis_ujian')] == jenis_ujian) {
         mapAbsen[allAbsen[i][headA.indexOf('id_siswa')]] = {
           sakit: allAbsen[i][headA.indexOf('sakit')], izin: allAbsen[i][headA.indexOf('izin')],
           alpha: allAbsen[i][headA.indexOf('alpha')], catatan: allAbsen[i][headA.indexOf('catatan')]
         };
      }
    }

    let resultData = [];
    siswaDiKelas.forEach(idS => {
      if(mapSiswa[idS]) {
        let ab = mapAbsen[idS] || {sakit:"", izin:"", alpha:"", catatan:""};
        resultData.push({ id_siswa: idS, nama: mapSiswa[idS], sakit: ab.sakit, izin: ab.izin, alpha: ab.alpha, catatan: ab.catatan });
      }
    });
    resultData.sort((a,b)=>a.nama.localeCompare(b.nama));
    return {status:"success", data: resultData};
  } catch(e) { return {status:"error", message:e.toString()}; }
}

function saveAbsenKolektif(id_tapel, id_kelas, jenis_ujian, arr_absen) {
  try {
    const sheet = DB.getSheet('trx_absen');
    let allData = sheet.getDataRange().getValues();
    if(allData.length === 0 || allData[0].length < 9) {
      sheet.getRange(1,1,1,9).setValues([['id_absen', 'id_tapel', 'id_kelas', 'jenis_ujian', 'id_siswa', 'sakit', 'izin', 'alpha', 'catatan']]);
      allData = sheet.getDataRange().getValues();
    }
    let head = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    
    // Hapus data lama (Wipe)
    for(let i = allData.length - 1; i > 0; i--) {
      if(allData[i][head.indexOf('id_tapel')] == id_tapel && allData[i][head.indexOf('id_kelas')] == id_kelas && allData[i][head.indexOf('jenis_ujian')] == jenis_ujian) {
         sheet.deleteRow(i + 1);
      }
    }

    // Insert data baru
    if(arr_absen && arr_absen.length > 0) {
      let newRows = [];
      arr_absen.forEach(a => {
        if(a.sakit || a.izin || a.alpha || a.catatan) {
          let row = new Array(head.length).fill("");
          row[head.indexOf('id_absen')] = Utilities.getUuid();
          row[head.indexOf('id_tapel')] = id_tapel;
          row[head.indexOf('id_kelas')] = id_kelas;
          row[head.indexOf('jenis_ujian')] = jenis_ujian;
          row[head.indexOf('id_siswa')] = a.id_siswa;
          row[head.indexOf('sakit')] = a.sakit;
          row[head.indexOf('izin')] = a.izin;
          row[head.indexOf('alpha')] = a.alpha;
          row[head.indexOf('catatan')] = a.catatan;
          newRows.push(row);
        }
      });
      if(newRows.length > 0) sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, head.length).setValues(newRows);
    }
    return {status:"success", message:"Kehadiran & Catatan tersimpan!"};
  } catch(e) { return {status:"error", message:e.toString()}; }
}

// --- FUNGSI BACKEND CETAK RAPORT (GABUNGAN SEMUA DATA) ---
function getRaportData(ta, id_kelas, jenis_ujian) {
  try {
    let madrasah = { nama: "NAMA MADRASAH", alamat: "-", nsm: "-", npsn: "-", kepala: "-", nip: "-", logo: "" };
    const sheetSet = DB.getSheet('set_madrasah');
    const dataSet = sheetSet.getDataRange().getDisplayValues();
    if (dataSet.length > 1) {
      const headSet = dataSet[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
      madrasah.nama = dataSet[1][headSet.indexOf('nama_madrasah')] || "";
      madrasah.nsm = dataSet[1][headSet.indexOf('nsm')] || ""; madrasah.npsn = dataSet[1][headSet.indexOf('npsn')] || "";
      madrasah.alamat = (dataSet[1][headSet.indexOf('alamat')] || "") + ", " + (dataSet[1][headSet.indexOf('kabupaten')] || "") + " - " + (dataSet[1][headSet.indexOf('provinsi')] || "");
      madrasah.kepala = dataSet[1][headSet.indexOf('kepala_madrasah')] || ""; madrasah.nip = dataSet[1][headSet.indexOf('nip_kepala')] || "";
      if(headSet.indexOf('logo_url_drive') > -1) madrasah.logo = sheetSet.getRange(2, headSet.indexOf('logo_url_drive') + 1).getValue();
    }

    let kelasName = "-";
    const sheetKelas = DB.getSheet('mst_kelas');
    const dataKelas = sheetKelas.getDataRange().getDisplayValues();
    const headK = dataKelas[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    for(let i=1; i<dataKelas.length; i++) {
      if(dataKelas[i][headK.indexOf('id_kelas')] === id_kelas) {
        kelasName = dataKelas[i][headK.indexOf('nama_kelas')] + " (Tkt " + dataKelas[i][headK.indexOf('tingkat')] + ")"; break;
      }
    }

    const sheetMapel = DB.getSheet('mst_mapel');
    const dataMapel = sheetMapel.getDataRange().getDisplayValues();
    const headM = dataMapel[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let mapelList = [];
    for(let i=1; i<dataMapel.length; i++) {
      if((dataMapel[i][headM.indexOf('deleted_at')] === "" || !dataMapel[i][headM.indexOf('deleted_at')]) && (dataMapel[i][headM.indexOf('is_active')] === 'TRUE' || dataMapel[i][headM.indexOf('is_active')] === true)) {
        mapelList.push({ id: dataMapel[i][headM.indexOf('id_mapel')], nama: dataMapel[i][headM.indexOf('nama_mapel')] });
      }
    }

    const sheetPlot = DB.getSheet('trx_kelas_siswa');
    let dataPlot = sheetPlot.getDataRange().getDisplayValues();
    const headP = dataPlot[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let idSiswaDiKelas = [];
    if(headP.indexOf('id_tapel') > -1) {
      for(let i=1; i<dataPlot.length; i++) {
        if(dataPlot[i][headP.indexOf('id_tapel')] == ta && dataPlot[i][headP.indexOf('id_kelas')] == id_kelas) {
           idSiswaDiKelas.push(dataPlot[i][headP.indexOf('id_siswa')]);
        }
      }
    }

    if(idSiswaDiKelas.length === 0) return { status: "error", message: "Belum ada santri di kelas ini." };

    const sheetSiswa = DB.getSheet('mst_siswa');
    const dataSiswa = sheetSiswa.getDataRange().getDisplayValues();
    const headS = dataSiswa[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let listSiswa = [];
    for(let i=1; i<dataSiswa.length; i++) {
      let idS = dataSiswa[i][headS.indexOf('id_siswa')];
      if(idSiswaDiKelas.indexOf(idS) > -1) {
        listSiswa.push({ id_siswa: idS, nis: dataSiswa[i][headS.indexOf('nis')], nama: dataSiswa[i][headS.indexOf('nama_lengkap')] });
      }
    }
    listSiswa.sort((a,b) => a.nama.localeCompare(b.nama));

    // AMBIL NILAI
    const sheetNilai = DB.getSheet('trx_nilai');
    let allNilai = sheetNilai.getDataRange().getDisplayValues();
    let dbNilai = {}; 
    if(allNilai.length > 1) {
      const headN = allNilai[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
      for(let i=1; i<allNilai.length; i++) {
        if(allNilai[i][headN.indexOf('id_tapel')] == ta && allNilai[i][headN.indexOf('id_kelas')] == id_kelas) {
           let idS = allNilai[i][headN.indexOf('id_siswa')];
           let idM = allNilai[i][headN.indexOf('id_mapel')];
           if(!dbNilai[idS]) dbNilai[idS] = {};
           dbNilai[idS][idM] = jenis_ujian === "UTS" ? allNilai[i][headN.indexOf('nilai_uts')] : allNilai[i][headN.indexOf('nilai_uas')];
        }
      }
    }

    // AMBIL ABSENSI
    const sheetAbsen = DB.getSheet('trx_absen');
    let allAbsen = sheetAbsen.getDataRange().getDisplayValues();
    let dbAbsen = {};
    if(allAbsen.length > 1) {
      const headA = allAbsen[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
      for(let i=1; i<allAbsen.length; i++) {
        if(allAbsen[i][headA.indexOf('id_tapel')] == ta && allAbsen[i][headA.indexOf('id_kelas')] == id_kelas && allAbsen[i][headA.indexOf('jenis_ujian')] == jenis_ujian) {
           dbAbsen[allAbsen[i][headA.indexOf('id_siswa')]] = {
             sakit: allAbsen[i][headA.indexOf('sakit')], izin: allAbsen[i][headA.indexOf('izin')],
             alpha: allAbsen[i][headA.indexOf('alpha')], catatan: allAbsen[i][headA.indexOf('catatan')]
           };
        }
      }
    }

    let finalData = [];
    listSiswa.forEach(s => {
      let nilaiMapelSiswa = [];
      mapelList.forEach(m => {
        let val = (dbNilai[s.id_siswa] && dbNilai[s.id_siswa][m.id]) ? dbNilai[s.id_siswa][m.id] : "";
        nilaiMapelSiswa.push({ nama_mapel: m.nama, nilai: val });
      });
      let absenCatatan = dbAbsen[s.id_siswa] || {sakit:"", izin:"", alpha:"", catatan:""};
      finalData.push({ info: s, nilai: nilaiMapelSiswa, absen: absenCatatan });
    });

    return { status: "success", madrasah: madrasah, kelas: kelasName, ta: ta, jenis_ujian: jenis_ujian, data_raport: finalData };
  } catch(e) { return { status: "error", message: e.toString() }; }
}
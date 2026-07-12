// ==========================================
// FILE: Kelas.js
// ==========================================
function getKelasInit() {
  try {
    // 1. Ambil Data Kelas
    const sheetKelas = DB.getSheet('mst_kelas');
    const allKelas = sheetKelas.getDataRange().getDisplayValues(); 
    const headersK = allKelas.length > 0 ? allKelas[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_')) : [];
    
    let dataKelas = [];
    for (let i = 1; i < allKelas.length; i++) {
      let obj = {};
      headersK.forEach((h, index) => { obj[h] = allKelas[i][index]; });
      if (obj.deleted_at === "" || obj.deleted_at === null || obj.deleted_at === undefined) {
        dataKelas.push(obj);
      }
    }
    // 2. Ambil Data Guru (Hanya yang Aktif)
    const sheetGuru = DB.getSheet('mst_guru');
    const allGuru = sheetGuru.getDataRange().getDisplayValues();
    const headersG = allGuru.length > 0 ? allGuru[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_')) : [];
    
    let dataGuru = [];
    for (let i = 1; i < allGuru.length; i++) {
      let obj = {};
      headersG.forEach((h, index) => { obj[h] = allGuru[i][index]; });
      // Filter guru yang belum dihapus DAN berstatus aktif
      if ((obj.deleted_at === "" || obj.deleted_at === undefined) && (obj.is_active === 'TRUE' || obj.is_active === true || obj.is_active === 'true')) {
        dataGuru.push({ id_guru: obj.id_guru, nama_lengkap: obj.nama_lengkap });
      }
    }
    // Kirim kedua data ke Frontend
    return { status: "success", dataKelas: dataKelas, dataGuru: dataGuru };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function saveKelas(dataObj) {
  try {
    const sheet = DB.getSheet('mst_kelas');
    const allData = sheet.getDataRange().getValues();
    if (allData.length === 0) return { status: 'error', message: 'Sheet mst_kelas kosong.' };
    
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_kelas');
    const idxNama = headers.indexOf('nama_kelas');
    const idxTingkat = headers.indexOf('tingkat');
    const idxWali = headers.indexOf('wali_kelas');
    const idxActive = headers.indexOf('is_active');
    let idxDeleted = headers.indexOf('deleted_at');
    
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
      headers.push('deleted_at');
    }
    if (idxId === -1 || idxNama === -1) return { status: 'error', message: 'Header id_kelas atau nama_kelas tidak ditemukan!' };
    const isActive = (dataObj.is_active === 'true' || dataObj.is_active === true);
    
    if (dataObj.id_kelas) {
      let rowIndex = -1;
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][idxId] == dataObj.id_kelas) { rowIndex = i + 1; break; }
      }
      
      if (rowIndex > -1) {
        if(idxNama > -1) sheet.getRange(rowIndex, idxNama + 1).setValue(dataObj.nama_kelas);
        if(idxTingkat > -1) sheet.getRange(rowIndex, idxTingkat + 1).setValue(dataObj.tingkat);
        if(idxWali > -1) sheet.getRange(rowIndex, idxWali + 1).setValue(dataObj.wali_kelas); // Menyimpan ID Guru
        if(idxActive > -1) sheet.getRange(rowIndex, idxActive + 1).setValue(isActive);
        return { status: 'success', message: 'Data Kelas diperbarui!' };
      }
      return { status: 'error', message: 'ID Kelas tidak ditemukan!' };
    } else {
      const isNamaExist = allData.some((row, idx) => idx > 0 && row[idxNama].toString().toLowerCase() == dataObj.nama_kelas.toLowerCase() && row[idxDeleted] === "");
      if(isNamaExist) return { status: 'error', message: 'Nama Kelas sudah digunakan!' };
      let newRow = new Array(headers.length).fill("");
      newRow[idxId] = Utilities.getUuid();
      if(idxNama > -1) newRow[idxNama] = dataObj.nama_kelas;
      if(idxTingkat > -1) newRow[idxTingkat] = dataObj.tingkat;
      if(idxWali > -1) newRow[idxWali] = dataObj.wali_kelas; // Menyimpan ID Guru
      if(idxActive > -1) newRow[idxActive] = isActive;
      newRow[idxDeleted] = ""; 
      
      sheet.appendRow(newRow);
      return { status: 'success', message: 'Kelas baru ditambahkan!' };
    }
  } catch(e) { return { status: 'error', message: e.toString() }; }
}

function deleteKelas(id_kelas) {
  try {
    const sheet = DB.getSheet('mst_kelas');
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_kelas');
    let idxDeleted = headers.indexOf('deleted_at');
    
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
    }
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idxId] == id_kelas) {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(i + 1, idxDeleted + 1).setValue(timestamp);
        return { status: 'success', message: 'Data Kelas dihapus!' };
      }
    }
    return { status: 'error', message: 'ID Kelas tidak ditemukan!' };
  } catch(e) { return { status: 'error', message: e.toString() }; }
}
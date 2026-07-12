// ==========================================
// FILE: Mapel.js
// FUNGSI: Modul CRUD Master Data Mata Pelajaran
// ==========================================
function getMapelList() {
  try {
    const sheet = DB.getSheet('mst_mapel');
    const allData = sheet.getDataRange().getDisplayValues(); 
    
    if (allData.length <= 1) return { status: "success", data: [] };
    
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    
    let result = [];
    for (let i = 1; i < allData.length; i++) {
      let row = allData[i];
      let obj = {};
      headers.forEach((h, index) => { obj[h] = row[index]; });
      
      if (obj.deleted_at === "" || obj.deleted_at === null || obj.deleted_at === undefined) {
        result.push(obj);
      }
    }
    return { status: "success", data: result };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function saveMapel(dataObj) {
  try {
    const sheet = DB.getSheet('mst_mapel');
    const allData = sheet.getDataRange().getValues();
    if (allData.length === 0) return { status: 'error', message: 'Sheet mst_mapel kosong.' };
    
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_mapel');
    const idxKode = headers.indexOf('kode_mapel');
    const idxNama = headers.indexOf('nama_mapel');
    const idxKkm = headers.indexOf('kkm');
    const idxActive = headers.indexOf('is_active');
    let idxDeleted = headers.indexOf('deleted_at');
    
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
      headers.push('deleted_at');
    }
    if (idxId === -1 || idxKode === -1 || idxNama === -1) {
      return { status: 'error', message: 'Header id_mapel / kode_mapel / nama_mapel tidak ditemukan!' };
    }
    const isActive = (dataObj.is_active === 'true' || dataObj.is_active === true);
    
    if (dataObj.id_mapel) {
      // PROSES EDIT
      let rowIndex = -1;
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][idxId] == dataObj.id_mapel) { rowIndex = i + 1; break; }
      }
      
      if (rowIndex > -1) {
        // Cek apakah kode mapel dipakai oleh mapel lain
        const isKodeExist = allData.some((row, idx) => idx !== (rowIndex-1) && row[idxKode].toString().toLowerCase() == dataObj.kode_mapel.toLowerCase() && row[idxDeleted] === "");
        if(isKodeExist) return { status: 'error', message: 'Kode Mapel sudah digunakan!' };
        if(idxKode > -1) sheet.getRange(rowIndex, idxKode + 1).setValue(dataObj.kode_mapel);
        if(idxNama > -1) sheet.getRange(rowIndex, idxNama + 1).setValue(dataObj.nama_mapel);
        if(idxKkm > -1) sheet.getRange(rowIndex, idxKkm + 1).setValue(dataObj.kkm);
        if(idxActive > -1) sheet.getRange(rowIndex, idxActive + 1).setValue(isActive);
        return { status: 'success', message: 'Data Mapel berhasil diperbarui!' };
      }
      return { status: 'error', message: 'ID Mapel tidak ditemukan!' };
      
    } else {
      // PROSES TAMBAH
      const isKodeExist = allData.some((row, idx) => idx > 0 && row[idxKode].toString().toLowerCase() == dataObj.kode_mapel.toLowerCase() && row[idxDeleted] === "");
      if(isKodeExist) return { status: 'error', message: 'Kode Mapel sudah digunakan!' };
      let newRow = new Array(headers.length).fill("");
      newRow[idxId] = Utilities.getUuid();
      if(idxKode > -1) newRow[idxKode] = dataObj.kode_mapel;
      if(idxNama > -1) newRow[idxNama] = dataObj.nama_mapel;
      if(idxKkm > -1) newRow[idxKkm] = dataObj.kkm;
      if(idxActive > -1) newRow[idxActive] = isActive;
      newRow[idxDeleted] = ""; 
      
      sheet.appendRow(newRow);
      return { status: 'success', message: 'Mapel baru berhasil ditambahkan!' };
    }
  } catch(e) { return { status: 'error', message: e.toString() }; }
}

function deleteMapel(id_mapel) {
  try {
    const sheet = DB.getSheet('mst_mapel');
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_mapel');
    let idxDeleted = headers.indexOf('deleted_at');
    
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
    }
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idxId] == id_mapel) {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(i + 1, idxDeleted + 1).setValue(timestamp);
        return { status: 'success', message: 'Data Mapel berhasil dihapus!' };
      }
    }
    return { status: 'error', message: 'Gagal hapus, ID Mapel tidak ditemukan!' };
  } catch(e) { return { status: 'error', message: e.toString() }; }
}
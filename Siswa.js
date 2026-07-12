/**
 * FILE: Siswa.js
 * FUNGSI: Modul CRUD Master Data Siswa
 */

function getSiswaList() {
  try {
    const sheet = DB.getSheet('mst_siswa');
    const allData = sheet.getDataRange().getDisplayValues(); 
    
    if (allData.length <= 1) return { status: "success", data: [] };
    
    // Normalisasi Header
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

function saveSiswa(dataObj) {
  try {
    const sheet = DB.getSheet('mst_siswa');
    const allData = sheet.getDataRange().getValues();
    if (allData.length === 0) return { status: 'error', message: 'Sheet mst_siswa kosong.' };
    
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_siswa');
    const idxNis = headers.indexOf('nis');
    const idxNama = headers.indexOf('nama_lengkap');
    const idxJk = headers.indexOf('jk');
    const idxTempat = headers.indexOf('tempat_lahir');
    const idxTgl = headers.indexOf('tgl_lahir');
    const idxOrtu = headers.indexOf('nama_ortu');
    const idxHp = headers.indexOf('no_hp');
    const idxActive = headers.indexOf('is_active');
    let idxDeleted = headers.indexOf('deleted_at');
    
    // Auto-Repair Kolom deleted_at
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
      headers.push('deleted_at');
    }

    if (idxId === -1 || idxNis === -1) return { status: 'error', message: 'Header id_siswa atau nis tidak ditemukan!' };

    const isActive = (dataObj.is_active === 'true' || dataObj.is_active === true);
    
    if (dataObj.id_siswa) {
      // PROSES EDIT
      let rowIndex = -1;
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][idxId] == dataObj.id_siswa) { rowIndex = i + 1; break; }
      }
      
      if (rowIndex > -1) {
        if(idxNis > -1) sheet.getRange(rowIndex, idxNis + 1).setValue(dataObj.nis);
        if(idxNama > -1) sheet.getRange(rowIndex, idxNama + 1).setValue(dataObj.nama_lengkap);
        if(idxJk > -1) sheet.getRange(rowIndex, idxJk + 1).setValue(dataObj.jk);
        if(idxTempat > -1) sheet.getRange(rowIndex, idxTempat + 1).setValue(dataObj.tempat_lahir);
        if(idxTgl > -1) sheet.getRange(rowIndex, idxTgl + 1).setValue(dataObj.tgl_lahir);
        if(idxOrtu > -1) sheet.getRange(rowIndex, idxOrtu + 1).setValue(dataObj.nama_ortu);
        if(idxHp > -1) sheet.getRange(rowIndex, idxHp + 1).setValue(dataObj.no_hp);
        if(idxActive > -1) sheet.getRange(rowIndex, idxActive + 1).setValue(isActive);
        return { status: 'success', message: 'Data Siswa berhasil diperbarui!' };
      }
      return { status: 'error', message: 'ID Siswa tidak ditemukan!' };
    } else {
      // PROSES TAMBAH
      const isNisExist = allData.some((row, idx) => idx > 0 && row[idxNis] == dataObj.nis && row[idxDeleted] === "");
      if(isNisExist) return { status: 'error', message: 'NIS sudah terdaftar!' };

      let newRow = new Array(headers.length).fill("");
      newRow[idxId] = Utilities.getUuid();
      if(idxNis > -1) newRow[idxNis] = dataObj.nis;
      if(idxNama > -1) newRow[idxNama] = dataObj.nama_lengkap;
      if(idxJk > -1) newRow[idxJk] = dataObj.jk;
      if(idxTempat > -1) newRow[idxTempat] = dataObj.tempat_lahir;
      if(idxTgl > -1) newRow[idxTgl] = dataObj.tgl_lahir;
      if(idxOrtu > -1) newRow[idxOrtu] = dataObj.nama_ortu;
      if(idxHp > -1) newRow[idxHp] = dataObj.no_hp;
      if(idxActive > -1) newRow[idxActive] = isActive;
      newRow[idxDeleted] = ""; 
      
      sheet.appendRow(newRow);
      return { status: 'success', message: 'Siswa baru berhasil ditambahkan!' };
    }
  } catch(e) { return { status: 'error', message: e.toString() }; }
}

function deleteSiswa(id_siswa) {
  try {
    const sheet = DB.getSheet('mst_siswa');
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_siswa');
    let idxDeleted = headers.indexOf('deleted_at');
    
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
    }
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idxId] == id_siswa) {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(i + 1, idxDeleted + 1).setValue(timestamp);
        return { status: 'success', message: 'Data Siswa berhasil dihapus!' };
      }
    }
    return { status: 'error', message: 'Gagal hapus, ID Siswa tidak ditemukan!' };
  } catch(e) { return { status: 'error', message: e.toString() }; }
}
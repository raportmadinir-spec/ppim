/**
 * FILE: Guru.js
 * FUNGSI: Mengelola data master Guru (CRUD)
 */

function getGuruList() {
  try {
    const sheet = DB.getSheet('mst_guru');
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

function saveGuru(dataObj) {
  try {
    const sheet = DB.getSheet('mst_guru');
    const allData = sheet.getDataRange().getValues();
    if (allData.length === 0) return { status: 'error', message: 'Sheet kosong.' };
    
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_guru');
    const idxNip = headers.indexOf('nip');
    const idxNama = headers.indexOf('nama_lengkap');
    const idxJk = headers.indexOf('jk');
    const idxHp = headers.indexOf('no_hp');
    const idxActive = headers.indexOf('is_active');
    let idxDeleted = headers.indexOf('deleted_at');
    
    // Auto-Repair Kolom deleted_at jika hilang
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
      headers.push('deleted_at');
    }

    if (idxId === -1 || idxNip === -1) return { status: 'error', message: 'Header id_guru atau nip tidak ditemukan di baris 1.' };

    const isActive = (dataObj.is_active === 'true' || dataObj.is_active === true);
    
    if (dataObj.id_guru) {
      let rowIndex = -1;
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][idxId] == dataObj.id_guru) { rowIndex = i + 1; break; }
      }
      if (rowIndex > -1) {
        sheet.getRange(rowIndex, idxNip + 1).setValue(dataObj.nip);
        sheet.getRange(rowIndex, idxNama + 1).setValue(dataObj.nama_lengkap);
        sheet.getRange(rowIndex, idxJk + 1).setValue(dataObj.jk);
        sheet.getRange(rowIndex, idxHp + 1).setValue(dataObj.no_hp);
        sheet.getRange(rowIndex, idxActive + 1).setValue(isActive);
        return { status: 'success', message: 'Data guru diperbarui!' };
      }
      return { status: 'error', message: 'ID tidak ditemukan!' };
    } else {
      const isNipExist = allData.some((row, idx) => idx > 0 && row[idxNip] == dataObj.nip && row[idxDeleted] === "");
      if(isNipExist) return { status: 'error', message: 'NIP sudah terdaftar!' };

      let newRow = new Array(headers.length).fill("");
      // GENERATE ID OTOMATIS
      newRow[idxId] = Utilities.getUuid();
      newRow[idxNip] = dataObj.nip;
      newRow[idxNama] = dataObj.nama_lengkap;
      newRow[idxJk] = dataObj.jk;
      newRow[idxHp] = dataObj.no_hp;
      newRow[idxActive] = isActive;
      newRow[idxDeleted] = ""; 
      sheet.appendRow(newRow);
      return { status: 'success', message: 'Guru baru ditambahkan!' };
    }
  } catch(e) { return { status: 'error', message: e.toString() }; }
}

function deleteGuru(id_guru) {
  try {
    const sheet = DB.getSheet('mst_guru');
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    const idxId = headers.indexOf('id_guru');
    let idxDeleted = headers.indexOf('deleted_at');
    
    // Auto-Repair Kolom deleted_at jika hilang
    if (idxDeleted === -1) {
      idxDeleted = headers.length;
      sheet.getRange(1, idxDeleted + 1).setValue('deleted_at');
    }
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idxId] == id_guru) {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(i + 1, idxDeleted + 1).setValue(timestamp);
        return { status: 'success', message: 'Data berhasil dihapus dari sistem!' };
      }
    }
    return { status: 'error', message: 'Gagal hapus, ID Guru tidak ditemukan!' };
  } catch(e) { return { status: 'error', message: e.toString() }; }
}
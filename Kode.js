/**
 * FILE: Kode.js
 * Deskripsi: File utama untuk konfigurasi aplikasi, routing halaman, dan fungsi autentikasi.
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('SIM RAPORT MADIN')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loadPage(pageName) {
  try { 
    return HtmlService.createHtmlOutputFromFile(pageName).getContent(); 
  } catch (e) { 
    return "<h2>Halaman tidak ditemukan!</h2>"; 
  }
}

// LOGIN DATABASE ASLI
function prosesLogin(username, password) {
  try {
    // 1. Cek Admin/Kepala di mst_users
    const dataUsers = DB.getAll('mst_users');
    const user = dataUsers.find(u => 
      u.username == username && 
      u.password_hash == password && 
      (u.is_active === true || u.is_active === "TRUE")
    );

    if (user) {
      return { 
        status: "success", 
        role: user.role, 
        name: user.username, 
        message: "Login Berhasil!" 
      };
    }

    // 2. Cek Guru di mst_guru (Username = NIP)
    const dataGuru = DB.getAll('mst_guru');
    const guru = dataGuru.find(g => 
      g.nip == username && 
      (g.is_active === true || g.is_active === "TRUE") && 
      g.deleted_at === ""
    );

    if (guru) {
      // Password Guru menggunakan NIP (karena tidak ada kolom password di mst_guru)
      if (password == guru.nip) {
         return { 
           status: "success", 
           role: "Guru", 
           name: guru.nama_lengkap, 
           nip: guru.nip, 
           message: "Login Berhasil!" 
         };
      } else {
         return { status: "error", message: "Password Guru salah!" };
      }
    }

    return { status: "error", message: "Username tidak ditemukan atau dinonaktifkan!" };
  } catch(e) {
    return { status: "error", message: "Error DB: " + e.toString() };
  }
}

function getDashboardStats() {
  return {
    status: "success",
    data: {
      guru: DB.getAll('mst_guru').filter(g => g.deleted_at === "").length,
      siswa: DB.getAll('mst_siswa').filter(s => s.deleted_at === "").length,
      kelas: DB.getAll('mst_kelas').filter(k => k.deleted_at === "").length,
      mapel: DB.getAll('mst_mapel').filter(m => m.deleted_at === "").length
    }
  };
}
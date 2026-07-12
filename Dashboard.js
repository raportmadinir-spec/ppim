/**
 * FILE: Dashboard.js
 * FUNGSI: Mengolah data statistik Dashboard
 */

function getDashboardStats() {
  try {
    // Membaca data dari database melalui modul DB
    const dataGuru = DB.getAll('mst_guru');
    const dataSiswa = DB.getAll('mst_siswa');
    const dataKelas = DB.getAll('mst_kelas');
    const dataMapel = DB.getAll('mst_mapel');

    // Menghitung jumlah data aktif (mengabaikan yang di-soft delete)
    const totalGuru = dataGuru.filter(g => g.deleted_at === "").length;
    const totalSiswa = dataSiswa.filter(s => s.deleted_at === "" && s.status_siswa !== "Keluar").length;
    const totalKelas = dataKelas.filter(k => k.deleted_at === "").length;
    const totalMapel = dataMapel.filter(m => m.deleted_at === "").length;

    return {
      status: "success",
      data: {
        guru: totalGuru,
        siswa: totalSiswa,
        kelas: totalKelas,
        mapel: totalMapel
      }
    };
  } catch (error) {
    return {
      status: "error",
      message: error.toString()
    };
  }
}
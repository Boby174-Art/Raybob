// ID Google Sheets yang digunakan sebagai database
const SPREADSHEET_ID = "1c-O1gKKXzVJSGJaKrxXz-ZyDuE7PgLQDMoeFnMelxDM";

/**
 * Fungsi utama untuk melayani web app (GET Request)
 */
function doGet() {
  // Inisialisasi sheet database jika belum ada
  inisialisasiDatabase();
  
  // Render halaman index.html
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('E-Fisio: Rekam Medis & Jadwal Fisioterapi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mengambil sheet database berdasarkan ID
 */
function getDb() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error("Gagal membuka spreadsheet. Pastikan ID Spreadsheet benar dan izin akses telah diberikan. Error: " + e.message);
  }
}

/**
 * Membuat tab/sheet default jika belum ada di Google Sheets
 */
function inisialisasiDatabase() {
  const ss = getDb();
  
  // 1. Sheet Pasien
  if (!ss.getSheetByName("Pasien")) {
    const sheet = ss.insertSheet("Pasien");
    sheet.appendRow([
      "ID_Pasien", "Nama", "NIK", "JK", "Gol_Darah", 
      "Tempat_Tanggal_Lahir", "Alamat", "No_WA", "Pekerjaan", 
      "Jenis_Administrasi", "Status_Terapi"
    ]);
    // Data dummy awal
    sheet.appendRow([
      "PS001", "Budi Santoso", "3171012345670001", "L", "O", 
      "Jakarta, 12-05-1988", "Jl. Merdeka No. 10", "08123456789", "Karyawan Swasta", 
      "BPJS", "Aktif"
    ]);
  }
  
  // 2. Sheet RekamMedis
  if (!ss.getSheetByName("RekamMedis")) {
    const sheet = ss.insertSheet("RekamMedis");
    sheet.appendRow([
      "ID_Rekam", "ID_Pasien", "Tanggal", "Keluhan", 
      "Diagnosa", "Tindakan", "Catatan_Fisio", "ID_Fisioterapis"
    ]);
    // Data dummy awal
    sheet.appendRow([
      "RM001", "PS001", "2026-06-20", "Nyeri pada lutut kanan setelah lari pagi", 
      "Mild Osteoarthritis Genu Dextra", "US, TENS, dan Quadriceps Strengthening Exercise", 
      "Pasien kooperatif, nyeri berkurang dari skala 6 menjadi 4", "Rayfaldi Boby Saputra, M.Fis"
    ]);
  }
  
  // 3. Sheet Jadwal
  if (!ss.getSheetByName("Jadwal")) {
    const sheet = ss.insertSheet("Jadwal");
    sheet.appendRow([
      "ID_Jadwal", "ID_Pasien", "Tanggal_Terapi", "Jam", "Status", 
      "Usulan_Tanggal_Baru", "Usulan_Jam_Baru"
    ]);
    // Data dummy awal
    sheet.appendRow([
      "JW001", "PS001", "2026-06-25", "10:00", "Dijadwalkan", "", ""
    ]);
  }
  
  // 4. Sheet Users
  if (!ss.getSheetByName("Users")) {
    const sheet = ss.insertSheet("Users");
    sheet.appendRow(["Username", "Password", "Role"]);
    // Akun default (Pastikan untuk mengganti password di lingkungan production!)
    sheet.appendRow(["admin", "admin123", "Admin"]);
    sheet.appendRow(["terapis", "terapis123", "Terapis"]);
  }
  
  // 5. Sheet Pengaturan_Beranda
  if (!ss.getSheetByName("Pengaturan_Beranda")) {
    const sheet = ss.insertSheet("Pengaturan_Beranda");
    sheet.appendRow(["Key", "Value"]);
    sheet.appendRow(["Nama_Fisioterapis", "Rayfaldi Boby Saputra, M.Fis"]);
    sheet.appendRow(["Kontak_WA", "0895-4022-16058"]);
    sheet.appendRow(["Link_Maps", "https://maps.app.goo.gl/PFQ6M2wqaCw5cZgh7"]);
    sheet.appendRow(["Deskripsi_Sistem", "Layanan fisioterapi profesional untuk memulihkan fungsi fisik, mengelola nyeri, dan meningkatkan kualitas hidup Anda dengan pendekatan klinis terpercaya."]);
    sheet.appendRow(["Informasi_Klinik", "Jam Operasional: Senin - Sabtu (08:00 - 17:00). Silakan hubungi admin melalui kontak WhatsApp jika memiliki pertanyaan mendesak."]);
  }
}

/**
 * Mengubah array 2D dari sheet menjadi array of objects agar mudah diolah di frontend
 */
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      let val = row[i];
      // Format tanggal ke string YYYY-MM-DD agar aman diparsing JS/HTML inputs
      if (val instanceof Date) {
        obj[header] = formatTanggalString(val);
      } else {
        obj[header] = val;
      }
    });
    return obj;
  });
}

/**
 * Format objek Date menjadi string "YYYY-MM-DD"
 */
function formatTanggalString(date) {
  const d = new Date(date);
  let month = '' + (d.getMonth() + 1);
  let day = '' + d.getDate();
  const year = d.getFullYear();

  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;

  return [year, month, day].join('-');
}

/**
 * Fungsi pembantu untuk menulis baris baru dengan ID auto-increment
 */
function generateNextId(sheet, prefix) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return prefix + "001";
  
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const idStr = String(data[i][0]);
    const num = parseInt(idStr.replace(prefix, ""), 10);
    if (!isNaN(num) && num > maxNum) {
      maxNum = num;
    }
  }
  const nextNum = maxNum + 1;
  return prefix + String(nextNum).padStart(3, '0');
}


// ==========================================
// API BACKEND (DAPAT DIPANGGIL DARI FRONTEND)
// ==========================================

/**
 * Ambil data konfigurasi beranda publik
 */
function dapatkanPengaturanBeranda() {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName("Pengaturan_Beranda");
    const rows = sheet.getDataRange().getValues();
    const config = {};
    for (let i = 1; i < rows.length; i++) {
      config[rows[i][0]] = rows[i][1];
    }
    return { success: true, data: config };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Memperbarui pengaturan beranda oleh Admin
 */
function simpanPengaturanBeranda(configData) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName("Pengaturan_Beranda");
    const rows = sheet.getDataRange().getValues();
    
    // Looping data baru dan update di Sheet
    for (const key in configData) {
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(configData[key]);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, configData[key]]);
      }
    }
    return { success: true, message: "Pengaturan beranda berhasil diperbarui!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Autentikasi Login Admin & Terapis
 */
function prosesLogin(username, password) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName("Users");
    const users = sheetToObjects(sheet);
    
    const user = users.find(u => u.Username.toLowerCase() === username.toLowerCase() && String(u.Password) === String(password));
    if (user) {
      return { success: true, role: user.Role, username: user.Username };
    } else {
      return { success: false, message: "Username atau Password salah!" };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Fitur Cek Pasien Tanpa Login menggunakan ID_Pasien
 */
function dapatkanDataPasienPublik(idPasien) {
  try {
    const ss = getDb();
    
    // Cari data Pasien
    const sheetPasien = ss.getSheetByName("Pasien");
    const listPasien = sheetToObjects(sheetPasien);
    const pasien = listPasien.find(p => p.ID_Pasien.toUpperCase() === idPasien.toUpperCase());
    
    if (!pasien) {
      return { success: false, message: "ID Pasien tidak ditemukan!" };
    }
    
    // Cari data Rekam Medis (Filter berdasarkan ID_Pasien)
    const sheetRM = ss.getSheetByName("RekamMedis");
    const listRM = sheetToObjects(sheetRM);
    const rekamMedis = listRM
      .filter(r => r.ID_Pasien.toUpperCase() === idPasien.toUpperCase())
      .sort((a, b) => new Date(b.Tanggal) - new Date(a.Tanggal)); // Urutkan dari yang terbaru
      
    // Cari Jadwal Terapi Berikutnya (Filter status aktif / pengajuan reschedule)
    const sheetJadwal = ss.getSheetByName("Jadwal");
    const listJadwal = sheetToObjects(sheetJadwal);
    const jadwal = listJadwal
      .filter(j => j.ID_Pasien.toUpperCase() === idPasien.toUpperCase())
      .sort((a, b) => new Date(b.Tanggal_Terapi) - new Date(a.Tanggal_Terapi))[0] || null; // Ambil jadwal terbaru/terakhir yang diinput
      
    return {
      success: true,
      pasien: pasien,
      rekamMedis: rekamMedis,
      jadwal: jadwal
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Pasien mengajukan perubahan tanggal periksa/terapi
 */
function ajukanReschedulePasien(idPasien, idJadwal, tanggalBaru, jamBaru) {
  try {
    const ss = getDb();
    const sheetJadwal = ss.getSheetByName("Jadwal");
    const listJadwal = sheetToObjects(sheetJadwal);
    
    // Cari baris jadwal yang sesuai
    let rowIndex = -1;
    for (let i = 0; i < listJadwal.length; i++) {
      if (listJadwal[i].ID_Jadwal === idJadwal && listJadwal[i].ID_Pasien === idPasien) {
        rowIndex = i + 2; // +2 karena index array mulai dari 0 dan ada baris header di spreadsheet
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Data jadwal tidak ditemukan!" };
    }
    
    // Update data usulan reschedule di spreadsheet
    sheetJadwal.getRange(rowIndex, 5).setValue("Pengajuan Reschedule"); // Status
    sheetJadwal.getRange(rowIndex, 6).setValue(tanggalBaru);            // Usulan_Tanggal_Baru
    sheetJadwal.getRange(rowIndex, 7).setValue(jamBaru);                // Usulan_Jam_Baru
    
    return { success: true, message: "Pengajuan perubahan jadwal berhasil dikirim ke terapis!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Mendapatkan seluruh list pasien (Untuk Admin & Terapis)
 */
function dapatkanSemuaPasien() {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName("Pasien");
    const data = sheetToObjects(sheet);
    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menambahkan pasien baru oleh Admin
 */
function tambahPasienBaru(dataPasien) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName("Pasien");
    const newId = generateNextId(sheet, "PS");
    
    sheet.appendRow([
      newId,
      dataPasien.Nama,
      dataPasien.NIK,
      dataPasien.JK,
      dataPasien.Gol_Darah,
      dataPasien.Tempat_Tanggal_Lahir,
      dataPasien.Alamat,
      dataPasien.No_WA,
      dataPasien.Pekerjaan,
      dataPasien.Jenis_Administrasi,
      "Aktif" // Status_Terapi default saat mendaftar
    ]);
    
    return { success: true, message: "Pasien baru berhasil didaftarkan dengan ID: " + newId, newId: newId };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menambahkan user baru oleh Admin
 */
function tambahUserBaru(username, password, role) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName("Users");
    const users = sheetToObjects(sheet);
    
    // Cek duplikasi username
    const exists = users.some(u => u.Username.toLowerCase() === username.toLowerCase());
    if (exists) {
      return { success: false, message: "Username sudah terdaftar!" };
    }
    
    sheet.appendRow([username, password, role]);
    return { success: true, message: "User akun baru berhasil dibuat!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Mendapatkan semua data rekam medis pasien tertentu (Untuk Terapis)
 */
function dapatkanRekamMedisPasien(idPasien) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName("RekamMedis");
    const allRM = sheetToObjects(sheet);
    const filtered = allRM.filter(r => r.ID_Pasien === idPasien);
    return { success: true, data: filtered };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menginput Rekam Medis Baru sekaligus menentukan/membuat Jadwal Terapi Baru (Oleh Terapis)
 */
function simpanSesiFisioterapi(idPasien, rmData, jadwalData) {
  try {
    const ss = getDb();
    
    // 1. Simpan Rekam Medis
    const sheetRM = ss.getSheetByName("RekamMedis");
    const newRmId = generateNextId(sheetRM, "RM");
    sheetRM.appendRow([
      newRmId,
      idPasien,
      rmData.Tanggal,
      rmData.Keluhan,
      rmData.Diagnosa,
      rmData.Tindakan,
      rmData.Catatan_Fisio,
      rmData.ID_Fisioterapis
    ]);
    
    // 2. Tandai jadwal yang lalu menjadi "Selesai Sesi" jika ada jadwal aktif sebelumnya
    const sheetJadwal = ss.getSheetByName("Jadwal");
    const listJadwal = sheetToObjects(sheetJadwal);
    for (let i = 0; i < listJadwal.length; i++) {
      if (listJadwal[i].ID_Pasien === idPasien && listJadwal[i].Status !== "Selesai Sesi") {
        sheetJadwal.getRange(i + 2, 5).setValue("Selesai Sesi");
      }
    }
    
    // 3. Buat Jadwal Terapi Berikutnya yang baru
    const newJadwalId = generateNextId(sheetJadwal, "JW");
    sheetJadwal.appendRow([
      newJadwalId,
      idPasien,
      jadwalData.Tanggal_Terapi,
      jadwalData.Jam,
      "Dijadwalkan",
      "", // Usulan_Tanggal_Baru kosong
      ""  // Usulan_Jam_Baru kosong
    ]);
    
    return { success: true, message: "Rekam medis dan jadwal terapi berikutnya berhasil disimpan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Mengambil semua jadwal berstatus 'Pengajuan Reschedule' (Untuk dashboard Terapis)
 */
function dapatkanPengajuanReschedule() {
  try {
    const ss = getDb();
    const sheetJadwal = ss.getSheetByName("Jadwal");
    const listJadwal = sheetToObjects(sheetJadwal);
    
    const sheetPasien = ss.getSheetByName("Pasien");
    const listPasien = sheetToObjects(sheetPasien);
    
    // Filter yang berstatus Pengajuan Reschedule
    const pengajuan = listJadwal.filter(j => j.Status === "Pengajuan Reschedule");
    
    // Map untuk menggabungkan nama pasien demi kemudahan UI terapis
    const result = pengajuan.map(p => {
      const pasien = listPasien.find(pas => pas.ID_Pasien === p.ID_Pasien);
      return {
        ...p,
        Nama_Pasien: pasien ? pasien.Nama : "Pasien Tidak Dikenal"
      };
    });
    
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menyetujui atau Menolak pengajuan reschedule pasien (Oleh Terapis)
 */
function prosesPersetujuanJadwal(idJadwal, tindakan, tanggalAlternatif, jamAlternatif) {
  try {
    const ss = getDb();
    const sheetJadwal = ss.getSheetByName("Jadwal");
    const listJadwal = sheetToObjects(sheetJadwal);
    
    let rowIndex = -1;
    let jadwal = null;
    for (let i = 0; i < listJadwal.length; i++) {
      if (listJadwal[i].ID_Jadwal === idJadwal) {
        rowIndex = i + 2;
        jadwal = listJadwal[i];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Data jadwal tidak ditemukan!" };
    }
    
    if (tindakan === "Terima") {
      // Salin tanggal usulan baru ke kolom tanggal aktif
      sheetJadwal.getRange(rowIndex, 3).setValue(jadwal.Usulan_Tanggal_Baru);
      sheetJadwal.getRange(rowIndex, 4).setValue(jadwal.Usulan_Jam_Baru);
      sheetJadwal.getRange(rowIndex, 5).setValue("Dijadwalkan");
      // Kosongkan kolom usulan
      sheetJadwal.getRange(rowIndex, 6).setValue("");
      sheetJadwal.getRange(rowIndex, 7).setValue("");
      
      return { success: true, message: "Pengajuan perubahan tanggal disetujui!" };
    } else {
      // Jika ditolak, bisa mengembalikan ke tanggal lama atau menetapkan tanggal alternatif
      if (tanggalAlternatif && jamAlternatif) {
        sheetJadwal.getRange(rowIndex, 3).setValue(tanggalAlternatif);
        sheetJadwal.getRange(rowIndex, 4).setValue(jamAlternatif);
      }
      sheetJadwal.getRange(rowIndex, 5).setValue("Dijadwalkan");
      sheetJadwal.getRange(rowIndex, 6).setValue("");
      sheetJadwal.getRange(rowIndex, 7).setValue("");
      
      return { success: true, message: "Pengajuan ditolak. Jadwal dikembalikan ke status aktif." };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Mengubah status terapi akhir pasien (Oleh Terapis)
 */
function setelSelesaiTerapi(idPasien, statusAkhir) {
  try {
    const ss = getDb();
    const sheetPasien = ss.getSheetByName("Pasien");
    const listPasien = sheetToObjects(sheetPasien);
    
    let rowIndex = -1;
    for (let i = 0; i < listPasien.length; i++) {
      if (listPasien[i].ID_Pasien === idPasien) {
        rowIndex = i + 2;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Data pasien tidak ditemukan!" };
    }
    
    sheetPasien.getRange(rowIndex, 11).setValue(statusAkhir); // Status_Terapi
    return { success: true, message: "Status terapi pasien berhasil diubah menjadi: " + statusAkhir };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

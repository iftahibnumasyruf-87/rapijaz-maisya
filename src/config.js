/**
 * PENGATURAN UTAMA APLIKASI (WHITE-LABEL CONFIG)
 * ==============================================
 * Anda dapat mengedit nilai-nilai di bawah ini sebelum mem-build
 * atau men-deploy aplikasi ini untuk sekolah lain.
 */

export const APP_CONFIG = {
    // Nama Singkat Aplikasi (muncul di judul sidebar dan beberapa header)
    appName: "Rapijaz",

    // Nama Sekolah / Instansi (muncul menyambung dengan appName, misal: Rapijaz-Maisya)
    schoolName: "Maisya",

    // Nama Lengkap Instansi (muncul di bagian Login, Footer, atau header utama)
    institutionName: "Ponpes Imam Syafi'i",

    // Deskripsi singkat di halaman Login
    loginDescription: "Sistem Informasi Raport dan Ijazah Terpadu",

    // Teks selamat datang di Dashboard
    welcomeMessage: "Selamat datang di Pusat Informasi Raport dan Ijazah",

    // Teks di bagian bawah layar (Footer)
    footerText: "Sistem Informasi Akademik",

    // Nama pengembang / copyright (misal: "Tim IT Maisya")
    developerName: "Maisya Dev",

    // Tahun rilis aplikasi
    year: "2024",

    // (Opsional) Jika punya URL logo sendiri, isi di sini. Jika kosong akan pakai default.
    // Contoh: "https://i.ibb.co/logo-sekolah.png"
    logoUrl: ""
};

/**
 * Helper untuk mendapatkan nama lengkap aplikasi
 * Contoh Output: "Rapijaz-Maisya"
 */
export const getFullAppName = () => {
    return `${APP_CONFIG.appName}-${APP_CONFIG.schoolName}`;
};

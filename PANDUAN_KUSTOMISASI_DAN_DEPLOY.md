# Panduan Kustomisasi & Deployment Aplikasi Rapijaz

Panduan ini ditujukan untuk Anda yang ingin mengkustomisasi (mengubah nama dan logo) aplikasi ini untuk digunakan oleh sekolah/instansi lain, dan kemudian meng-online-kannya (deploy) secara gratis.

---

## TAHAP 1: Kustomisasi Aplikasi (White-Label)

Sebelum aplikasi di-deploy untuk klien/sekolah baru, Anda dapat mengubah identitas aplikasi dengan sangat mudah tanpa harus mencari-cari kode yang rumit.

### 1. Ubah Teks Utama & Logo
Buka file `src/config.js` menggunakan Notepad, VS Code, atau text editor apa saja. Di dalamnya Anda akan menemukan pengaturan berikut:

```javascript
export const APP_CONFIG = {
    appName: "Rapijaz",
    schoolName: "Maisya", // Ganti dengan nama sekolah klien, misal: "Al-Falah"
    institutionName: "Ponpes Imam Syafi'i", // Ganti nama institusinya
    loginDescription: "Sistem Informasi Raport dan Ijazah Terpadu",
    welcomeMessage: "Selamat datang di Pusat Informasi Raport dan Ijazah",
    // Isi dengan link gambar logo sekolah klien. 
    // Anda bisa mengupload gambar ke imgbb.com lalu tempel link-nya di sini
    logoUrl: "" 
};
```
Silakan ubah teks di dalam tanda kutip `"` sesuai kebutuhan.

### 2. Ubah Judul Tab Browser
Buka file `index.html` yang ada di folder paling luar.
Cari baris ini:
```html
<title>Rapijaz - Aplikasi Raport dan Ijazah</title>
```
Ubah teks di antara `<title>` dan `</title>` sesuai keinginan Anda.

---

## TAHAP 2: Deployment (Meng-online-kan Aplikasi)

Karena aplikasi ini dibangun menggunakan Vite + React dan databasenya berada di Cloud (Firebase), Anda bisa menghostingnya secara **gratis** menggunakan **Vercel** atau **Netlify**.

### Pilihan A: Deploy ke Vercel (Sangat Direkomendasikan)
Vercel sangat dioptimalkan untuk React dan Vite.

1. Buat akun di [Vercel.com](https://vercel.com).
2. Jika Anda memindahkan kode ini ke GitHub Anda sendiri:
   - Klik **"Add New..."** > **"Project"**
   - Import repository GitHub Anda.
3. Jika Anda ingin deploy langsung dari komputer (tanpa GitHub):
   - Install Vercel CLI (buka terminal/CMD, ketik: `npm i -g vercel`)
   - Buka terminal di folder project ini, lalu ketik `vercel` dan tekan Enter. Ikuti petunjuk di layar.
4. **Pengaturan Framework**: Pastikan Vercel mendeteksi *Framework Preset* sebagai **Vite**.
5. **Penting untuk Vercel**: Agar saat me-refresh halaman tidak error "404 Not Found", pastikan Anda sudah memiliki file `vercel.json` di folder utama aplikasi dengan isi:
   ```json
   {
     "rewrites": [ { "source": "/(.*)", "destination": "/index.html" } ]
   }
   ```
6. Klik **Deploy**. Selesai!

### Pilihan B: Deploy ke Netlify
1. Buat akun di [Netlify.com](https://netlify.com).
2. Masuk ke halaman Dashboard > klik **"Add new site"**.
3. Sama seperti Vercel, Anda bisa Import dari GitHub, atau deploy manual dengan menu **"Deploy manually"** (cukup drag-and-drop folder `dist` setelah Anda menjalankan perintah `npm run build` di lokal).
4. **Penting untuk Netlify**: Agar tidak error 404, pastikan ada file bernama `_redirects` di dalam folder `public/` yang isinya hanya satu baris ini:
   ```text
   /*    /index.html   200
   ```
5. Tunggu proses selesai dan aplikasi Anda sudah online!

---

*Selamat berkreasi dan mempromosikan aplikasi ini!*

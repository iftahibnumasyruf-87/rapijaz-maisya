# Panduan Impor Nilai dari Excel

## Pengenalan

Fitur impor nilai Excel memungkinkan Anda untuk **mengimpor data nilai siswa secara massal** dari file Excel (.xlsx) dengan cepat. Ini sangat berguna untuk:
- Mengisi nilai banyak siswa sekaligus
- Update nilai dari dokumen eksternal
- Backup dan restore data

## Cara Menggunakan

### 1. **Unduh Template Excel**
   - Masuk ke menu **Input Nilai**
   - Pilih **Kelas** yang akan diisi nilainya
   - Klik tombol **"Unduh Template Excel"** (ikon Download)
   - File `template_nilai_[KELAS].xlsx` akan terunduh

### 2. **Isi Data Nilai di Excel**
   - Buka file template yang sudah diunduh
   - **Kolom yang ada:**
     - `No` - Nomor urut (otomatis)
     - `NIS` - Nomor Induk Siswa (jangan diubah)
     - `Nama Santri` - Nama siswa (jangan diubah)
     - `[MAPEL] - UTS` - Nilai UTS mata pelajaran (dapat diisi dengan angka Arab atau Latin)
     - `[MAPEL] - UAS` - Nilai UAS mata pelajaran (dapat diisi dengan angka Arab atau Latin)

   **Contoh:**
   | No | NIS | Nama Santri | Matematika - UTS | Matematika - UAS | B. Indonesia - UTS | B. Indonesia - UAS |
   |----|-----|-------------|------------------|------------------|-------------------|-------------------|
   | 1  | 001 | Ahmad       | 85               | 88               | 75                | 80                |
   | 2  | 002 | Badria      | 90               | 92               | 85                | 88               |

### 3. **Impor File ke Aplikasi**
   - Kembali ke menu **Input Nilai**
   - Pilih **Kelas yang sama** seperti saat unduh template
   - Klik tombol **"Impor dari Excel"** (ikon Upload)
   - Pilih file Excel yang sudah diisi
   - Tunggu proses selesai (akan muncul notifikasi sukses)

### 4. **Verifikasi Data**
   - Data akan langsung tampil di tabel
   - Nilai raport akan otomatis terhitung (40% UTS + 60% UAS)
   - Tekan **"Simpan Manual"** untuk menyimpan perubahan

## Aturan Penting

### Format Data
- **NIS & Nama Siswa**: Jangan diubah dari template yang diunduh
- **Nilai**: Gunakan format angka
  - Format Latin: `85.5`, `80`, `92`
  - Format Arab: `٨٥٫٥`, `٨٠`, `٩٢`
  - Aplikasi akan otomatis mengkonversi ke format Latin
- **Nilai Kosong**: Biarkan kosong jika tidak ada nilai
- **Tidak ada header tambahan**: Jangan menambah kolom baru

### Validasi
- ✅ Hanya kolom UTS dan UAS per mata pelajaran yang diproses
- ✅ Siswa dicocokkan berdasarkan NIS atau Nama
- ⚠️ Jika siswa tidak ditemukan, baris tersebut akan dilewati
- ⚠️ Jika kolom mata pelajaran tidak sesuai, nilai tidak akan diimpor

### Batas Ukuran
- File Excel maksimal **10 MB**
- Impor dapat dilakukan untuk **semua siswa di kelas sekaligus**
- Jika data banyak (>1000 siswa), pecah ke beberapa file dan impor bertahap

## Troubleshooting

### "Pilih kelas dan pastikan ada mata pelajaran"
**Solusi:**
- Pastikan Anda sudah memilih kelas di dropdown "Pilih Kelas"
- Pastikan ada mata pelajaran yang diatur di Master Data untuk kelas tersebut

### Siswa tidak ditemukan / data tidak terimport
**Penyebab:**
- NIS atau Nama Siswa di Excel tidak cocok dengan database
- Siswa belum didaftarkan di sistem

**Solusi:**
- Cek kembali NIS dan nama siswa di Excel
- Pastikan siswa sudah terdaftar di menu Master Data > Siswa
- Unduh template ulang untuk mendapatkan data yang akurat

### Nilai masuk tapi dengan angka aneh
**Penyebab:**
- Format angka di Excel tidak valid (misal: text, formula, dll)

**Solusi:**
- Pastikan semua nilai dalam format angka plain (bukan formula)
- Ubah format sel Excel ke "Number" untuk semua kolom nilai

### File tidak bisa diunduh
**Solusi:**
- Periksa koneksi internet
- Coba refresh halaman browser
- Gunakan browser yang berbeda

## Tips & Trik

### 1. **Backup Nilai Reguler**
Setiap minggu, unduh template Excel untuk backup data nilai secara otomatis.

### 2. **Update Bertahap**
- Unduh template pada minggu pertama
- Edit sebagian kolom nilai
- Impor kembali - hanya kolom yang diisi yang terupdate
- Data lama yang tidak diubah akan tetap utuh

### 3. **Merge Nilai dari Multiple Source**
- Dari Excel eksternal (guru, arsip, dll) dapat langsung diimpor
- Nilai yang ada di sistem akan di-update sesuai data terbaru

### 4. **Angka Arab Support**
Anda dapat mengetik nilai langsung menggunakan keyboard Arab (٠-٩) atau Latin (0-9), sistem akan otomatis mengenali dan mengkonversi.

## Contoh Skenario

### Skenario 1: Isi Nilai Baru Dari Nol
1. Pilih Kelas 7
2. Unduh template
3. Tulis semua nilai UTS dan UAS untuk Matematika dan B. Indonesia
4. Impor file
5. Nilai langsung muncul di tabel

### Skenario 2: Update Nilai UAS saja
1. Nilai UTS sudah ada di sistem
2. Unduh template (yang sudah berisi UTS)
3. Edit hanya kolom "UAS"
4. Impor file
5. Hanya UAS yang terupdate, UTS tetap sama

### Skenario 3: Pindahan Data dari Aplikasi Lain
1. Export dari aplikasi lain ke format Excel dengan kolom: NIS, Nama, [Mapel]-UTS, [Mapel]-UAS
2. Sesuaikan nama-nama kolom dengan template yang diunduh
3. Impor ke sistem ini
4. Data berhasil dimigrasikan

## FAQ

**Q: Bisakah impor dari file CSV?**
A: Tidak, hanya format Excel (.xlsx atau .xls) yang didukung.

**Q: Apakah nilai lama akan terhapus saat impor?**
A: Tidak, impor akan **merge** dengan data lama. Jika ada nilai baru untuk siswa yang sama, nilai lama akan diupdate.

**Q: Berapa lama proses impor?**
A: Biasanya hanya beberapa detik untuk 100+ siswa, tergantung ukuran file dan kecepatan koneksi.

**Q: Apakah bisa impor nilai untuk multiple kelas sekaligus?**
A: Tidak, harus per kelas. Untuk multiple kelas, lakukan impor berkali-kali.

**Q: Kemana data tua disimpan jika ada update?**
A: Data tua akan ditimpa, tapi Anda dapat unduh backup Excel sebelum impor jika perlu.

---

**Pertanyaan atau Masalah?** Hubungi admin atau baca dokumentasi lebih lengkap.

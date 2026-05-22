# Panduan Penggunaan Fitur Excel Import Nilai

## Ringkas

Aplikasi Rapijaz Maisya sekarang mendukung **import nilai dari Excel** untuk mengisi nilai siswa secara massal dengan cepat dan efisien.

---

## 📋 Tabel Isi

1. [Fitur-Fitur](#fitur-fitur)
2. [Cara Kerja](#cara-kerja)
3. [Format Template](#format-template)
4. [Panduan Langkah-Langkah](#panduan-langkah-langkah)
5. [Troubleshooting](#troubleshooting)
6. [Tips & Trik](#tips--trik)

---

## 🎯 Fitur-Fitur

### ✅ Yang Didukung
- **Unduh Template Excel**: Download template kosong yang sudah berisi daftar siswa kelas
- **Import Nilai**: Impor nilai UTS dan UAS untuk semua mata pelajaran sekaligus
- **Merge Data**: Impor akan merge dengan nilai existing (tidak menghapus data lama)
- **Validasi Otomatis**: Sistem cek kecocokan NIS/Nama siswa
- **Konversi Angka**: Mendukung input angka Arab (٠-٩) atau Latin (0-9)
- **Auto-Calculate**: Nilai raport otomatis dihitung (40% UTS + 60% UAS)
- **Freeze Panes**: Template sudah memiliki frozen headers untuk kemudahan editing

---

## 🔄 Cara Kerja

```
┌─────────────────────────────────────────────────────┐
│ 1. Pilih Kelas di Menu Input Nilai                 │
└─────────────────────────────────────────────────────┘
                      ↓
┌─────────────────────────────────────────────────────┐
│ 2. Klik "Unduh Template Excel"                      │
│    → File .xlsx download dengan daftar siswa kelas  │
└─────────────────────────────────────────────────────┘
                      ↓
┌─────────────────────────────────────────────────────┐
│ 3. Isi nilai UTS & UAS di Excel                     │
│    → Edit kolom [MAPEL] - UTS dan UAS               │
│    → Simpan file Excel                              │
└─────────────────────────────────────────────────────┘
                      ↓
┌─────────────────────────────────────────────────────┐
│ 4. Klik "Impor dari Excel"                          │
│    → Pilih file yang sudah diisi                    │
│    → Tunggu proses selesai                          │
└─────────────────────────────────────────────────────┘
                      ↓
┌─────────────────────────────────────────────────────┐
│ 5. Verifikasi & Simpan                              │
│    → Data muncul di tabel                           │
│    → Nilai raport otomatis dihitung                 │
│    → Klik "Simpan Manual" untuk save                │
└─────────────────────────────────────────────────────┘
```

---

## 📊 Format Template

### Struktur Kolom

| Kolom | Tipe | Keterangan | Edit? |
|-------|------|-----------|-------|
| No | Angka | Nomor urut | ❌ Jangan diubah |
| NIS | Text | Nomor Induk Siswa | ❌ Jangan diubah |
| Nama Santri | Text | Nama siswa | ❌ Jangan diubah |
| [MAPEL] - UTS | Angka | Nilai Ujian Tengah Semester | ✅ **ISI DI SINI** |
| [MAPEL] - UAS | Angka | Nilai Ujian Akhir Semester | ✅ **ISI DI SINI** |

### Contoh Template (Kelas 7)

```
No | NIS  | Nama Santri      | Matematika-UTS | Matematika-UAS | B.Indo-UTS | B.Indo-UAS
---|------|------------------|----------------|----------------|------------|-------------
1  | 001  | Ahmad Abdullah   | 85             | 88             | 75         | 80
2  | 002  | Badria Siti      | 90             | 92             | 85         | 88
3  | 003  | Cilak Maulana    | 78             | 82             | 70         | 75
```

### Format Nilai

**Angka Arab (٠-٩)**
- Contoh: `٨٥`, `٩٢`, `٧٨٫٥`
- Sistem akan auto-konversi ke Latin

**Angka Latin (0-9)**
- Contoh: `85`, `92`, `78.5`
- Format standar

**Catatan:**
- Desimal: gunakan `.` (titik) atau `,` (koma), keduanya diterima
- Nilai kosong: biarkan kosong untuk kolom yang tidak ada nilai
- Batas: angka 0-100

---

## 🚀 Panduan Langkah-Langkah

### Langkah 1: Masuk ke Menu Input Nilai

1. Buka aplikasi Rapijaz Maisya
2. Di sidebar kiri, klik **"Input Nilai"** → **"Pelajaran"**
3. Anda akan melihat halaman Input Nilai dengan:
   - Dropdown "Pilih Kelas"
   - Tombol "Unduh Template Excel" (biru)
   - Tombol "Impor dari Excel" (oranye)
   - Tabel nilai di bawahnya

### Langkah 2: Pilih Kelas

1. Klik dropdown **"Pilih Kelas"**
2. Pilih kelas yang akan diisi nilainya (misal: Kelas 7)
3. Pastikan sudah ada mata pelajaran untuk kelas tersebut

### Langkah 3: Unduh Template

1. Klik tombol biru **"Unduh Template Excel"** 
2. Browser akan download file `template_nilai_[KELAS].xlsx`
3. File berisi daftar semua siswa di kelas itu

### Langkah 4: Edit di Excel

1. Buka file Excel yang diunduh
2. Anda akan melihat:
   - **Header membeku** (tidak geser saat scroll)
   - **Kolom siswa membeku** (kolom No, NIS, Nama tidak geser horizontal)
   - Siswa dengan data sudah terisi dengan nilai lama (jika ada)
   - Baris kosong siap untuk nilai baru
3. **Edit kolom UTS dan UAS** untuk setiap mata pelajaran
4. Contoh pengisian:
   ```
   Ahmad Abdullah: Matematika UTS=85, Matematika UAS=88, B.Indo UTS=75, B.Indo UAS=80
   ```
5. **Simpan file** (Ctrl+S atau File → Save)

### Langkah 5: Impor Kembali

1. Kembali ke aplikasi Rapijaz
2. Pastikan kelas yang sama masih terpilih
3. Klik tombol oranye **"Impor dari Excel"**
4. Jendela file picker akan terbuka
5. Pilih file Excel yang sudah diisi
6. Tunggu notifikasi "✓ [X] siswa berhasil diimpor!"

### Langkah 6: Verifikasi & Simpan

1. Data akan langsung muncul di tabel
2. Nilai raport otomatis terhitung (40% UTS + 60% UAS)
3. Periksa apakah ada yang perlu koreksi
4. Klik **"Simpan Manual"** untuk menyimpan semua perubahan ke database
5. Notifikasi "Tersimpan otomatis" akan muncul

---

## ❓ Troubleshooting

### Error: "Pilih kelas dan pastikan ada mata pelajaran"

**Penyebab:**
- Kelas belum dipilih
- Mata pelajaran belum didaftarkan untuk kelas tersebut

**Solusi:**
```
1. Pastikan dropdown "Pilih Kelas" sudah memilih kelas
2. Jika belum ada mata pelajaran:
   - Pergi ke "Master Data" → "Mata Pelajaran"
   - Pastikan mata pelajaran sudah dibuat dan ditugaskan ke kelas
3. Refresh halaman jika perlu
```

### Error: "Gagal mengimpor file Excel. Pastikan format sesuai template."

**Penyebab:**
- File yang diupload bukan dari template resmi
- Format kolom tidak sesuai
- File rusak atau corrupted

**Solusi:**
```
1. Unduh template yang baru dari aplikasi
2. Copy-paste data dari file lama ke template baru
3. Pastikan kolom UTS/UAS nama sesuai: "[MAPEL] - UTS" dan "[MAPEL] - UAS"
4. Coba buka file Excel dengan Excel (bukan Google Sheets) sebelum upload
5. Pastikan no NA ada (tidak ada simbol khusus di nama kolom)
```

### Siswa tidak ditemukan / data tidak terimport

**Penyebab:**
- NIS di Excel tidak cocok dengan database
- Nama siswa di Excel berbeda dengan di database (typo, spasi, huruf kapital)
- Siswa belum terdaftar

**Solusi:**
```
1. Cek database siswa:
   - Pergi ke "Master Data" → "Siswa"
   - Lihat daftar NIS dan nama yang benar
2. Edit file Excel sesuai data di database
3. Pastikan tidak ada spasi berlebih di awal/akhir nama
4. Jika masih gagal, buat siswa baru di Master Data terlebih dahulu
```

### Nilai masuk tapi dengan angka aneh

**Penyebab:**
- Format Excel salah (formula, text, dll)
- Angka Arab tercampur Latin

**Solusi:**
```
1. Di Excel, format kolom UTS/UAS sebagai "Number"
2. Hapus semua formula, hanya isi angka plain
3. Gunakan angka standar (0-100) tanpa simbol
4. Untuk desimal, gunakan: 85.5 atau 85,5 (keduanya diterima)
```

### File Excel tidak bisa diunduh

**Penyebab:**
- Browser block download
- Koneksi terputus
- Error server

**Solusi:**
```
1. Cek koneksi internet
2. Disable popup blocker di browser
3. Coba browser lain (Chrome, Firefox, Edge)
4. Refresh halaman dan coba lagi
5. Lihat folder "Downloads" apakah file sudah ada
```

### Nilai lama hilang setelah impor

**Penyebab:**
- Data lama di-overwrite dengan data baru

**Catatan:**
- Impor menggunakan **merge**, bukan replace
- Jika ada siswa dan mata pelajaran yang sama, nilai akan di-update
- Jika ada siswa/mapel yang berbeda, akan ditambahkan

**Solusi:**
```
1. Jika ingin restore, unduh backup dari export sebelumnya (jika ada)
2. Untuk mencegah di masa depan: selalu backup sebelum impor besar
3. Jika file Excel masih ada, unduh template lagi dan impor dengan data lama
```

---

## 💡 Tips & Trik

### ✅ Best Practice

**1. Selalu Backup Sebelum Impor Besar**
```
- Unduh template untuk backup
- Edit di backup, bukan di original
```

**2. Update Bertahap**
```
- Jangan semuanya sekaligus
- Impor per mata pelajaran atau per minggu
```

**3. Validasi Data Dulu**
```
- Sebelum impor, cek di Excel:
  - Tidak ada NIS duplikat
  - Tidak ada nama yang kosong/aneh
  - Nilai hanya angka (0-100)
```

**4. Gunakan Format Standar**
```
- Angka: 85, 90.5, 78 (bukan "delapan puluh lima")
- Nama: "Ahmad Abdullah" (tidak ada spasi berlebih)
- NIS: "001" atau "1" (keduanya OK)
```

### 🎯 Skenario Penggunaan

**Skenario 1: Input Nilai Baru (Blank to Filled)**
```
1. Unduh template (semua kosong)
2. Isi semua nilai UTS/UAS
3. Impor
4. Otomatis tersimpan dengan nilai raport dihitung
```

**Skenario 2: Update UAS Saja (Existing Value)**
```
1. Unduh template (nilai UTS sudah ada)
2. Edit hanya kolom UAS
3. Impor (hanya UAS yang terupdate)
4. UTS tetap sama
```

**Skenario 3: Migrasi dari Aplikasi Lain**
```
1. Export dari aplikasi lain ke Excel
2. Sesuaikan kolom: No, NIS, Nama Santri, [Mapel]-UTS, [Mapel]-UAS
3. Unduh template dari sini
4. Copy-paste data ke template yang benar
5. Impor
```

**Skenario 4: Perbaikan Massal**
```
1. Jika banyak nilai salah, unduh template lagi
2. Copy nilai yang benar ke template
3. Impor (hanya nilai yang diubah akan terupdate)
```

### 🔧 Advanced Tips

**Merge Multiple Sources**
- Import dari file A (matematika)
- Import dari file B (bahasa)
- Keduanya akan merge di database

**Batch Processing**
- Untuk >1000 siswa, pecah jadi 2-3 file
- Import satu per satu
- Sistem akan auto-merge

**Backup Rutin**
- Setiap akhir bulan, unduh template sebagai backup
- Simpan di folder archive
- Gunakan untuk restore jika ada masalah

---

## 📞 Support

Jika ada masalah atau pertanyaan:
1. Lihat [PANDUAN_IMPOR_NILAI.md](PANDUAN_IMPOR_NILAI.md) untuk detail lengkap
2. Hubungi admin atau guru IT
3. Check file [template_nilai_sample.xlsx](template_nilai_sample.xlsx) untuk contoh format

---

**Last Updated:** Mei 2026  
**Version:** 1.0

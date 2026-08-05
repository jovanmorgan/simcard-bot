# Rekap Burn Kuota per City — Web

Versi web dari script `rekap-data.js`. Bedanya: pemilihan "mulai dari sheet mana"
sekarang dilakukan lewat dropdown di browser, dan sekarang juga bisa pilih
sampai sheet mana (rentang, bukan cuma titik awal).

## 1. Persiapan

1. Copy folder ini ke komputer/server kamu.
2. Taruh file `credentials.json` (service account Google yang sama dengan
   yang dipakai script lama) di folder root project ini (sejajar dengan
   `server.js`).
3. Copy `.env.example` jadi `.env`, lalu isi `SHEET_URL` dengan link
   spreadsheet kamu (boleh dikosongkan — nanti bisa diisi manual dari web).
4. Install dependency:

   ```bash
   npm install
   ```

## 2. Jalankan

```bash
npm start
```

Lalu buka `http://localhost:3000` di browser.

## 3. Cara pakai di web

1. Isi/cek link spreadsheet (otomatis terisi dari `.env` kalau ada), klik
   **Muat daftar sheet**.
2. Pilih **Dari sheet** dan **Sampai sheet** — preview pita di bawahnya akan
   menyorot sheet mana saja yang bakal ikut diproses.
3. Klik **Export ke Excel** — file `.xlsx` otomatis ke-download di browser,
   dan juga tersimpan di folder `Hasil-Rekap/` di server (sama seperti
   perilaku script CLI lama).
4. Ringkasan hasil (jumlah CITY, bulan dominan, sheet yang dilewati, baris
   tanpa MSISDN, dll) muncul di bagian "Hasil".

## Struktur project

```
rekap-web/
├─ server.js          # Express server (endpoint API + serving frontend)
├─ lib/rekap.js        # Logic inti (baca Google Sheets, susun Excel) — hasil refactor dari script CLI lama, tanpa readline
├─ public/             # Frontend (HTML/CSS/JS polos, tanpa build step)
├─ Hasil-Rekap/         # Folder tempat file .xlsx hasil export tersimpan
├─ credentials.json     # (kamu tambahkan sendiri) service account Google
└─ .env                 # (kamu buat dari .env.example)
```

## Catatan

- Logic pembacaan sheet, deteksi kolom CITY/MSISDN, parsing "Tanggal Pasang",
  dan styling Excel **sama persis** dengan script CLI aslinya — cuma dibungkus
  jadi fungsi yang bisa dipanggil dari web, dan sumber input sheet-nya lewat
  dropdown, bukan lewat pertanyaan di terminal.
- Kalau sheet awal/akhir yang dipilih ternyata terbalik posisinya, sistem
  otomatis menukar urutan supaya tetap jalan dengan benar.

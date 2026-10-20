# LABSOMA PPS — Apps Script

Aplikasi penugasan dan pembuatan soal ujian madrasah berbasis Google Apps Script dan Google Spreadsheet.

## Isi proyek

- `Code.gs` — fungsi server dan API Apps Script.
- `Index.html` — antarmuka aplikasi.
- `appsscript.json` — manifest runtime V8.

## Menjalankan di Google Apps Script

1. Buat atau buka Google Spreadsheet yang akan menjadi basis data.
2. Buka **Ekstensi → Apps Script**.
3. Ganti isi `Code.gs` dan buat file HTML bernama `Index`, lalu salin isi `Index.html`.
4. Aktifkan tampilan manifest di pengaturan proyek Apps Script, lalu salin `appsscript.json` bila belum otomatis digunakan.
5. Simpan, muat ulang Spreadsheet, lalu pilih **LABSOMA PPS → Siapkan sheet**.
6. Jika menu belum muncul, jalankan fungsi `setup` sekali dari editor dan berikan izin Google yang diminta.

## Sinkronisasi dengan GitHub menggunakan clasp

Simpan proyek ini di repository GitHub privat jika data atau konfigurasi madrasah bersifat internal. Jangan commit `.clasp.json`, karena file itu berisi ID proyek Apps Script.

1. Pasang Node.js, lalu jalankan `npm install -g @google/clasp`.
2. Jalankan `clasp login` dan masuk dengan akun Google pemilik proyek.
3. Buat atau pilih proyek Apps Script yang terikat ke Spreadsheet.
4. Dari folder yang tertaut ke proyek jalankan `clasp clone SCRIPT_ID` untuk menarik file proyek yang sudah ada. Tinjau perubahan sebelum menimpa proyek produksi.
5. Jalankan `clasp push` untuk mengirim perubahan ke Apps Script.
6. Inisialisasi Git, tambahkan remote GitHub, lalu commit dan push:

```sh
git init
git add Code.gs Index.html appsscript.json README.md .claspignore
git commit -m "Optimize LABSOMA PPS Apps Script"
git remote add origin URL_REPOSITORY
git push -u origin main
```

`clasp push` mengirim kode ke Apps Script, sedangkan `git push` mengirim salinan kode ke GitHub. Deployment web app dan izin akses perlu ditinjau sesuai kebijakan madrasah.

## Perubahan performa

`api_init()` sebelumnya memformat kolom teks sampai baris maksimum setiap sheet pada setiap pembukaan aplikasi. Sekarang pemuatan awal melewati pemformatan berulang, dan menu **Siapkan sheet** membatasi pemformatan ke baris yang sudah terpakai. Ini mengurangi pekerjaan Spreadsheet service saat aplikasi dibuka.
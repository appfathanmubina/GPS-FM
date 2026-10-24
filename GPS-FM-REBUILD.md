# GPS FM — Rebuild awal

Aplikasi mencatat **kehadiran guru GPS** pada kegiatan pesantren dan menyediakan rekap bulanan. Absensi santri tidak termasuk dalam tahap awal.

## Hak akses
- Superadmin: kelola akun guru, kegiatan, penugasan PJ, dan penguncian periode.
- Guru GPS: dapat melihat rekap bulanan seluruh kegiatan.
- PJ: dapat menginput kehadiran hanya pada kegiatan yang ditugaskan. Pemeriksaan akses dilakukan di backend.

## Struktur data
`Users`, `Activities`, `ActivityPJ`, `Attendance`, `Periods`, `AuditLogs`, dan `AppSettings`.

## File baru
- `apps-script/Code.gs`
- `apps-script/SetupDatabase.gs`
- `apps-script/Index.html`
- `pwa/index.html`, `pwa/config.js`, `pwa/manifest.json`, `pwa/sw.js`, `pwa/icon.svg`

## Setup Apps Script
1. Buat project Apps Script baru.
2. Salin `Code.gs`, `SetupDatabase.gs`, dan HTML `Index.html` dari folder `apps-script/`.
3. Di Script Properties, set `GPSFM_INITIAL_ADMIN_PASSWORD` ke password awal minimal 12 karakter. Jangan menyimpan password di kode.
4. Jalankan `setupGpsFmDatabase()` satu kali. Database Spreadsheet dibuat otomatis dan ID disimpan ke Script Properties.
5. Deploy sebagai Web App; untuk skenario PWA, pastikan pengaturan akses deployment sesuai kebijakan pesantren.
6. Isi `pwa/config.js` dengan URL deployment berakhiran `/exec`.
7. Atur GitHub Pages agar menyajikan folder `pwa/`, atau pindahkan file PWA ke root hanya setelah memastikan tidak menimpa file aktif.

## Status
Ini adalah baseline implementasi di branch `gps-fm-rebuild-v1`, bukan rilis produksi yang sudah lulus pengujian. Uji alur login, seluruh role, PJ lintas kegiatan, duplikasi absensi, penguncian periode, dan instalasi PWA sebelum digunakan secara resmi.
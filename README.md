# SOP CIMS — Kopi Kenangan

**Continuous Improvement Management System** untuk Kopi Kenangan, berbasis Google Workspace (Apps Script + Spreadsheet).

## Isi Repository

### Apps Script Project (`1QK2mARkWVjDKXnw9CfLo1hDIw6JYecsU2P7LFKO1ooOBEiy0LVCY4rgg`)

| File | Deskripsi |
|------|-----------|
| `appsscript.json` | Manifest konfigurasi project Apps Script |
| `Code.js` | Logic utama sistem CIMS |
| `Config.js` | Konfigurasi (URL, ID, settings) |
| `Dashboard.js` | Logic dashboard monitoring |
| `Mail.js` | Pengiriman email notifikasi & reminder |
| `Reminders.js` | Logic pengingat otomatis |
| `SOP.js` | Generator dokumen SOP & slide presentasi |
| `Utility.js` | Fungsi-fungsi helper/utilitas |
| `WebApp.js` | Web App handler (doGet, doPost, routing) |
| `Approval.html` | Template halaman approval |
| `Closing.html` | Template halaman closing |
| `EmailTemplate.html` | Template email |
| `Revision.html` | Template halaman revisi |

### Data

| File | Deskripsi |
|------|-----------|
| `CI Database (1).xlsx` | Database Continuous Improvement (CI) dalam format spreadsheet |

## Sync Apps Script ↔ GitHub

Project ini terhubung ke Google Apps Script via [clasp](https://github.com/google/clasp).

### Pull (download dari Apps Script ke lokal):
```bash
clasp pull
git add .
git commit -m "Sync from Apps Script"
git push
```

### Push (upload dari lokal ke Apps Script):
```bash
clasp push
```

### Clone pertama kali:
```bash
clasp login
clasp clone 1QK2mARkWVjDKXnw9CfLo1hDIw6JYecsU2P7LFKO1ooOBEiy0LVCY4rgg --rootDir .
```

## Struktur Sistem

- **Initiator / Staff** — mengajukan proyek CI via Google Form
- **Reviewer** — meninjau dan merevisi pengajuan
- **Approver** — memberikan persetujuan final
- **CIMS Dashboard** — monitoring real-time via spreadsheet

## Lisensi

Internal use only — Kopi Kenangan.

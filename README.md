# SOP CIMS — Kopi Kenangan

**Continuous Improvement Management System** untuk Kopi Kenangan, berbasis Google Workspace (Apps Script + Spreadsheet).

## Isi Repository

| File | Deskripsi |
|------|-----------|
| `SOP.gs` | Google Apps Script — generator dokumen SOP & slide presentasi otomatis di Google Drive |
| `CI Database (1).xlsx` | Database Continuous Improvement (CI) dalam format spreadsheet |

## Cara Pakai Apps Script

1. Buka [Google Apps Script](https://script.google.com)
2. Buat project baru, lalu copy-paste isi `SOP.gs`
3. Jalankan fungsi `createSOPDocumentInDrive()` — dokumen SOP akan terbuat otomatis di Google Drive Anda

## Struktur Sistem

- **Initiator / Staff** — mengajukan proyek CI via Google Form
- **Reviewer** — meninjau dan merevisi pengajuan
- **Approver** — memberikan persetujuan final
- **CIMS Dashboard** — monitoring real-time via spreadsheet

## Lisensi

Internal use only — Kopi Kenangan.

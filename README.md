# CAM Recovery Action Dashboard

Dashboard Streamlit untuk membaca `MASTER_PROGRESS.xlsx`, menampilkan ringkasan executive, dan mendukung alur otomatis dari Google Drive.

## Struktur Folder

```text
.
├── app.py
├── combiner.py
├── Process.py
├── drive_utils.py
├── requirements.txt
├── data/
│   ├── raw/
│   │   └── *.xlsx
│   └── main/
│       └── MASTER_PROGRESS.xlsx
├── secret/
│   └── *.json
└── Software Requirement Document.md
```

## Cara Menjalankan

1. Install dependency:

```bash
pip install -r requirements.txt
```

2. Jalankan dashboard:

```bash
streamlit run app.py
```

## Alur Data

- Upload file project raw ke folder Google Drive `raw`
- Jalankan `combiner.py` untuk membangun `MASTER_PROGRESS.xlsx`
- `MASTER_PROGRESS.xlsx` otomatis di-upload ke folder Google Drive `main`
- Streamlit membaca `MASTER_PROGRESS.xlsx` dari Drive

## Konfigurasi Google Drive

- Simpan service account JSON di folder `secret/`
- Share folder Drive ke email service account
- Gunakan folder utama Drive yang berisi subfolder `raw` dan `main`
- App akan mencari `MASTER_PROGRESS.xlsx` di Drive secara otomatis
- Jika pakai GitHub Actions, tambahkan secrets:
  - `GOOGLE_SERVICE_ACCOUNT_JSON`
  - `GOOGLE_DRIVE_ROOT_FOLDER`

## Setup Cepat

1. Upload raw Excel ke folder Drive `raw`.
2. Pastikan folder root Drive ini di-share ke service account.
3. Simpan JSON service account ke `secret/` untuk lokal, atau ke GitHub Secrets untuk cloud.
4. Tambahkan secret GitHub:
   - `GOOGLE_SERVICE_ACCOUNT_JSON`
   - `GOOGLE_DRIVE_ROOT_FOLDER`
5. Jalankan workflow `Sync MASTER_PROGRESS from Google Drive` atau tunggu jadwal otomatis.
6. Streamlit akan membaca `MASTER_PROGRESS.xlsx` terbaru dari folder `main`.

## Setup Streamlit Cloud

Salin `.streamlit/secrets.toml.example` ke `.streamlit/secrets.toml`, lalu isi:

- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_DRIVE_ROOT_FOLDER`

## Otomatisasi Penuh

Workflow `.github/workflows/sync_master.yml` menjalankan `combiner.py` setiap 15 menit atau manual via `workflow_dispatch`.
Alurnya:

`Google Drive raw -> combiner.py -> Google Drive main -> Streamlit`

## Fitur Utama

- KPI task, completion, dan achievement
- Chart status task dan due status
- Tren `Daily Plan vs Actual`
- Tren achievement dari workbook
- Perbandingan beberapa workbook sekaligus

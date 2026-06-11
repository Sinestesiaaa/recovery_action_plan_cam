# CAM Recovery Action Dashboard

Dashboard Streamlit untuk membaca workbook CAM yang sudah ada, menampilkan ringkasan task, tren harian, dan perbandingan antar file.

## Struktur Folder

```text
.
├── app.py
├── Process.py
├── requirements.txt
├── data/
│   └── raw/
│       └── *.xlsx
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

- Letakkan workbook `.xlsx` di `data/raw/`
- Atau upload file langsung dari sidebar dashboard
- Dashboard otomatis membaca sheet `DATA` dan `Activity Plan`

## Fitur Utama

- KPI task, completion, dan achievement
- Chart status task dan due status
- Tren `Daily Plan vs Actual`
- Tren achievement dari workbook
- Perbandingan beberapa workbook sekaligus

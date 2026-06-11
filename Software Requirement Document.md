# Software Requirement Document

## Tujuan

Membangun dashboard Streamlit untuk workbook CAM agar data olahan Excel bisa dibaca sebagai ringkasan operasional, tren, dan perbandingan antar project.

## Input

- File workbook `.xlsx`
- Sheet utama: `DATA`
- Sheet pendukung: `Activity Plan`

## Kebutuhan Fungsional

- Menampilkan daftar workbook yang tersedia
- Memilih workbook aktif dari sidebar
- Menampilkan KPI utama task dan achievement
- Menampilkan chart status, due status, dan tren harian
- Menampilkan tabel detail yang bisa difilter
- Membandingkan beberapa workbook sekaligus
- Menerima file upload tambahan dari pengguna

## Keluaran

- Dashboard interaktif berbasis Streamlit
- Tabel ringkasan data yang sudah dibersihkan
- Visualisasi tren dan status

## Struktur Data

- `Process.py` sebagai parser dan normalizer workbook
- `app.py` sebagai entry point dashboard
- `data/raw/` sebagai folder workbook sumber

## Catatan

- Workbook lama tetap bisa dipakai tanpa diubah
- File baru cukup dimasukkan ke `data/raw/` atau di-upload dari sidebar

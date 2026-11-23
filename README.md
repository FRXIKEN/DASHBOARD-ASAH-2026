# 🌸 Dashboard by.ficlo

Dashboard Business Intelligence untuk UMKM **by.ficlo** (buket bunga & dekorasi), dibuat dengan Streamlit.

**5 halaman:** Overview · Sales Analytics · Inventory Monitoring · Order Monitoring · Financial Analytics

## Struktur folder

```text
.
├── app.py                          # dashboard Streamlit
├── data_loader.py                  # pembersihan & perhitungan data dari Excel
├── Dataset_Dasboard_DS_byficlo.xlsx # data
├── Dashboard_byficlo.ipynb         # notebook (membuat file-file di atas + menjalankan dashboard)
├── requirements.txt
└── .streamlit/config.toml          # tema hijau-pink
```

## Deploy ke Streamlit Community Cloud

1. Upload seluruh isi folder ini ke repository GitHub (file harus ada di root repo, bukan di dalam subfolder).
2. Buka <https://share.streamlit.io> → login dengan GitHub → **Create app**.
3. Pilih repository dan branch, isi **Main file path** dengan `app.py`, lalu klik **Deploy**.

## Jalankan lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Memperbarui data

- Ganti file `Dataset_Dasboard_DS_byficlo.xlsx` di GitHub (commit ulang), lalu klik **🔄 Muat ulang data** di dashboard; atau
- unggah file `.xlsx` baru langsung lewat panel **🔎 Filter & sumber data** di dashboard (tanpa mengubah GitHub).

## Catatan data

Kolom `[AUTO]` di Excel dihitung ulang oleh Python, dan masalah data ditampilkan di bagian
**🧹 Catatan kualitas data** pada halaman Overview.

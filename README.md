# V.6 Padebuolo Fresh

Aplikasi kas rumah dinas berbasis Streamlit + SQLite, dibangun ulang dari nol supaya lebih bersih.

## Fitur

- Dashboard saldo kas Azka/Rayhan
- Input transaksi masuk/keluar
- Inject dana / top up
- Split pengeluaran
- Buku besar dengan rekap sisa kas
- Kategorisasi massal
- Budget bulanan standar:
  - Rayhan: Rp715.000
  - Azka: Rp760.000
- Pergerakan belanja bulanan per kategori
- Import/export Excel, CSV, JSON
- GitHub cloud backup supaya data tetap aman setelah reboot/redeploy

## Cara jalan lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

Default password: `rumdin123`

## Cara deploy Streamlit Cloud

1. Upload semua isi folder ini ke repo GitHub.
2. Streamlit Cloud -> New app.
3. Repository: repo app ini.
4. Branch: `main`.
5. Main file path: `app.py`.
6. Secrets minimal:
   ```toml
   APP_PASSWORD = "password_app_lo"
   ```

## Supaya data tidak hilang

Streamlit Cloud bisa menghapus file lokal saat idle/reboot/redeploy. Karena database SQLite berupa file lokal, aktifkan GitHub backup.

Saran: buat repo private terpisah, misalnya `padebuolo-data`.
Lalu buat GitHub fine-grained token dengan akses repo data tersebut dan permission `Contents: Read and Write`.

Secrets:

```toml
APP_PASSWORD = "password_app_lo"
PERSISTENCE_PROVIDER = "github"
GITHUB_TOKEN = "github_pat_xxx"
GITHUB_REPO = "username/padebuolo-data"
GITHUB_BRANCH = "main"
GITHUB_DATA_FILE = "padebuolo_live_backup.json"
```

Setelah app jalan, buka menu Pengaturan -> Simpan backup cloud sekarang.

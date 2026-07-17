# Cara Upload ke GitHub

Cara paling gampang:

1. Extract ZIP ini.
2. Buka repo GitHub lama, atau bikin repo baru.
3. Upload semua isi folder hasil extract ke root repo.
4. Pastikan struktur repo seperti ini:

```text
app.py
requirements.txt
.gitignore
.streamlit/config.toml
SECRETS_TEMPLATE.toml
README.md
data/
  seed_transactions.csv
  seed_budgets.csv
  Rekap Kas Rumdin.xlsx
  Rekap_Kas_Rumdin_import_ready.xlsx
```

5. Deploy di Streamlit Cloud dengan main file path: `app.py`.

Catatan:
- Jangan upload `.streamlit/secrets.toml`.
- API AI tidak ada lagi. File ini sudah bersih dari AI Assistant.
- Untuk data persisten, pakai Secrets GitHub backup seperti di README.

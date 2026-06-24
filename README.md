# XLSForm Converter

Aplikasi web untuk mengkonversi dokumen kuesioner (PDF/DOCX) menjadi file **XLSForm (.xlsx)** yang siap di-upload ke **KoboToolbox** / ODK, menggunakan AI Claude via Sumopod API.

---

## 🗂️ Struktur Folder

```
xlsform-converter/
├── backend/
│   ├── __init__.py
│   ├── app.py              # Flask app & API endpoint
│   ├── file_parser.py      # Ekstraksi teks dari PDF/DOCX
│   ├── llm_client.py       # Koneksi ke Sumopod AI API
│   ├── xlsform_builder.py  # Build file Excel XLSForm
│   └── prompt.md           # System prompt untuk AI
├── frontend/
│   ├── templates/
│   │   └── index.html      # UI utama
│   └── static/             # CSS/JS tambahan (opsional)
├── .env.example            # Template konfigurasi
├── .gitignore
├── Dockerfile
├── Procfile                # Untuk Railway/Heroku
├── railway.toml            # Konfigurasi Railway
├── requirements.txt
├── run.py                  # Entry point lokal
├── setup.bat               # Setup otomatis Windows
├── setup.sh                # Setup otomatis Mac/Linux
└── README.md
```

---

## ⚡ Quick Start — Lokal

### Windows

```bash
# 1. Clone / extract project
# 2. Jalankan setup otomatis
setup.bat

# 3. Edit .env — isi API key
notepad .env

# 4. Jalankan
venv\Scripts\activate
python run.py
```

### Mac / Linux

```bash
# 1. Clone / extract project
# 2. Jalankan setup otomatis
chmod +x setup.sh
./setup.sh

# 3. Edit .env — isi API key
nano .env

# 4. Jalankan
source venv/bin/activate
python run.py
```

Buka browser: **http://localhost:5000**

---

## 🔑 Konfigurasi `.env`

Edit file `.env` (salin dari `.env.example`):

```env
# API Sumopod — dapatkan di https://sumopod.com/dashboard/ai/keys
SUMOPOD_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
SUMOPOD_BASE_URL=https://ai.sumopod.com/v1
SUMOPOD_MODEL=claude-sonnet-4-6

# Flask
FLASK_ENV=development
PORT=5000
SECRET_KEY=ganti-dengan-string-acak-yang-panjang
```

> **Model yang tersedia di Sumopod:** `claude-sonnet-4-6`, `gpt-4o`, `gpt-4o-mini`, `deepseek-chat`
> Untuk konversi kuesioner kompleks, **claude-sonnet-4-6** direkomendasikan.

---

## 🚂 Deploy ke Railway

### Cara 1 — Via GitHub (Rekomendasi)

1. Push project ke GitHub repository
2. Buka [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub**
3. Pilih repository ini
4. Tambahkan **Environment Variables** di Railway dashboard:
   ```
   SUMOPOD_API_KEY = sk-xxxxx
   SUMOPOD_BASE_URL = https://ai.sumopod.com/v1
   SUMOPOD_MODEL = claude-sonnet-4-6
   FLASK_ENV = production
   SECRET_KEY = string-acak-panjang
   ```
5. Railway otomatis detect `railway.toml` dan deploy

### Cara 2 — Via Railway CLI

```bash
npm install -g @railway/cli
railway login
railway init
railway up
railway variables set SUMOPOD_API_KEY=sk-xxxxx
railway variables set SUMOPOD_BASE_URL=https://ai.sumopod.com/v1
railway variables set SUMOPOD_MODEL=claude-sonnet-4-6
```

### Cara 3 — Via Docker

```bash
docker build -t xlsform-converter .
docker run -p 5000:5000 \
  -e SUMOPOD_API_KEY=sk-xxxxx \
  -e SUMOPOD_BASE_URL=https://ai.sumopod.com/v1 \
  -e SUMOPOD_MODEL=claude-sonnet-4-6 \
  xlsform-converter
```

---

## 🔌 API Endpoint

### `POST /api/convert`

Konversi file kuesioner ke XLSForm.

**Request:**
```
Content-Type: multipart/form-data
Body: file = <PDF/DOC/DOCX file>
```

**Response (sukses):**
```
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename="nama_file_xlsform.xlsx"
Body: <binary xlsx>
```

**Response (error):**
```json
{ "error": "Pesan error detail" }
```

### `GET /health`
Health check endpoint.
```json
{ "status": "ok", "version": "1.0.0" }
```

---

## 📝 Cara Kerja

1. **Upload** → User upload PDF/DOCX via browser
2. **Parse** → Server ekstrak teks dari file:
   - PDF: menggunakan `pdfminer.six` (fallback: `pypdf`)
   - DOCX: menggunakan `python-docx`
   - DOC: menggunakan `antiword` atau `docx2txt`
3. **AI Convert** → Teks dikirim ke Sumopod API (Claude) bersama system prompt XLSForm
4. **Build Excel** → Response JSON dari AI dikonversi ke file `.xlsx` dengan 3 sheet:
   - `survey` — daftar pertanyaan + logic
   - `choices` — daftar pilihan jawaban
   - `settings` — metadata form
5. **Download** → User download file `.xlsx` siap upload ke KoboToolbox

---

## ⚠️ Catatan Penting

- **Kuesioner panjang** (>100 pertanyaan): AI mungkin tidak bisa menangani semuanya sekaligus. Pertimbangkan untuk konversi per bagian.
- **File maksimal 10MB** — untuk file lebih besar, kompres dulu
- **Validasi wajib**: Selalu validasi hasil XLSForm di [XLSForm Online Validator](https://getodk.org/xlsform/) sebelum upload ke KoboToolbox
- **Timeout**: Konversi kuesioner kompleks bisa memakan waktu 1-3 menit

---

## 🛠️ Troubleshooting

| Masalah | Solusi |
|---------|--------|
| `SUMOPOD_API_KEY is not configured` | Pastikan file `.env` sudah diisi dan berada di root folder |
| `Could not extract readable content` | File mungkin berisi gambar/scan, coba OCR dulu |
| `LLM returned invalid response` | Coba lagi, atau gunakan model yang lebih kuat (`gpt-4o`) |
| `.doc file not readable` | Konversi ke `.docx` atau `.pdf` terlebih dahulu |
| Railway deploy gagal | Pastikan semua env variables sudah diset di Railway dashboard |

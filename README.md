# 🚀 SuperTools Excel Add-in — Panduan Setup

## 📦 Isi File

```
supertools-addin/
├── index.html      ← UI utama add-in
├── app.js          ← Semua logika fitur
├── manifest.xml    ← Konfigurasi yang dibaca Excel
└── README.md       ← Panduan ini
```

---

## ✅ Cara Deploy ke GitHub Pages (Tanpa Node.js)

### Langkah 1 — Buat repo GitHub
1. Buka https://github.com → klik **New Repository**
2. Nama repo: `supertools-addin`
3. Pilih **Public** → klik **Create repository**

### Langkah 2 — Upload file
1. Klik **Add file** → **Upload files**
2. Upload semua file (`index.html`, `app.js`, `manifest.xml`)
3. Klik **Commit changes**

### Langkah 3 — Aktifkan GitHub Pages
1. Buka tab **Settings** di repo
2. Klik **Pages** di sidebar kiri
3. Di bagian "Source" pilih: **Deploy from a branch**
4. Branch: **main**, Folder: **/ (root)**
5. Klik **Save**
6. Tunggu 1-2 menit → URL akan muncul:
   `https://USERNAME.github.io/supertools-addin`

### Langkah 4 — Edit manifest.xml
Ganti semua `YOUR_USERNAME` dengan username GitHub kamu:
```xml
<SourceLocation DefaultValue="https://USERNAME.github.io/supertools-addin/index.html"/>
```

### Langkah 5 — Pasang ke Excel
1. Buka **Excel Desktop**
2. Klik tab **Insert** → **Add-ins** → **My Add-ins**
3. Klik **Upload My Add-in**
4. Pilih file `manifest.xml` dari komputer
5. Klik **Upload**
6. 🎉 Tombol **SuperTools** muncul di ribbon tab Home!

---

## 🔧 Cara Pakai Fiturnya

### 📊 Tab DATA
| Fitur | Cara Pakai |
|-------|-----------|
| Statistik | Pilih range angka → klik **Hitung Statistik** |
| Duplikat | Pilih range → klik **Highlight Duplikat** |
| Kosong | Pilih range → klik **Highlight Kosong** |
| Konversi | Pilih range → pilih jenis konversi → klik **Terapkan** |

### 🎨 Tab FORMAT
| Fitur | Cara Pakai |
|-------|-----------|
| Template Style | Pilih range → klik template yang diinginkan |
| Heatmap | Pilih range angka → pilih warna → klik **Terapkan Heatmap** |
| Warna Kustom | Pilih range → pilih warna → klik **Terapkan Warna** |

### 📈 Tab CHART
| Fitur | Cara Pakai |
|-------|-----------|
| Buat Chart | Pilih data → pilih jenis → isi judul → klik **Buat Chart** |
| Rekomendasi | Pilih data → klik **Rekomendasikan Chart** |

### 🔧 Tab TOOLS
| Fitur | Cara Pakai |
|-------|-----------|
| Find & Replace | Isi teks → klik **Ganti Semua** |
| Split Kolom | Pilih satu kolom teks → pilih pemisah → klik **Split** |
| Formula Pintar | Pilih sel kosong → klik formula yang diinginkan |
| Sort | Pilih range → klik **A→Z** atau **Z→A** |

### 📤 Tab EXPORT
| Fitur | Cara Pakai |
|-------|-----------|
| Export CSV | Pilih range → atur opsi → klik **Download CSV** |
| Export JSON | Pilih range → atur opsi → klik **Download JSON** |
| Copy Clipboard | Pilih range → pilih format → klik **Salin** |

---

## ⚠️ Troubleshooting

**Add-in tidak muncul setelah upload manifest?**
→ Coba tutup dan buka ulang Excel

**Error "We can't load the app"?**
→ Pastikan GitHub Pages sudah aktif dan URL di manifest.xml sudah benar

**Fitur tidak berjalan?**
→ Pastikan memilih range/sel terlebih dahulu sebelum menekan tombol

**Excel Web tidak support?**
→ Untuk Excel Web, tambahkan domain GitHub Pages di AppDomains manifest

---

## 📝 Lisensi
MIT License — bebas digunakan dan dimodifikasi

# AutoCover Studio X7

AutoCover Studio X7 membantu membuat layout hardcover A3+ dari file cover PDF. Aplikasi ini menyediakan dua antarmuka:

- **Desktop Python/Tkinter** (`app.py`) untuk parsing metadata PDF dan otomatisasi CorelDRAW.
- **Web statis** (`dekstop_app.html`) untuk preview layout, ekspor PNG, SVG, CSV, dan salin macro/script.

## Fitur Utama

- Preset ukuran cover A4, B5, dan A5.
- Pengaturan tebal punggung dan gap/paret dalam milimeter.
- Parsing metadata dari halaman pertama PDF: kategori, judul, penulis/NIM, dan tahun.
- Render halaman pertama PDF ke PNG 300 DPI sebelum diimpor ke CorelDRAW agar hasil cover lebih stabil.
- Mode CLI non-interaktif untuk batch parsing metadata.
- Ekspor dari web: PNG layout, SVG embed cover, CSV Print Merge, VBA Macro, dan script Python.

## Kebutuhan Sistem

### Umum

- Python 3.9 atau lebih baru.
- `pip` untuk memasang dependensi.

### Untuk parsing/render PDF

- PyMuPDF (`fitz`).

### Untuk otomatisasi CorelDRAW

- Windows.
- CorelDRAW X7 atau CorelDRAW yang mendukung COM Automation.
- `pywin32`.

> Catatan: Di Linux/macOS, aplikasi tetap dapat menjalankan mode CLI untuk parsing metadata, tetapi otomatisasi CorelDRAW tidak tersedia.

## Instalasi

```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows PowerShell/CMD
python -m pip install -r requirements.txt
```

## Cara Menjalankan Aplikasi Desktop

```bash
python app.py
```

Langkah penggunaan:

1. Klik **Pilih Banyak File PDF**.
2. Pilih satu atau beberapa file cover PDF.
3. Pilih preset ukuran cover dan masukkan tebal punggung.
4. Klik tombol proses untuk membuat dokumen di CorelDRAW.

Jika aplikasi dijalankan di lingkungan tanpa layar/GUI, program otomatis beralih ke mode CLI.

## Cara Menjalankan Mode CLI

Parsing metadata tanpa CorelDRAW:

```bash
python app.py --cli --no-corel --preset B5 --spine 12 /path/ke/folder-pdf
```

Memproses beberapa file PDF secara langsung:

```bash
python app.py --cli --preset B5 --spine 12 cover-1.pdf cover-2.pdf
```

Opsi penting:

- `--cli` atau `-c`: menjalankan mode command line.
- `--no-corel`: hanya ekstrak metadata, tidak menjalankan otomatisasi CorelDRAW.
- `--preset {A4,A5,B5}`: memilih preset ukuran cover.
- `--spine <mm>`: tebal punggung buku dalam milimeter.
- `paths`: file PDF atau folder berisi PDF.

## Cara Menjalankan Versi Web Statis

Buka file berikut langsung di browser modern:

```bash
xdg-open dekstop_app.html  # Linux
# start dekstop_app.html   # Windows
# open dekstop_app.html    # macOS
```

Atau jalankan server statis lokal:

```bash
python -m http.server 8000
```

Lalu buka:

```text
http://localhost:8000/dekstop_app.html
```

> Versi web memakai CDN Tailwind CSS, Font Awesome, dan PDF.js sehingga membutuhkan koneksi internet ketika pertama dibuka.

## Validasi Layout

Layout dihitung pada lembar A3+ **297 x 430 mm**. Program akan menolak ukuran cover, spine, atau gap yang menghasilkan total lebar lebih besar dari 297 mm, atau tinggi cover lebih besar dari 430 mm.

Rumus utama:

```text
total_lebar = tebal_punggung + gap + lebar_cover
margin_kiri = (297 - total_lebar) / 2
margin_atas = (430 - tinggi_cover) / 2
```

## Troubleshooting

- **`pywin32` tidak ditemukan**: install dependensi di Windows dengan `python -m pip install -r requirements.txt`.
- **CorelDRAW gagal terhubung**: pastikan CorelDRAW sudah terpasang, mendukung COM Automation, dan coba jalankan CorelDRAW sebelum memulai script.
- **Metadata tidak akurat**: PDF hasil scan biasanya tidak memiliki teks; gunakan nama file dengan pola `Nama Penulis - Judul.pdf` sebagai fallback.
- **Versi web tidak memuat PDF**: periksa koneksi internet karena PDF.js dimuat dari CDN.

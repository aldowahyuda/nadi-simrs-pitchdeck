# Nadi SIMRS - Pitch Deck & Presentasi (Modular)

Repositori ini berisi berkas presentasi pitch deck untuk produk **Nadi SIMRS** (Sistem Informasi Manajemen Rumah Sakit & Klinik), yang disesuaikan secara modular untuk target **RSUD (Enterprise)** dan **Klinik Pratama (Core)**.

---

## 📂 Isi Folder

1. **`index.html`** (Interactive Web Pitch Deck)
   - Presentasi berbasis web interaktif premium (Dark/Light mode, penyeleksi target audiens Klinik/RSUD secara dinamis, navigasi penuh keyboard/tombol, dan mode layar penuh).
   - **Cara Membuka**: Cukup double-click file `index.html` ini pada komputer Anda untuk langsung membukanya di browser (Google Chrome, Edge, Safari, dll). Tidak memerlukan server lokal.

2. **`NADI_Presentasi.pptx`** (File PowerPoint)
   - File PowerPoint siap pakai yang dihasilkan secara otomatis menggunakan skema warna modern (Slate & Teal).
   - Cocok untuk presentasi konvensional atau dikirim langsung ke client.

3. **`generate_pptx.js`** (Skrip Generator)
   - Skrip Node.js yang digunakan untuk merancang dan mengekspor slide presentasi ke file `.pptx` secara programatik menggunakan pustaka `pptxgenjs`.

---

## 🛠️ Cara Mengedit & Mengembangkan

### 1. Mengubah Konten Slide Web (`index.html`)
- Buka file `index.html` dengan text editor pilihan Anda (misalnya VS Code atau Notepad).
- Cari elemen dengan kelas `.slide` dan sesuaikan teks atau ikon SVG di dalamnya.

### 2. Memperbarui File PowerPoint (`NADI_Presentasi.pptx`)
Jika Anda mengedit skrip `generate_pptx.js` untuk merubah slide PowerPoint, Anda dapat mengekspor ulang dengan cara:
1. Pastikan Node.js terinstal di komputer Anda.
2. Buka terminal di folder ini dan instal dependensi:
   ```bash
   npm install
   ```
3. Jalankan skrip generator:
   ```bash
   node generate_pptx.js
   ```
4. File `NADI_Presentasi.pptx` baru akan terbuat secara otomatis menggantikan file yang lama.

# Dashboard Pertumbuhan Inklusif — Kabupaten/Kota

Dashboard Streamlit untuk menganalisis PDRB sektoral, indikator sosial-ekonomi, dan
kontribusi tiap sektor terhadap **Pertumbuhan Inklusif** kabupaten/kota.

## Cara menjalankan

1. Install dependensi (sekali saja):
   ```bash
   pip install -r requirements.txt
   ```
2. Jalankan dashboard:
   ```bash
   streamlit run app.py
   ```
3. Browser akan terbuka otomatis (biasanya di `http://localhost:8501`).
4. Di sidebar kiri, klik **Browse files** dan unggah file dataset Anda:
   `OKK_Panel_PDRB_PDRB_KAB-Kota_Sulsel_3_Inclusive_Growth.xlsx`
   (atau versi `.csv`-nya). Sebelum diunggah, dashboard menampilkan data contoh
   (Bantaeng & Barru) agar Anda bisa langsung melihat tampilannya.

## Struktur dataset yang didukung

Kolom wajib: `kabupaten`, `tahun`, `PDRB`, `Pertumbuhan_Inklusif`, `Kategori_Inklusif`.
Kolom opsional (akan otomatis dipakai jika ada): 17 kolom sektor (`A_Pertanian_KP` … `R_S_T_U_Jasa_Lain`),
`ipm`, `miskin`, `pertumbuhan`, `gini`, `unemploy`, `pop_ribujiwa`, `pdrb_perkap`, `pct_formal`, `pend_miskin`.

Format angka dengan koma desimal ala Indonesia (mis. `1040,57`) otomatis dikonversi ke angka.

## Fitur

- **Ringkasan**: KPI utama (PDRB, IPM, Kemiskinan, Gini, Indeks Pertumbuhan Inklusif) beserta
  perubahan tahun-ke-tahun, tren gabungan PDRB & Pertumbuhan Inklusif, distribusi kategori
  inklusif, dan peringkat kabupaten/kota.
- **Kontribusi Sektoral** *(fitur utama)*:
  - Komposisi sektoral PDRB dari waktu ke waktu (area chart) per kabupaten atau rata-rata semua wilayah.
  - Kontribusi tiap sektor terhadap PDRB pada tahun terakhir.
  - **Korelasi pertumbuhan tiap sektor dengan Indeks Pertumbuhan Inklusif** — menjawab
    langsung sektor mana yang paling berasosiasi positif/negatif dengan pertumbuhan inklusif.
  - Peta korelasi (heatmap) antar sektor dan Pertumbuhan Inklusif.
- **Indikator Sosial**: tren IPM, kemiskinan, gini, pengangguran, dsb per kabupaten, serta
  scatter plot Gini vs IPM dan Pengangguran vs Kemiskinan.
- **Perbandingan Wilayah**: heatmap indikator per kabupaten × tahun, dan scatter PDRB vs
  Pertumbuhan Inklusif dengan garis tren.
- **Data**: tabel data terfilter + unduh CSV.

Semua tab mengikuti filter di sidebar (kabupaten, rentang tahun, kategori inklusif).

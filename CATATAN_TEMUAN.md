# Catatan & Temuan — Cancellation & Return Analysis

Proyek end-to-end data analyst: pola order yang **gagal menjadi revenue** (Batal + Retur Penuh) pada data penjualan marketplace, Jan 2024 – Nov 2025 (23 bulan).

- **Teknik:** descriptive + diagnostic drill-down (Pareto, cross-tab, chi-square). Bukan root cause analysis formal. Semua temuan dibingkai sebagai **faktor yang berasosiasi**, bukan penyebab.
- **Label keyakinan:** [High] = dihitung langsung & stabil; [Medium] = konsisten tapi ada asumsi/sampel kecil; [Low] = indikasi, belum diuji.
- Terakhir diperbarui: 22 Okt 2026.

---

## Insight Utama (Ringkasan Eksekutif)

Dari 20.154 order settled, **14,07% (2.835 order) gagal menjadi revenue**. Lima insight berikut berlaku untuk seluruh periode (23 bulan) dan bersifat asosiasi, bukan sebab-akibat.

### 1. Kegagalan terbesar bukan penolakan produk, melainkan pembeli yang ingin *mengubah* pesanan [High]
- 38,6% pembatalan (1.056 order) adalah *ubah pesanan*, *ubah alamat*, atau *ubah voucher*. 26,4% lagi pembeli tidak jadi membeli (berubah pikiran, menemukan yang lebih murah, dll.).
- **Artinya:** kelompok terbesar adalah pembeli yang masih berniat membeli tetapi perlu mengoreksi pesanan. Kemungkinan mereka membatalkan lalu memesan ulang. Ini **tidak bisa dibuktikan** karena identitas pembeli tidak konsisten [Low].
- **Rekomendasi:** cek dulu apakah platform mengizinkan edit alamat/varian sebelum dikirim. Kalau tidak bisa, kirim pesan konfirmasi alamat & varian segera setelah order masuk. Kebijakan ini perlu diuji dengan membandingkan rate sebelum & sesudah.

### 2. Masalah "belum dibayar" ada di Online Payment, bukan di gerai [High]
- Dari 431 order batal karena *belum dibayar*, **56,4% (243) berasal dari Online Payment**; Indomaret + Alfamart hanya 5,8% (25).
- Tanpa alasan *belum dibayar*, rate Online Payment turun dari 18,93% menjadi **11,04%**, di bawah rata-rata. Artinya, kelebihan rate-nya sepenuhnya dari order yang tidak dibayar.
- Rate tinggi Indomaret (50%) dan Alfamart (45,7%) juga hampir seluruhnya dari *belum dibayar* (71% & 63% pembatalannya). Tanpa alasan ini, rate-nya 14,3% & 17,1%. Namun volumenya kecil (77 order), jadi dampak absolutnya kecil.
- **Rekomendasi:** prioritaskan pengingat pembayaran untuk Online Payment (volume 3.079 order), bukan untuk gerai.

### 3. COD menyumbang lebih dari separuh order gagal walaupun rate-nya di bawah rata-rata [High]
- Rate COD 13,83% (di bawah rata-rata), tetapi karena volumenya 55% order, **COD menyumbang 54,4% dari seluruh order gagal**.
- *Pengiriman gagal* hampir eksklusif terjadi di COD (243 dari 247). Penyebab di lapangan (pembeli menolak paket, tidak bisa dihubungi, dll.) tidak tercatat di data [Low untuk mekanismenya].
- **Rekomendasi:** jangan simpulkan "COD aman" hanya dari rate-nya. Untuk mengurangi jumlah absolut order gagal, COD tetap target terbesar. Contoh tindakan: konfirmasi order COD sebelum dikirim.

### 4. Sekitar 8% pembatalan ada di kendali penjual [High]
- 223 order (8,1% pembatalan): penjual telat kirim (113), telat atur pengiriman (63), tidak responsif (33), produk habis (14).
- 176 di antaranya tercatat sebagai pembatalan *sistem*, jadi tersembunyi kalau hanya melihat sumber pembatalan.
- **Rekomendasi:** ini *quick win* operasional (SLA pengiriman, sinkron stok, respons chat) karena tidak bergantung pada perilaku pembeli.

### 5. Periode Mar–Jun 2025 memburuk, lalu membaik sejak Jul 2025 [High untuk polanya, penyebab belum diketahui]
- Rate naik ke 16–22% di Mar–Jun 2025 (puncak Mei 22,10%), lalu turun ke 9,55–12,37% sejak Jul 2025.
- Penyebabnya belum dianalisis (langkah 4). Hipotesis yang perlu diuji: perubahan komposisi metode bayar, kategori, atau masalah pengiriman penjual di periode itu.

### Catatan kategori [High]
- **Rak / Rak Serbaguna** adalah prioritas kategori: rate 31,87% dengan volume besar (568 order), sehingga menghasilkan **±101 order gagal lebih banyak** daripada jika rate-nya setara rata-rata. Pengolah Bumbu/Sayur punya rate tertinggi (32,26%) tetapi kelebihannya hanya ±28 order.
- Data tidak memuat alasan di level produk, jadi penyebab per kategori (deskripsi, ukuran, ongkir barang besar, dll.) masih dugaan [Low].

---

## 1. Alur Proyek

| Tahap | Alat | File | Status |
|---|---|---|---|
| Raw → Clean | Python | `data_cleaning.ipynb` → `Clean_data/master_clean_orders_FIXED.xlsx` | ✅ |
| Clean → Dataset analisis | Python | `02_analysis.ipynb` → `Clean_data/analysis_dataset.xlsx` | ✅ |
| Dashboard otomatis | Python + Excel COM | `build_dashboard_auto.py` → `Dashboard/dashboard_cancel_return_auto.xlsx` | ✅ |
| Dashboard manual (latihan) | Excel | `Dashboard/dashboard_cancel_return.xlsx` | ⏳ Langkah 1/7 |
| Pendalaman Mei 2025 | Python | — | ⏳ |
| Laporan insight & rekomendasi | — | — | ⏳ |

**Prinsip:** semua transformasi data dikerjakan di Python (terdokumentasi, bisa dijalankan ulang); Excel hanya mengagregasi dan menampilkan. Data masuk Excel lewat Power Query, sehingga dashboard cukup di-*Refresh All*.

---

## 2. Data & Definisi

- **Master:** 20.404 order × 23 kolom, 1 baris = 1 `order_id` (tidak ada duplikat).
- **Order settled** (`is_settled = 1`): 20.154. Sebanyak 250 order *Dalam Proses* (semuanya Nov 2025) dikeluarkan dari denominator.
- **Gagal-Revenue Rate** = (Batal + Retur Penuh) / order settled.
- **Komposisi status efektif:** Selesai 17.298 · Batal 2.738 · Dalam Proses 250 · Retur Penuh 97 · Retur Parsial 21.
- **Basis berbeda per metrik:**
  - KPI & rate → 2.835 order gagal dari 20.154 settled.
  - Sumber & alasan pembatalan → 2.738 order **batal** (retur penuh tidak punya alasan pembatalan).

---

## 3. Hasil Verifikasi Ulang

Semua angka di brief awal dihitung ulang dari `master_clean_orders_FIXED.xlsx`.

### Cocok persis [High]
- Overall rate **14,07%** (2.835 / 20.154).
- Tren bulanan 23 bulan, termasuk Nov 2025 **12,37%** (akan turun palsu ke 9,64% bila Dalam Proses ikut denominator).
- Sumber pembatalan: Pembeli **67,35%**, Sistem Otomatis **32,14%**, Penjual Manual **0,51%**.
- Rate per metode pembayaran.
- 724 order artefak opsi pengiriman (3,59%), semuanya 100% gagal.
- Top 5 kategori rawan.

### Koreksi
| Klaim di brief | Hasil cek |
|---|---|
| File di `Clean_data/master_clean_orders_FIXED.xlsx` | Awalnya hanya ada di root sebagai `master_clean_orders_FIXED (1).xlsx`. Sudah disalin ke `Clean_data/` (hash identik). |
| Penjual Manual mulai **Mar 2025** | Mulai **Apr 2025** (14 order, semuanya "Produk habis"). |
| Mei 2025 = "anomali" | 22,10% benar, tetapi Mar (18,47%), Apr (17,96%), Jun (16,26%) juga tinggi → **puncak dari periode tinggi Mar–Jun 2025**, bukan lonjakan tunggal. |
| Volume "dua kali lipat sejak Jul 2024" | Lompatan Jun → Jul 2024 benar (697 → 1.270), tetapi tidak bertahan: Mar–Jul 2025 turun ke 615–824 order/bulan. |
| Pareto top 5 = 78,96% | Benar untuk label mentah, tetapi ada label sinonim (lihat 4.3). Setelah digabung: **85,46%**. |
| Chi-square opsi pengiriman = 87,67 | Benar, **tetapi hanya untuk opsi n ≥ 30** (dof 15). Tanpa filter: χ² = 123,46, dan 31 sel expected < 5, jadi uji tidak valid. |

### Peringatan file
`Clean_data/master_clean_orders.xlsx` (versi lama) **jangan dipakai**: 579 baris `cancel_source` berbeda dan `is_settled` belum terisi benar.

---

## 4. Temuan Penting

### 4.1 Tren bulanan
- Rate bergerak di rentang **9,55% (Agu 2025) – 22,10% (Mei 2025)**. [High]
- **Periode tinggi Mar–Jun 2025** (16–22%), lalu membaik sejak Jul 2025 (12,14% → 9,55% → 10,70% → 10,12%) dan naik ke 12,37% di Nov 2025. [High]
- Per tahun: 2024 = 14,38% (1.472 / 10.236); 2025 = 13,74% (1.363 / 9.918). [High]

### 4.2 Sumber pembatalan
- **Pembeli 67,35%** (1.844), **Sistem Otomatis 32,14%** (880), **Penjual Manual 0,51%** (14). [High]
- **"Sistem Otomatis" tidak sepenuhnya di luar kendali penjual:** 176 dari 880 (20,0%) adalah *Penjual telat kirim* (113) + *Penjual telat atur kirim* (63). Ini sumber gagal-revenue yang bisa dikendalikan penjual. [High]

### 4.3 Alasan pembatalan (Pareto, label sinonim digabung)
Pasangan label sinonim muncul bersamaan di 21–23 bulan (kemungkinan beda bahasa/versi aplikasi pembeli, bukan pergantian versi di satu tanggal). Digabung 18 → 16 label; "Lainnya" **tidak** digabung dengan "berubah pikiran" karena itu kategori sisa.

| # | Alasan | Sumber | n | % | Kumulatif |
|---|---|---|---|---|---|
| 1 | Ubah pesanan | Pembeli | 644 | 23,52% | 23,52% |
| 2 | Berubah pikiran | Pembeli | 639 | 23,34% | 46,86% |
| 3 | Belum dibayar | Sistem | 431 | 15,74% | 62,60% |
| 4 | Ubah alamat | Pembeli | 379 | 13,84% | 76,44% |
| 5 | Pengiriman gagal | Sistem | 247 | 9,02% | 85,46% |

- Tiga alasan (ubah pesanan, berubah pikiran, ubah alamat) = **90,1% pembatalan oleh pembeli**. [High]
- "Ubah pesanan" dan "ubah alamat" kemungkinan diikuti pemesanan ulang (bukan kehilangan penjualan), tetapi **tidak bisa dibuktikan** karena identitas pembeli tidak konsisten di 23 bulan. [Low]
- "Belum dibayar" = **49,0% pembatalan sistem** (431 / 880). [High]

### 4.4 Metode pembayaran (n ≥ 30; χ² = 202,00, dof 7, p < 0,001)
| Metode | n settled | Rate |
|---|---|---|
| Indomaret/i.Saku | 42 | **50,00%** |
| Alfamart/Alfamidi/Dan+Dan | 35 | **45,71%** |
| Online Payment | 3.079 | 18,93% |
| SPayLater | 1.443 | 15,59% |
| Kartu Kredit/Debit | 222 | 14,41% |
| COD (Bayar di Tempat) | 11.160 | 13,83% |
| SeaBank Bayar Instan | 577 | 13,17% |
| Saldo ShopeePay | 3.571 | 9,41% |

- Bayar tunai di gerai (Indomaret, Alfamart) punya rate jauh di atas rata-rata, tetapi **n kecil** (42 & 35). [Medium]
- **Sudah diuji silang:** rate tinggi di gerai memang didorong *belum dibayar* (71% & 63% pembatalannya), tetapi **volume *belum dibayar* terbesar ada di Online Payment** (243 dari 431 = 56,4%). Lihat Insight Utama #2. [High]
- Rate tanpa *belum dibayar*: Online Payment 11,04%, SPayLater 11,09%, Indomaret 14,29%, Alfamart 17,14%, COD 13,83% (tidak berubah), Saldo ShopeePay 7,45%.
- COD, metode dengan volume terbesar (55% order settled), justru **di bawah rata-rata**. [High]
- Saldo ShopeePay terendah (9,41%). [High]
- 4 metode dibuang karena n < 30 (total 25 order): Pembayaran dibebaskan 18, BCA OneKlik 3, Cicilan Kartu Kredit 3, Mitra Shopee 1.

### 4.5 Kategori produk (order 1 kategori, kategori n ≥ 30)
27 dari 37 kategori tunggal memenuhi syarat; 18.582 order settled masuk analisis.

| Kategori | n | Rate |
|---|---|---|
| Pengolah Bumbu / Sayur | 155 | 32,26% |
| Rak / Rak Serbaguna | 568 | 31,87% |
| Stempel / Alat Kantor | 117 | 25,64% |
| Piring | 141 | 24,82% |
| Plastik / Wadah Plastik | 107 | 24,30% |

- Untuk **prioritas tindakan**, Rak Serbaguna lebih penting: 181 order gagal, 3,6× lebih banyak daripada Pengolah Bumbu (50) dengan rate hampir sama. [High]

### 4.6 Opsi pengiriman — temuan terlemah, tidak masuk dashboard [Low]
- Setelah 724 artefak dikeluarkan: Hemat Kargo-J&T Economy 22,29%, Agen SPX Express 21,74%, J&T Cargo 20,00% vs dominan Hemat Kargo-SPX Hemat 11,28% (12.088 order).
- **Masalah utama:** kurir yang sama tercatat dengan nama berbeda di periode berbeda.
  - `Kargo-J&T Cargo` (Jan 2024 – Jun 2025): 5,71%.
  - `J&T Cargo` (Apr – Nov 2025): 20,00%.
  - `Hemat Kargo-J&T Cargo` (Jun – Nov 2025): 11,27%.
  - **Digabung: 10,33%** (28 / 271), di bawah rata-rata.
- Format "nama kurir saja" hampir tidak ada sebelum Nov 2024, lalu naik ke 40–100 order/bulan → perbandingan antar-opsi **tercampur efek waktu**.

---

## 5. Batasan Metodologi

1. **Kerugian Rupiah tidak dihitung.** `total_pembayaran_idr` tidak konsisten setelah retur (retur penuh kadang 0, kadang tetap nilai awal; retur parsial tidak disesuaikan). Kolom ini sengaja tidak dibawa ke dataset dashboard.
2. **`waktu_pesanan` sebagian proxy** (jadwal kirim batch) di Des 2024 dan sebagian 2025. Aman untuk agregasi bulanan (`bulan_file`), **tidak** untuk analisis jam/hari. Kolom ini juga kosong di 237 baris.
3. **Artefak opsi pengiriman:** 724 order dengan opsi generik tanpa kurir selalu Batal (efek pencatatan). Tetap dihitung di KPI, dikecualikan dari analisis opsi pengiriman.
4. **Retur sangat sedikit** (118 order: 97 penuh, 21 parsial); kolom retur Jan 2024 kosong total. Retur hanya catatan sekunder.
5. **Kategori** hanya dianalisis untuk order 1 kategori dan n ≥ 30.
6. **Label alasan sinonim digabung** (18 → 16). Pemetaan lengkap ada di `02_analysis.ipynb` bagian 2.
7. **RFM ditunda:** kolom `Username (Pembeli)` hanya ada di **1 dari 24 file mentah** (Sep 2024), jadi identitas pembeli tidak tersedia untuk 23 bulan. (Brief awal menyebut "mulai Sep 2024"; hasil cek menunjukkan hanya bulan itu.)

---

## 6. Keputusan Desain Dashboard

| Keputusan | Alasan |
|---|---|
| Rate via calculated field `rate_gagal = is_gagal_revenue / is_settled` | Excel menghitung SUM/SUM, jadi order Dalam Proses otomatis keluar dari denominator. Aman karena tidak ada order gagal yang belum settled (diverifikasi). |
| KPI "Total Order Settled" = *Sum of is_settled*, bukan *Count of order_id* | Count memberi 20.404, bukan 20.154. |
| Semua pivot dari **1 pivot cache** | Syarat agar 1 slicer bisa mengendalikan beberapa pivot. |
| Slicer **Bulan** hanya ke KPI, Sumber, Alasan | Rate per grup tidak stabil pada n kecil (satu bulan). |
| Chart **Tren** tidak ikut slicer Tahun & Bulan | Konteks 23 bulan harus utuh; rata-rata tetap sebanding. |
| Chart **Metode** tidak ikut slicer Metode | Fungsinya membandingkan semua metode. |
| Garis rata-rata = rata-rata **tertimbang** (grand total) | Bukan rata-rata dari 23 rate bulanan. |
| % alasan dihitung terhadap **total order batal** | Grand total pivot ikut terpotong filter Top 5, sehingga % akan salah bila memakai grand total itu. |
| Label chart **tanpa** "(n=…)"; n hanya ada di sheet `chart_data` | Permintaan user: label bersih & konsisten. Syarat n ≥ 30 ditulis di subjudul chart. |
| Subjudul chart seragam: `Basis: … · syarat … · slicer yang tidak berlaku`, judul rata kiri | Konsistensi penulisan antar-chart. |
| Latar halaman abu kebiruan `#D5DCE6`, kartu putih dengan aksen warna, footnote navy | Kontras kartu & teks lebih jelas, senada dengan header navy. |
| Slicer Bulan hanya nama bulan (Jan…Des), tanpa nomor | Kolom `bulan` disimpan sebagai tanggal dummy (tgl 15, tahun 2000) dan diformat `[$-421]mmm`, sehingga slicer urut kronologis dan tampil dalam bahasa Indonesia di komputer mana pun (tidak bergantung custom list). |
| Panel **Insight & Keputusan** (6 baris) di bawah chart, sebelum footnote | Alur baca: bukti (chart) → kesimpulan → keputusan. Teks statis, diberi label "berlaku seluruh periode · tidak ikut slicer". Angka mengikuti format locale Excel (14.07%) agar konsisten dengan angka lain di dashboard. |
| Warna metode: pink = di atas rata-rata, teal = di bawah | Aturan warna di mockup tidak konsisten (COD di bawah rata-rata tetapi oranye). |
| Judul "Cancel Rate per Metode" → **"Gagal-Revenue Rate per Metode"** | Metriknya juga mencakup retur penuh. |
| Footnote: Rp loss · basis settled · label digabung | Poin artefak pengiriman dipindah ke laporan karena dashboard tidak punya chart pengiriman. |

**Perilaku dashboard yang perlu diketahui pembaca:**
- Slicer Bulan tanpa Tahun **menggabungkan dua tahun** ("Mei" = Mei 2024 + Mei 2025 = 1.439 order, rate 19,32%).
- Kelayakan n ≥ 30 dihitung global, jadi saat difilter (mis. Tahun = 2024) kategori dengan n kecil bisa muncul di Top 5 (contoh: Aksesoris ID Card n = 25). Label chart tidak menampilkan n, jadi saat membaca hasil filter, cek kolom `n_settled` di sheet `chart_data`.
- Indomaret (n = 42) dan Alfamart (n = 35) punya sampel kecil; sebutkan ini secara lisan/di laporan saat mempresentasikan rate 50% & 45,7%.
- Path Power Query absolut; ubah di *Queries & Connections* bila folder dipindah.

---

## 7. Pertanyaan Terbuka untuk Pendalaman (Langkah 4)

1. **Apa yang berbeda di Mar–Jun 2025?** Cek kombinasi metode bayar, kategori, dan alasan pada periode itu dibanding periode lain.
2. **"Penjual telat kirim"** masuk Top 5 alasan saat slicer Bulan = Mei (10,6%, gabungan Mei 2024 + 2025), padahal secara keseluruhan hanya 4,1%. Perlu dipisah per tahun dan diuji. [Low]
3. ~~Apakah rate tinggi Indomaret/Alfamart didominasi "Belum dibayar"?~~ Terjawab: ya, tetapi volume terbesar *belum dibayar* ada di Online Payment (Insight Utama #2). Pertanyaan lanjutan: apakah *belum dibayar* di Online Payment ikut naik pada Mar–Jun 2025?
4. Apakah tingginya rate Rak Serbaguna terkait periode tertentu atau stabil sepanjang 23 bulan?

---

## 8. Catatan Housekeeping

- `prompt_sesi_baru_cli.md` masih memuat klaim lama (Penjual Manual mulai Mar 2025, top 5 alasan 78,96%, label "anomali" Mei 2025). Gunakan file ini sebagai acuan terbaru.
- Nama folder `RFM E-Commerce` tidak sesuai isi proyek; ganti di akhir proyek (setelah path Power Query diperbarui).
- Library yang dipasang di sesi ini: `openpyxl`, `nbformat`, `nbclient`.

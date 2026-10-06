# PaySim Fraud Audit (Otomasi Analisis & Deteksi Kecurangan Keuangan)

[![Python Version](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://www.python.org/)
[![Dataset HF](https://img.shields.io/badge/Dataset-HuggingFace-orange.svg)](https://huggingface.co/datasets/purulalwani/Synthetic-Financial-Datasets-For-Fraud-Detection)
[![License MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

> **Otomasi sistem audit berbasis aturan (rule-based) untuk mengidentifikasi anomali dan transaksi keuangan mencurigakan pada simulasi data transaksi perbankan PaySim.**

---

## 📌 Tentang Project

Dalam era digitalisasi industri pembiayaan dan perbankan, volume transaksi harian melonjak sangat pesat. Metode audit sampling tradisional sering kali gagal menangkap kecurangan (*fraud*) karena keterbatasan jangkauan manual. Project ini dirancang untuk mendemonstrasikan bagaimana **Otomasi Audit Keuangan** dapat menyaring jutaan baris data keuangan dengan cepat guna mendeteksi kecurangan.

Menggunakan dataset **PaySim** (data sintetis yang merepresentasikan transaksi keuangan seluler riil sebanyak 6,3 juta baris data), project ini menerapkan **5 aturan penapisan investigatif** untuk mencari ketidaksesuaian saldo, pola pencilan (*outliers*), mule account, dan kegagalan pengendalian internal. Hasil akhirnya adalah laporan audit interaktif berstandar eksekutif dalam bentuk file Excel dengan visualisasi grafik otomatis.

Project ini dibuat oleh **Rizal Taufik Rifaldi** sebagai portofolio pemahaman audit berbasis data (*data-driven auditing*) untuk proses lamaran posisi **Audit Junior di Nusantara Sakti Group**.

---

## 🛠️ Deteksi yang Diimplementasikan

Berikut adalah 5 aturan deteksi (*detection rules*) yang dirancang untuk menguji keandalan transaksi dan kepatuhan sistem kontrol:

| Nama Aturan | Deskripsi Singkat Anomali | Kolom yang Dianalisis |
| :--- | :--- | :--- |
| **Ketidaksesuaian Saldo<br>*(Balance Mismatch)*** | Menyaring transaksi `TRANSFER` dan `CASH_OUT` di mana saldo pengirim tidak terpotong tepat sebesar nominal transaksi (selisih > 1 unit). Menunjukkan celah sistematis atau manipulasi pembukuan. | `oldbalanceOrg`, `newbalanceOrig`, `amount`, `type` |
| **Transaksi Nominal Besar<br>*(Large Transactions Outliers)*** | Mengidentifikasi transaksi pencilan yang nilainya jauh melampaui batas wajar rata-rata berdasarkan Z-Score statistik (`amount > mean + 3 * std`) per jenis transaksi. | `amount`, `type` |
| **Saldo Pengirim Terkuras<br>*(Zero Balance Origin)*** | Menyaring transaksi yang menguras habis saldo pengirim hingga tepat `0`, padahal saldo awal bernilai positif, namun transaksi tersebut terlewat (`isFraud == 0`). | `oldbalanceOrg`, `newbalanceOrig`, `isFraud` |
| **Akun Penerima Mencurigakan<br>*(Suspicious Dest Accounts - Mule)*** | Mengidentifikasi akun penerima (`nameDest`) yang menerima dana dari $\ge 4$ pengirim berbeda dalam kurun waktu sangat sempit (24 jam/step). Terindikasi sebagai akun penampung uang kecurangan. | `nameDest`, `step` |
| **Fraud Lolos Flag Sistem<br>*(Missed by System Audit)*** | Menyaring kasus transaksi fraud yang terbukti terjadi (`isFraud == 1`) namun gagal dideteksi atau ditandai oleh sistem bawaan (`isFlaggedFraud == 0`). Berguna menguji efektivitas sistem pengendalian intern. | `isFraud`, `isFlaggedFraud` |

---

## 📂 Struktur Folder

```text
portofolio-fraud-detection/
├── .gitignore
├── requirements.txt
├── README.md
├── setup_dataset.py           # Script unduh dataset, pembersihan, & stratified sampling (50k baris)
├── fraud_detector.py          # Script eksekusi 5 aturan deteksi kecurangan
├── generate_excel_report.py   # Script pembuatan laporan Excel profesional 4 sheet
├── data/                      # Folder penyimpanan sampel data (contoh disertakan)
│   └── transactions_sample.csv
└── report/                    # Folder laporan hasil audit (contoh disertakan)
    ├── fraud_findings.csv     # Daftar semua baris temuan rinci
    ├── summary.txt            # Ringkasan ringkas eksekusi audit dalam format teks
    └── Laporan_Audit_Fraud_PaySim.xlsx # Laporan Excel interaktif dengan Dashboard
```

---

## 🚀 Cara Menjalankan

Ikuti langkah-langkah berikut untuk mereplikasi analisis secara lokal pada komputer Anda:

1. **Clone repositori ini:**
   ```bash
   git clone https://github.com/rizal-taufik-rifaldi/portofolio-fraud-detection.git
   cd portofolio-fraud-detection
   ```

2. **Buat dan aktifkan virtual environment Python:**
   ```bash
   python -m venv venv
   # Di macOS/Linux:
   source venv/bin/activate
   # Di Windows:
   venv\Scripts\activate
   ```

3. **Instal seluruh dependensi:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Opsional: unduh ulang dan siapkan data:**
   Sampel transaksi sudah tersedia di folder `data/`. Lewati langkah ini untuk memakai sampel yang disertakan.
   Script ini mengunduh dataset asli dari Hugging Face dan mengambil sample berstrata 50.000 baris berdasarkan proporsi transaksi asli agar ringan dijalankan secara lokal.
   ```bash
   python setup_dataset.py
   ```

5. **Jalankan modul detektor kecurangan:**
   Script ini menjalankan 5 filter audit dan menampilkan waktu eksekusi serta jumlah temuan ke konsol.
   ```bash
   python fraud_detector.py
   ```

6. **Hasilkan laporan Excel profesional:**
   Script ini menyatukan hasil temuan dan data historis ke dalam bentuk laporan Excel visual.
   ```bash
   python generate_excel_report.py
   ```

---

## 📊 Contoh Output

### 1. Cuplikan Visual Dashboard Excel (`report/Laporan_Audit_Fraud_PaySim.xlsx`)
Laporan Excel dirancang menggunakan font premium *Segoe UI*, skema warna *Navy Corporate*, pembatas kolom tipis, format desimal/ribuan otomatis, dan kolom bantu perhitungan yang disembunyikan (*hidden columns*) agar rapi.

* **Sheet 1 (Data Transaksi)**: 1.000 baris sample transaksi dengan penanda merah muda jika transaksi terbukti fraud (`isFraud == 1`).
* **Sheet 2 (Temuan Fraud)**: Daftar lengkap 32.228 deteksi yang terpicu dengan warna latar berbeda per jenis aturan untuk klasifikasi cepat auditor.
* **Sheet 3 (Ringkasan per Tipe)**: Pivot table ringkasan total nominal beserta rasio fraud per tipe transaksi dilengkapi grafik batang & pie chart.
* **Sheet 4 (Dashboard Eksekutif)**: Menyajikan 5 kartu KPI utama, 3 temuan audit dinamis, serta grafik tren transaksi harian & sebaran fraud per tipe.

[Unduh contoh laporan Excel](report/Laporan_Audit_Fraud_PaySim.xlsx)

### 2. Output Eksekusi Terminal (`fraud_detector.py`)
```text
Reading transaction data from 'data/transactions_sample.csv'...
Loaded 50,000 transactions successfully.

================================================================================
RUNNING AUDIT RULE-BASED DETECTIONS
================================================================================

Running: Balance Mismatch...
Logic  : oldbalanceOrg - amount != newbalanceOrig (TRANSFER & CASH_OUT only)
Time   : 0.003647 seconds
Found  : 19,576 suspicious transactions (39.15%)
--------------------------------------------------------------------------------

Running: Large Transactions...
Logic  : amount > mean + 3*std per transaction type
Time   : 0.006255 seconds
Found  : 532 suspicious transactions (1.06%)
--------------------------------------------------------------------------------

Running: Zero Balance Origin...
Logic  : newbalanceOrig == 0 AND oldbalanceOrg > 0 AND isFraud == 0
Time   : 0.000989 seconds
Found  : 12,063 suspicious transactions (24.13%)
--------------------------------------------------------------------------------

Running: Suspicious Destination Accounts...
Logic  : nameDest receiving > 3 transactions in 24 hours
Time   : 0.045694 seconds
Found  : 4 suspicious transactions (0.01%)
--------------------------------------------------------------------------------

Running: Missed by System...
Logic  : isFraud == 1 AND isFlaggedFraud == 0
Time   : 0.000308 seconds
Found  : 53 suspicious transactions (0.11%)
--------------------------------------------------------------------------------

================================================================================
AUDIT DETECTION SUMMARY REPORT
================================================================================
Source File             : data/transactions_sample.csv
Total Transactions      : 50,000
Total Unique Flagged    : 22,556 (45.11%)
Multi-Rule Overlaps     : 9,672
--------------------------------------------------------------------------------
Rule Name                           | Count    | % of Data  | Time (sec)
--------------------------------------------------------------------------------
Balance Mismatch                    | 19,576   | 39.15    % | 0.003647
Large Transactions                  | 532      | 1.06     % | 0.006255
Zero Balance Origin                 | 12,063   | 24.13    % | 0.000989
Suspicious Destination Accounts     | 4        | 0.01     % | 0.045694
Missed by System                    | 53       | 0.11     % | 0.000308
--------------------------------------------------------------------------------
```

---

## 💼 Relevansi dengan Peran Audit Junior

Project otomasi ini merefleksikan kesiapan kerja dan keahlian praktis yang dicari pada seorang **Audit Junior** di **Nusantara Sakti Group**:

1. **Kemampuan Analitis & Pengolahan Data (*Data Analytics Capability*)**: Mampu mengolah ribuan sampel data transaksi keuangan secara terstratifikasi untuk menganalisis risiko fraud secara menyeluruh, melampaui kemampuan audit manual konvensional.
2. **Pengujian Efektivitas Pengendalian Intern (*Internal Control Testing*)**: Aturan "Missed by System" memetakan kegagalan kontrol internal (flagging sistem yang tidak sensitif) versus kejadian fraud riil. Ini mencerminkan pemikiran kritis auditor dalam mengevaluasi kelemahan sistem operasional perusahaan.
3. **Penyajian Laporan Berstandar Eksekutif (*Professional Reporting*)**: Menyadari pentingnya komunikasi audit yang efektif. Hasil visualisasi dashboard Excel yang bersih, ringkas, dan fokus pada temuan utama memudahkan auditor junior menyajikan laporan kepada manajer audit maupun jajaran direksi.

---

## 💾 Informasi Dataset

* **Kredit Dataset**: Puru Lalwani di Hugging Face ([Hubungkan ke Dataset](https://huggingface.co/datasets/purulalwani/Synthetic-Financial-Datasets-For-Fraud-Detection)).
* **Deskripsi Singkat PaySim**: PaySim adalah simulator transaksi finansial sintetis berbasis agen yang merepresentasikan pola operasional uang seluler (*mobile money*). Simulator ini menghasilkan data transaksi berlabel fraud berlandaskan transaksi logis dari penyedia layanan keuangan sungguhan, menjadikannya standar industri untuk riset deteksi fraud tanpa membocorkan data pribadi nasabah asli.

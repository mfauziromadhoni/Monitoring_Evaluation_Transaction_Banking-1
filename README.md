# 🏦 Digital Banking Transaction Monitoring — Excel Data Cleaning
### Bank Jatim Portfolio Series (Part 1 of 2) | Microsoft Excel

---

## Overview

This project simulates the data cleaning workflow performed by a **Monitoring & Evaluation (M&E) team** in a digital banking environment. Raw transaction data from multiple digital channels is cleaned, standardized, and prepared for further analysis and dashboard visualization.

> *Part 1 of 2 — Data cleaning in Excel. Dashboard visualization continues in Part 2 (Power BI).*

---

## Dataset

| Attribute | Detail |
|---|---|
| Source | Simulated Digital Banking Transaction Data |
| Records | 730 raw transactions |
| Period | January — December 2024 |
| Channels | Mobile Banking, Internet Banking, ATM, EDC |

### Columns

| Column | Description |
|---|---|
| `ID_Transaksi` | Unique transaction code |
| `Tanggal_Waktu` | Transaction datetime — inconsistent formats |
| `ID_Nasabah` | Customer code |
| `Channel` | Digital transaction channel |
| `Jenis_Transaksi` | Transfer, Pembayaran, Tarik Tunai, Top Up, Pembelian |
| `Nominal` | Transaction amount — inconsistent formats |
| `Status` | Transaction status — inconsistent labels |
| `Wilayah` | Transaction origin city |
| `Bank_Tujuan` | Destination bank (null for non-transfers) |
| `OS_Device` | Device operating system |
| `Biaya_Admin` | Administration fee |
| `Response_Time_ms` | System response time in milliseconds |
| `Jam_Tidak_Wajar` | Flag for transactions at 02.00–04.00 (1=yes, 0=no) |

---

## Data Quality Issues Found

| Issue | Description |
|---|---|
| Inconsistent status labels | `sukses`, `SUKSES`, `success`, `berhasil` refer to the same value |
| Inconsistent channel labels | `mobile banking`, `MOBILE BANKING`, `Mobilebanking` |
| Inconsistent amount formats | `50000`, `Rp 50.000`, `Rp50000`, `50000.00` |
| Missing values | ID, datetime, amount, status, fee, response time |
| Duplicate transaction IDs | ~20 duplicate rows |
| Invalid amounts | Negative values and zero-value transactions |
| Inconsistent datetime formats | 5 different date/time formats across rows |
| Response time outliers | Values exceeding 30,000 ms |

---

## Cleaning Workflow

### Sheet Structure

| Sheet | Purpose |
|---|---|
| `RAW_DATA` | Original data — never modified |
| `CLEAN_DATA` | Working sheet — all cleaning applied here |
| `CLEANING_REPORT` | Audit table documenting cleaning decisions |

---

### Step 1 — Missing Values Audit

Counted blank cells per column using `COUNTBLANK()` to identify the scale of missing data before taking action.

**Decision per column:**

| Column | Action |
|---|---|
| `ID_Transaksi` | Delete row — transaction without ID is invalid |
| `Tanggal_Waktu` | Delete row — cannot be analyzed without datetime |
| `Nominal` | Delete row — transaction without value is meaningless |
| `Status` | Label as `"Unknown"` using `IF()` |
| `Bank_Tujuan` | Retain null — expected for non-transfer transactions |
| `OS_Device` | Retain null — expected for ATM/EDC channels |
| `Biaya_Admin` | Fill with `0` |
| `Response_Time_ms` | Flag as `-1` to mark as missing |

---

### Step 2 — Remove Duplicates

Used **Data → Remove Duplicates** on column `ID_Transaksi` to eliminate double-posted transactions.

---

### Step 3 — Standardize Status Labels

Created new column `Status_Bersih` using nested `IF()` and `LOWER()` to normalize all label variations into three standard values: `Sukses`, `Gagal`, `Pending`.

```excel
=IF(OR(LOWER(G2)="sukses",LOWER(G2)="success",
LOWER(G2)="berhasil"),"Sukses",
IF(OR(LOWER(G2)="gagal",LOWER(G2)="failed",
LOWER(G2)="error"),"Gagal",
IF(OR(LOWER(G2)="pending",LOWER(G2)="menunggu"),
"Pending","Unknown")))
```

---

### Step 4 — Standardize Channel Labels

Created new column `Channel_Bersih` using the same approach to normalize channel name variations.

```excel
=IF(OR(LOWER(D2)="mobile banking",
LOWER(D2)="mobilebanking",
LOWER(D2)="mb"),"Mobile Banking",
IF(OR(LOWER(D2)="internet banking",
LOWER(D2)="ib"),"Internet Banking",
IF(LOWER(D2)="atm","ATM",
IF(LOWER(D2)="edc","EDC","Lainnya"))))
```

---

### Step 5 — Clean Amount Column

Removed currency symbols and formatting inconsistencies from `Nominal` using **Find & Replace**:

| Find | Replace |
|---|---|
| `Rp ` | *(empty)* |
| `Rp` | *(empty)* |
| `.` | *(empty)* |
| `,00` | *(empty)* |

Column was then converted to **Number** format. A validation column was added to flag invalid amounts:

```excel
=IF(OR(Nominal_Bersih<0, Nominal_Bersih=0), "Invalid", "Valid")
```

Rows flagged as `Invalid` were removed and recorded in `CLEANING_REPORT`.

---

### Step 6 — Standardize Datetime Column

The `Tanggal_Waktu` column contained 5 different formats:

```
2024-03-15 14:23:45
15/03/2024 14:23
15-03-2024 14:23:45
2024/03/15
24 Jul 2024
```

**Cleaning process:**

**i. Separate date from time component**

Used formula to extract date portion only, handling both space-separated datetime strings and text-format dates like `24 Jul 2024`:

```excel
=IF(MID(C2,3,1)=" ",
    C2,
    IF(ISNUMBER(FIND(" ",C2)),
        LEFT(C2, FIND(" ",C2)-1),
        C2))
```

**ii. Standardize separators**

Used Find & Replace to convert all `/` to `-` for consistent formatting.

**iii. Convert month names to numbers**

Used Find & Replace to replace abbreviated month names (` Jan `, ` Feb `, etc.) with zero-padded numeric equivalents (`-01-`, `-02-`, etc.).

**iv. Convert text to Excel date**

Created column `Tanggal_Bersih` using `DATE()` with logic to handle both `YYYY-MM-DD` and `DD-MM-YYYY` formats:

```excel
=IF(MID(D2,5,1)="-",
    DATE(LEFT(D2,4), MID(D2,6,2), MID(D2,9,2)),
    IF(MID(D2,3,1)="-",
        DATE(RIGHT(D2,4), MID(D2,4,2), LEFT(D2,2)),
        "Cek Manual"))
```

Column formatted as `YYYY-MM-DD`.

**v. Derived columns from `Tanggal_Bersih`:**

| Column | Formula | Example Output |
|---|---|---|
| `Bulan` | `=TEXT(Tanggal_Bersih,"MMMM YYYY")` | `Januari 2024` |
| `Kuartal` | `=IF(MONTH(...)<=3,"Q1",IF(...)...)` | `Q2` |

---

### Step 7 — Flag Response Time Outliers

Created column `Response_Flag` to classify system performance per transaction:

```excel
=IF(L2="","Missing",
IF(L2>30000,"Lambat Kritis",
IF(L2>10000,"Lambat",
IF(L2>5000,"Perlu Perhatian","Normal"))))
```

| Flag | Threshold |
|---|---|
| Normal | ≤ 5,000 ms |
| Perlu Perhatian | 5,001 – 10,000 ms |
| Lambat | 10,001 – 30,000 ms |
| Lambat Kritis | > 30,000 ms |
| Missing | No value recorded |

---

### Step 8 — Cleaning Report

All cleaning actions were documented in the `CLEANING_REPORT` sheet:

| Item | Before | After |
|---|---|---|
| Total Rows | 730 | *[result]* |
| Deleted (Missing Critical) | — | *[result]* |
| Duplicates Removed | — | *[result]* |
| Invalid Amounts Removed | — | *[result]* |
| **Total Clean Rows** | — | *[result]* |

---

## Final Clean Columns

| Column | Type | Description |
|---|---|---|
| `ID_Transaksi` | Text | Validated unique transaction ID |
| `Tanggal_Bersih` | Date | Standardized date (YYYY-MM-DD) |
| `Bulan` | Text | Month label for visualization |
| `Kuartal` | Text | Quarter label (Q1–Q4) |
| `ID_Nasabah` | Text | Customer code |
| `Channel_Bersih` | Text | Standardized channel label |
| `Jenis_Transaksi` | Text | Transaction type |
| `Nominal_Bersih` | Number | Cleaned transaction amount |
| `Status_Bersih` | Text | Standardized status label |
| `Wilayah` | Text | Transaction origin city |
| `Bank_Tujuan` | Text | Destination bank (nullable) |
| `OS_Device` | Text | Device OS (nullable) |
| `Biaya_Admin` | Number | Administration fee |
| `Response_Time_ms` | Number | System response time |
| `Response_Flag` | Text | System performance classification |
| `Jam_Tidak_Wajar` | Number | Odd-hour transaction flag |

---

## Tools

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)

---

## Project Series

| Part | Tool | Focus | Status |
|---|---|---|---|
| **Part 1** | Microsoft Excel | Data Cleaning & Preparation | ✅ Complete |
| Part 2 | Microsoft Power BI | M&E Dashboard & Visualization | 🔜 In Progress |

---

## About

**Muhammad Fauzi Romadhoni**
D4 Industrial Electrical Engineering — Politeknik Elektronika Negeri Surabaya
K3 Umum Certified (BNSP) | Internship: PT PLN UPT Gresik (Planning & Evaluation)
Data Analysis Enthusiast | Transitioning into Data Science

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/mfauziromadhoni/)
[![Email](https://img.shields.io/badge/Email-Contact-EA4335?style=flat&logo=gmail&logoColor=white)](mailto:mfauzir2477@gmail.com)
[![GitHub](https://img.shields.io/badge/GitHub-Portfolio-181717?style=flat&logo=github&logoColor=white)](https://github.com/mfauziromadhoni?tab=repositories)

---

*Part of a data analytics portfolio series targeting Monitoring & Evaluation
roles in the banking and financial services industry.*

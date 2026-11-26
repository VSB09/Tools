# SP Shah & Co. — Import Job Register Audit Tool

An auditor-grade analytical evaluation tool designed specifically for **SP Shah & Co.** to audit import Job Registers, evaluate billing turnaround time (TAT), identify Late Billing SLA breaches (> 2 Days), and export formatted 7-sheet Excel audit reports.

---

## 🚀 Standalone Browser Dashboard (`index.html`)

Open `index.html` directly in any web browser (Chrome, Edge, Firefox, Brave, Safari) — **no installation or server required**:

- **100% Offline & Private**: All file parsing, evaluation, and calculations happen in-memory inside the browser using embedded SheetJS. Zero network requests.
- **Pre-loaded Dataset**: Opens immediately with all base import jobs pre-loaded into interactive KPI cards and tables.
- **Drag-and-Drop Batch Processing**: Drag-and-drop any new `.xlsx` or legacy Crystal Reports `.xls` (BIFF8) batch file for real-time audit evaluation.
- **Filterable Audit Tables**:
  - **🚨 Late Billing Register**: Real-time filtering by delay severity (`> 2 Days`, `> 5 Days`, `> 10 Days`, `> 15 Days`, `> 30 Days`), search by client name or B/L number, with customs duty and assessable value details.
  - **📋 All Jobs Register**: Complete transaction listing with delivery date, bill date, SLA turnaround, and status classification.
  - **⏳ Delay Aging Buckets**: Distribution across standard audit aging buckets with severity levels and auditor control procedures.
  - **🏢 Client Portfolio Analysis**: Portfolio ranking by job volume, late billing rate %, and audit risk classification.
  - **⚠️ Unbilled Deliveries Risk**: Flagged completed deliveries missing billing records with revenue leakage risk alerts.
- **7-Sheet Excel Report Export**: Click **"📥 Export Excel Audit Report"** to generate an auditor-formatted workbook (`SP_Shah_Audit_Report.xlsx`).

---

## 🎯 Audit Business Rules & Evaluation

1. **Turnaround Difference**:
   $$\text{Billing Turnaround} = \text{Bill Date} - \text{Delivery Date}$$
   - **Delivery Date**: Extracted from title row column matching `Delivery Date / IGM Date / BE Date` (Column G).
   - **Bill Date**: Extracted from subordinate row under `Bill No. / Date` (Column M).

2. **Late Billing SLA Classification**:
   - **On-Time Billing** ($\le 2\text{ Days}$): Within SLA threshold.
   - **Late Billing** ($> 2\text{ Days}$): Highlighted with late billing alert and exact day count.
   - **Advance Billed** ($< 0\text{ Days}$): Billed prior to recorded delivery.
   - **Unbilled Delivery**: Delivery recorded, bill pending (revenue leakage risk).

---

## 📊 Exported Audit Workbook (`SP_Shah_Audit_Report.xlsx`)

The exported 7-sheet workbook includes:
1. **`Late Billing Summary`**: KPI dashboard and executive metrics.
2. **`Late Billing Register`**: Full listing of jobs exceeding the 2-day delivery-to-billing SLA threshold with root-cause comments.
3. **`All Jobs Master`**: Complete transaction-level job register with turnaround metrics.
4. **`Aging Analysis`**: Billing TAT aging buckets ($\le 2\text{d}$, $3\text{--}5\text{d}$, $6\text{--}10\text{d}$, $11\text{--}15\text{d}$, $16\text{--}30\text{d}$, $> 30\text{d}$).
5. **`Client Risk Analysis`**: Client-level breakdown with total jobs, late counts, late percentage, and risk ratings.
6. **`Unbilled Deliveries`**: Delivered jobs with pending billing representing revenue exposure.
7. **`Audit Methodology`**: Field detection logic, classification criteria, and statutory notes.

---

## 🛠️ Headless Python Suite (`python/`)

For command-line runs or local server hosting:
- **CLI Evaluator**: `python/run_audit.py` or `./python/run_tool.sh`
- **FastAPI Web Server**: `python/app.py` or `./python/start_web_ui.sh` (launches `http://localhost:8055`)
- **Core Processing Engine**: `python/job_register_processor.py`

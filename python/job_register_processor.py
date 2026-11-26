"""
SP Shah - Import Job Register Auditor & Analytical Automation Engine
=====================================================================
Processes port/customs import job registers (Crystal Reports XLS or formatted XLSX),
detects client entry groups, extracts Delivery Date, Bill Date, Client Address,
evaluates Late Billing (> 2 days difference), and generates an auditor-grade
multi-sheet formatted Excel workbook with revenue analysis.
"""

import os
import re
import datetime
from decimal import Decimal
from typing import List, Dict, Any, Tuple, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Try importing xlrd for legacy BIFF8 .xls files
try:
    import xlrd
except ImportError:
    xlrd = None


def parse_date(val: Any) -> Optional[datetime.date]:
    """Robust date parser handling datetime objects, Excel floats, and various string formats."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    
    # Handle xlrd float dates if passed
    if isinstance(val, (int, float)):
        # Excel date epoch: 1899-12-30
        try:
            if 30000 <= val <= 60000:
                base = datetime.date(1899, 12, 30)
                return base + datetime.timedelta(days=int(val))
        except Exception:
            pass

    s = str(val).strip()
    if not s or s.lower() in ('none', 'null', 'nan', '-', ''):
        return None

    # Handle string with time part e.g. '2026-05-14 00:00:00'
    if ' ' in s:
        s = s.split(' ')[0]

    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y', '%d.%m.%Y', '%m/%d/%Y', '%Y/%m/%d'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    # Regex search for DD/MM/YYYY or DD-MM-YYYY in mixed string
    match = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})', s)
    if match:
        d, m, y = match.groups()
        if len(y) == 2:
            y = '20' + y
        try:
            return datetime.date(int(y), int(m), int(d))
        except ValueError:
            pass

    return None


def parse_number(val: Any) -> float:
    """Safe numeric parser for currencies and quantities."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float, Decimal)):
        return float(val)
    s = str(val).replace(',', '').replace(' ', '').replace('₹', '').replace('$', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_file_rows(file_path: str) -> List[List[Any]]:
    """Loads all rows from an XLSX or XLS file as a list of lists."""
    ext = os.path.splitext(file_path)[1].lower()
    rows = []

    if ext in ('.xlsx', '.xlsm'):
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        for r in ws.iter_rows(values_only=True):
            rows.append(list(r))
        wb.close()
    elif ext == '.xls':
        if xlrd is None:
            raise RuntimeError("xlrd is required to parse .xls files. Please install xlrd.")
        wb = xlrd.open_workbook(file_path)
        sheet = wb.sheet_by_index(0)
        for r_idx in range(sheet.nrows):
            rows.append(sheet.row_values(r_idx))
    else:
        raise ValueError(f"Unsupported file format: {ext}. Expected .xlsx or .xls")

    return rows


def detect_headers_and_columns(rows: List[List[Any]]) -> Tuple[int, Dict[str, int], List[str]]:
    """
    Scans the top rows to find the main header and any subheaders,
    combining multi-row headers if present, and mapping standard audit columns.
    """
    header_idx = 0
    best_score = -1

    search_keywords = ['sr.no', 'sr no', 'jobno', 'client', 'b/l', 'delivery', 'bill no', 'duty', 'invo']

    for idx, r in enumerate(rows[:20]):
        text = ' '.join(str(c or '').lower() for c in r)
        score = sum(1 for kw in search_keywords if kw in text)
        if score > best_score:
            best_score = score
            header_idx = idx

    raw_headers = [str(c or '').strip() for c in rows[header_idx]]
    combined_headers = list(raw_headers)
    data_start_row = header_idx + 1

    # Check for subheaders within next 2 rows
    for next_idx in range(header_idx + 1, min(header_idx + 4, len(rows))):
        next_r = [str(c or '').strip() for c in rows[next_idx]]
        sub_text = ' '.join(next_r).lower()
        if any(kw in sub_text for kw in ['recd. dt', 'recd dt', 'address', 'invo. no', 'thoka no', 'net wt']):
            data_start_row = next_idx + 1
            for col_i in range(len(combined_headers)):
                if col_i < len(next_r) and next_r[col_i]:
                    if combined_headers[col_i] and next_r[col_i].lower() not in combined_headers[col_i].lower():
                        combined_headers[col_i] = f"{combined_headers[col_i]}|{next_r[col_i]}"
                    elif not combined_headers[col_i]:
                        combined_headers[col_i] = next_r[col_i]

    # Dynamic Column Mapping
    col_map = {}
    best_del_score = -1
    best_del_col = None

    for i, h in enumerate(combined_headers):
        h_clean = h.lower().replace('\n', ' ').strip()

        # Sr.No
        if any(k in h_clean for k in ['sr.no', 'sr no', 's.no']) and 'sr_no' not in col_map:
            col_map['sr_no'] = i
        # JobNo
        elif ('jobno' in h_clean or 'job no' in h_clean) and 'job_no' not in col_map:
            col_map['job_no'] = i
        # Client Name / Address
        elif ('client' in h_clean or 'customer' in h_clean or 'party' in h_clean) and 'client_name' not in col_map:
            col_map['client_name'] = i
        # B/L No
        elif ('b/l' in h_clean or 'bl no' in h_clean or 'blno' in h_clean) and 'bl_no' not in col_map:
            col_map['bl_no'] = i
        # Date column next to BL
        elif h_clean == 'date' and 'bl_date' not in col_map:
            col_map['bl_date'] = i
        # IGM Item
        elif ('igm item' in h_clean or 'igm' in h_clean) and 'item' in h_clean and 'igm_item' not in col_map:
            col_map['igm_item'] = i
        # Marks
        elif 'marks' in h_clean and 'marks_no' not in col_map:
            col_map['marks_no'] = i
        # Packages / Wt
        elif ('pkg' in h_clean or 'wt' in h_clean) and 'pkgs_wt' not in col_map:
            col_map['pkgs_wt'] = i
        # Cargo description
        elif ('cth' in h_clean or 'desc' in h_clean) and 'cargo_desc' not in col_map:
            col_map['cargo_desc'] = i
        # Invo / Ass / CIF Val
        elif ('invo' in h_clean or 'cif' in h_clean or 'ass. val' in h_clean) and 'val' in h_clean and 'invo_val' not in col_map:
            col_map['invo_val'] = i
        # Duty Paid
        elif 'duty' in h_clean and 'duty_paid' not in col_map:
            col_map['duty_paid'] = i

        # Find delivery header with most match for "Delivery", "Date", "IGM Date", "BE Date"
        del_keywords = ['delivery', 'igm date', 'be date', 'date']
        del_score = sum(1 for kw in del_keywords if kw in h_clean)
        if 'delivery' in h_clean and ('igm' in h_clean or 'be date' in h_clean):
            del_score += 5
        if del_score > best_del_score:
            best_del_score = del_score
            best_del_col = i

        # Find Bill No and Date column
        if 'bill no' in h_clean and 'date' in h_clean and 'bill_col' not in col_map:
            col_map['bill_col'] = i
        elif 'bill no' in h_clean and 'bill_col' not in col_map:
            col_map['bill_col'] = i

    if best_del_col is not None:
        col_map['delivery_col'] = best_del_col

    # Fallback column defaults if not matched
    defaults = {
        'sr_no': 0, 'job_no': 1, 'client_name': 2, 'bl_no': 3,
        'bl_date': 4, 'igm_item': 5, 'delivery_col': 6, 'marks_no': 7,
        'pkgs_wt': 8, 'cargo_desc': 9, 'invo_val': 10, 'duty_paid': 11,
        'bill_col': 12
    }
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v

    return data_start_row, col_map, combined_headers


def parse_job_register(rows: List[List[Any]]) -> Tuple[List[Dict[str, Any]], List[List[Any]], Dict[str, Any]]:
    """
    Parses the job register rows:
    - Identifies Title Rows (where Sr.No is numeric)
    - Groups sub-rows under each job
    - Assigns dedicated Group_Sr_No (1, 2, 3...)
    - Extracts Delivery Date, Bill Date, Client Address, Values
    - Evaluates Late Billing (> 2 Days) and categorizes exceptions
    """
    data_start_row, col_map, headers = detect_headers_and_columns(rows)

    c_sr = col_map['sr_no']
    c_job = col_map['job_no']
    c_client = col_map['client_name']
    c_bl = col_map['bl_no']
    c_bldate = col_map.get('bl_date', 4)
    c_igm = col_map.get('igm_item', 5)
    c_del = col_map['delivery_col']
    c_marks = col_map.get('marks_no', 7)
    c_pkg = col_map.get('pkgs_wt', 8)
    c_cargo = col_map.get('cargo_desc', 9)
    c_invo = col_map.get('invo_val', 10)
    c_duty = col_map.get('duty_paid', 11)
    c_bill = col_map['bill_col']

    jobs_data: List[Dict[str, Any]] = []
    enriched_rows: List[List[Any]] = []

    # Enriched header for raw view: Group Sr.No is Column 0
    enriched_headers = ['Group Sr. No'] + [h for h in headers]
    if len(enriched_headers) <= 14:
        enriched_headers += ['Delivery Date', 'Bill Date', 'Difference', 'Late Billing Status', 'Comments']
    enriched_rows.append(enriched_headers)

    current_job: Optional[Dict[str, Any]] = None
    group_counter = 0

    for r_idx in range(data_start_row, len(rows)):
        raw_row = rows[r_idx]
        sr_val = raw_row[c_sr] if c_sr < len(raw_row) else None

        # Check if title row (Sr.No is a valid number)
        is_title = False
        parsed_sr = None
        if sr_val is not None:
            s_str = str(sr_val).strip()
            if s_str and not any(w in s_str.lower() for w in ['sr', 'total', 'grand', 'page']):
                try:
                    parsed_sr = int(float(s_str))
                    is_title = True
                except ValueError:
                    is_title = False

        if is_title:
            group_counter += 1
            if current_job is not None:
                jobs_data.append(current_job)

            current_job = {
                'group_sr_no': group_counter,
                'sr_no': parsed_sr,
                'start_row': r_idx + 1,
                'title_row': raw_row,
                'sub_rows': [],
                'job_no': str(raw_row[c_job] or '').strip() if c_job < len(raw_row) else '',
                'client_name': str(raw_row[c_client] or '').strip() if c_client < len(raw_row) else '',
                'bl_no': str(raw_row[c_bl] or '').strip() if c_bl < len(raw_row) else '',
                'bl_date_raw': raw_row[c_bldate] if c_bldate < len(raw_row) else None,
                'igm_item': str(raw_row[c_igm] or '').strip() if c_igm < len(raw_row) else '',
                'marks_no': str(raw_row[c_marks] or '').strip() if c_marks < len(raw_row) else '',
                'pkgs_info': str(raw_row[c_pkg] or '').strip() if c_pkg < len(raw_row) else '',
                'cargo_desc': str(raw_row[c_cargo] or '').strip() if c_cargo < len(raw_row) else '',
                'invo_val_raw': raw_row[c_invo] if c_invo < len(raw_row) else None,
                'duty_paid_raw': raw_row[c_duty] if c_duty < len(raw_row) else None,
                'bill_no_raw': raw_row[c_bill] if c_bill < len(raw_row) else None,
                'delivery_date_raw': raw_row[c_del] if c_del < len(raw_row) else None,
            }
        else:
            if current_job is not None:
                current_job['sub_rows'].append(raw_row)

    if current_job is not None:
        jobs_data.append(current_job)

    # Now extract sub-row fields, calculate late billing, and assemble auditor records
    parsed_jobs = []

    for job in jobs_data:
        t_row = job['title_row']
        s_rows = job['sub_rows']

        del_date = parse_date(job['delivery_date_raw'])
        bl_date = parse_date(job['bl_date_raw'])
        duty_paid = parse_number(job['duty_paid_raw'])
        invo_val = parse_number(job['invo_val_raw'])

        bill_date = None
        client_address = ''
        recd_date = None
        invo_no_full = ''
        thoka_no = ''
        igm_date = None
        be_date = None
        ass_val = 0.0
        cif_val = 0.0
        gross_wt = None
        net_wt = None

        # Inspect sub-rows
        for s in s_rows:
            # Check for Bill Date in bill column
            if c_bill < len(s) and s[c_bill] is not None and bill_date is None:
                b_cand = parse_date(s[c_bill])
                if b_cand:
                    bill_date = b_cand

            # Check for Recd Dt in job_no column
            if c_job < len(s) and s[c_job] is not None and recd_date is None:
                r_cand = parse_date(s[c_job])
                if r_cand:
                    recd_date = r_cand

            # Check for Address in client column (usually on row with recd_date or non-empty string)
            if c_client < len(s) and s[c_client] is not None:
                s_addr = str(s[c_client]).strip()
                if s_addr and not client_address:
                    client_address = s_addr

            # Check for Invoice No in bl_no column (often contains '@' with invoice date)
            if c_bl < len(s) and s[c_bl] is not None:
                s_inv = str(s[c_bl]).strip()
                if s_inv and not invo_no_full:
                    invo_no_full = s_inv

            # Check for Thoka No in igm column
            if c_igm < len(s) and s[c_igm] is not None:
                s_thoka = str(s[c_igm]).strip()
                if s_thoka and not thoka_no:
                    thoka_no = s_thoka

            # Check for IGM Date and BE Date in delivery column
            if c_del < len(s) and s[c_del] is not None:
                d_cand = parse_date(s[c_del])
                if d_cand:
                    if igm_date is None:
                        igm_date = d_cand
                    elif be_date is None:
                        be_date = d_cand

            # Check for Assessable Value and CIF Value in invo_val column
            if c_invo < len(s) and s[c_invo] is not None:
                num = parse_number(s[c_invo])
                if num > 0:
                    if ass_val == 0.0:
                        ass_val = num
                    elif cif_val == 0.0:
                        cif_val = num

            # Check for Gross WT and Net WT in pkgs column
            if c_pkg < len(s) and s[c_pkg] is not None:
                w_num = parse_number(s[c_pkg])
                if w_num > 0:
                    if gross_wt is None:
                        gross_wt = w_num
                    elif net_wt is None:
                        net_wt = w_num

        # If assessable value was not found in sub-rows, fallback to invo_val
        if ass_val == 0.0 and invo_val > 0.0:
            ass_val = invo_val

        # Late Billing Calculation
        diff_days: Optional[int] = None
        billing_status = ""
        is_late = False
        comments = ""
        aging_bucket = ""

        if del_date is not None and bill_date is not None:
            diff_days = (bill_date - del_date).days

            # Aging bucket
            if diff_days <= 2 and diff_days >= 0:
                aging_bucket = "0-2 Days (On-Time)"
            elif 3 <= diff_days <= 5:
                aging_bucket = "3-5 Days (Low Delay)"
            elif 6 <= diff_days <= 10:
                aging_bucket = "6-10 Days (Moderate Delay)"
            elif 11 <= diff_days <= 20:
                aging_bucket = "11-20 Days (Significant Delay)"
            elif 21 <= diff_days <= 30:
                aging_bucket = "21-30 Days (High Delay)"
            elif diff_days > 30 and diff_days <= 100:
                aging_bucket = "> 30 Days (Severe Delay)"
            elif diff_days > 100:
                aging_bucket = "Date Anomaly (> 100d)"
            elif diff_days < 0:
                aging_bucket = "Advance Billed (< 0d)"

            # Status classification
            if diff_days > 100 or diff_days < -60:
                billing_status = "Date Anomaly"
                is_late = True
                comments = f"⚠️ DATA ERROR: Difference is {diff_days} days. Possible year typo (Del: {del_date.strftime('%d/%m/%Y')}, Bill: {bill_date.strftime('%d/%m/%Y')})."
            elif diff_days > 2:
                billing_status = "Late Billing"
                is_late = True
                comments = f"🚨 LATE BILLING: Delayed by {diff_days} days (Exceeds 2-day SLA limit). Mention Reason for delay."
            elif diff_days == 0:
                billing_status = "Same-Day Billing"
                is_late = False
                comments = "✅ SAME-DAY: Billed on delivery date (Compliant)."
            elif 1 <= diff_days <= 2:
                billing_status = "On-Time Billing"
                is_late = False
                comments = f"✅ ON-TIME: Billed in {diff_days} day(s) (Within 2-day limit)."
            elif diff_days < 0:
                billing_status = "Advance Billed"
                is_late = False
                comments = f"ℹ️ ADVANCE: Bill raised {abs(diff_days)} day(s) before recorded delivery date."
        elif del_date is not None and bill_date is None:
            billing_status = "Unbilled Delivery"
            is_late = True
            aging_bucket = "Unbilled Revenue Exposure"
            comments = f"🔥 CRITICAL: Cargo delivered on {del_date.strftime('%d/%m/%Y')} but NO Bill Date found. Working capital & revenue leakage risk."
        elif del_date is None and bill_date is not None:
            billing_status = "Undelivered / Billed"
            is_late = False
            aging_bucket = "Undelivered / In-Transit"
            comments = f"ℹ️ IN-TRANSIT: Bill date {bill_date.strftime('%d/%m/%Y')} present, but cargo delivery date pending."
        else:
            billing_status = "Incomplete Dates"
            is_late = False
            aging_bucket = "Incomplete Dates"
            comments = "⚠️ Missing both Delivery Date and Bill Date. Verify cargo clearance status."

        # Customs TAT (BE Date to Delivery Date)
        clearance_tat = None
        if be_date and del_date:
            clearance_tat = (del_date - be_date).days

        # Total Cycle TAT (Recd Date to Bill Date)
        total_cycle_tat = None
        if recd_date and bill_date:
            total_cycle_tat = (bill_date - recd_date).days

        job_record = {
            'group_sr_no': job['group_sr_no'],
            'sr_no': job['sr_no'],
            'job_no': job['job_no'],
            'recd_date': recd_date,
            'client_name': job['client_name'],
            'client_address': client_address,
            'bl_no': job['bl_no'],
            'bl_date': bl_date,
            'invoice_no_full': invo_no_full,
            'thoka_no': thoka_no,
            'igm_item': job['igm_item'],
            'igm_date': igm_date,
            'be_date': be_date,
            'bill_no': str(job['bill_no_raw'] or '').strip(),
            'delivery_date': del_date,
            'bill_date': bill_date,
            'difference_days': diff_days,
            'is_late': is_late,
            'billing_status': billing_status,
            'aging_bucket': aging_bucket,
            'comments': comments,
            'duty_paid': duty_paid,
            'assessable_val': ass_val,
            'cif_val': cif_val,
            'cargo_desc': job['cargo_desc'],
            'pkgs_info': job['pkgs_info'],
            'gross_wt': gross_wt,
            'net_wt': net_wt,
            'clearance_tat': clearance_tat,
            'total_cycle_tat': total_cycle_tat
        }
        parsed_jobs.append(job_record)

        # Enriched raw rows for Sheet 7 (Enriched Raw Register)
        e_title = [job['group_sr_no']] + list(t_row)
        while len(e_title) < len(enriched_headers) - 5:
            e_title.append('')
        e_title += [
            del_date.strftime('%d/%m/%Y') if del_date else '',
            bill_date.strftime('%d/%m/%Y') if bill_date else '',
            diff_days if diff_days is not None else '',
            billing_status,
            comments
        ]
        enriched_rows.append(e_title)

        for s in s_rows:
            e_sub = [job['group_sr_no']] + list(s)
            while len(e_sub) < len(enriched_headers):
                e_sub.append('')
            enriched_rows.append(e_sub)

    summary_stats = compute_summary_statistics(parsed_jobs)
    return parsed_jobs, enriched_rows, summary_stats


def compute_summary_statistics(jobs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Computes comprehensive executive auditor statistics."""
    total_jobs = len(jobs)
    on_time_jobs = [j for j in jobs if j['billing_status'] in ('On-Time Billing', 'Same-Day Billing')]
    late_jobs = [j for j in jobs if j['billing_status'] == 'Late Billing']
    unbilled_jobs = [j for j in jobs if j['billing_status'] == 'Unbilled Delivery']
    advance_jobs = [j for j in jobs if j['billing_status'] == 'Advance Billed']
    anomaly_jobs = [j for j in jobs if j['billing_status'] == 'Date Anomaly']
    undelivered_jobs = [j for j in jobs if j['billing_status'] == 'Undelivered / Billed']

    valid_diffs = [j['difference_days'] for j in jobs if j['difference_days'] is not None and -30 <= j['difference_days'] <= 100]
    avg_diff = sum(valid_diffs) / len(valid_diffs) if valid_diffs else 0.0

    total_duty = sum(j['duty_paid'] for j in jobs)
    total_ass_val = sum(j['assessable_val'] for j in jobs)
    late_duty = sum(j['duty_paid'] for j in late_jobs)
    unbilled_duty = sum(j['duty_paid'] for j in unbilled_jobs)

    buckets = {
        '0-2 Days (On-Time)': 0,
        '3-5 Days (Low Delay)': 0,
        '6-10 Days (Moderate Delay)': 0,
        '11-20 Days (Significant Delay)': 0,
        '21-30 Days (High Delay)': 0,
        '> 30 Days (Severe Delay)': 0,
        'Advance Billed (< 0d)': 0,
        'Unbilled Revenue Exposure': 0,
        'Date Anomaly / Typo': 0
    }
    for j in jobs:
        b = j.get('aging_bucket', '')
        if b in buckets:
            buckets[b] += 1
        elif 'Anomaly' in b:
            buckets['Date Anomaly / Typo'] += 1

    client_agg: Dict[str, Dict[str, Any]] = {}
    for j in jobs:
        c_name = j['client_name'] or 'UNKNOWN'
        if c_name not in client_agg:
            client_agg[c_name] = {
                'client_name': c_name,
                'total_jobs': 0,
                'on_time_jobs': 0,
                'late_jobs': 0,
                'unbilled_jobs': 0,
                'total_duty': 0.0,
                'total_ass_val': 0.0,
                'diffs': []
            }
        c_entry = client_agg[c_name]
        c_entry['total_jobs'] += 1
        if j['billing_status'] in ('On-Time Billing', 'Same-Day Billing'):
            c_entry['on_time_jobs'] += 1
        elif j['billing_status'] == 'Late Billing':
            c_entry['late_jobs'] += 1
        elif j['billing_status'] == 'Unbilled Delivery':
            c_entry['unbilled_jobs'] += 1

        c_entry['total_duty'] += j['duty_paid']
        c_entry['total_ass_val'] += j['assessable_val']
        if j['difference_days'] is not None and -30 <= j['difference_days'] <= 100:
            c_entry['diffs'].append(j['difference_days'])

    client_summary = []
    for c_name, c_data in client_agg.items():
        tot = c_data['total_jobs']
        late = c_data['late_jobs']
        late_pct = (late / tot * 100) if tot > 0 else 0.0
        avg_lag = sum(c_data['diffs']) / len(c_data['diffs']) if c_data['diffs'] else 0.0
        max_lag = max(c_data['diffs']) if c_data['diffs'] else 0

        if late_pct > 65 or c_data['unbilled_jobs'] > 5:
            risk = "HIGH RISK"
        elif late_pct > 35 or c_data['unbilled_jobs'] > 0:
            risk = "MEDIUM RISK"
        else:
            risk = "COMPLIANT"

        client_summary.append({
            'client_name': c_name,
            'total_jobs': tot,
            'on_time_jobs': c_data['on_time_jobs'],
            'late_jobs': late,
            'late_pct': round(late_pct, 1),
            'unbilled_jobs': c_data['unbilled_jobs'],
            'avg_billing_lag': round(avg_lag, 1),
            'max_billing_lag': max_lag,
            'total_duty': c_data['total_duty'],
            'total_ass_val': c_data['total_ass_val'],
            'risk_tier': risk
        })

    client_summary.sort(key=lambda x: x['total_jobs'], reverse=True)

    return {
        'total_jobs': total_jobs,
        'on_time_count': len(on_time_jobs),
        'on_time_pct': round((len(on_time_jobs) / total_jobs * 100) if total_jobs else 0, 2),
        'late_billing_count': len(late_jobs),
        'late_billing_pct': round((len(late_jobs) / total_jobs * 100) if total_jobs else 0, 2),
        'unbilled_count': len(unbilled_jobs),
        'unbilled_pct': round((len(unbilled_jobs) / total_jobs * 100) if total_jobs else 0, 2),
        'advance_count': len(advance_jobs),
        'anomaly_count': len(anomaly_jobs),
        'undelivered_count': len(undelivered_jobs),
        'avg_billing_lag': round(avg_diff, 2),
        'total_duty_paid': total_duty,
        'total_assessable_value': total_ass_val,
        'late_duty_exposure': late_duty,
        'unbilled_duty_exposure': unbilled_duty,
        'buckets': buckets,
        'client_summary': client_summary
    }


def generate_auditor_excel_workbook(parsed_jobs: List[Dict[str, Any]],
                                    enriched_rows: List[List[Any]],
                                    stats: Dict[str, Any],
                                    output_filepath: str):
    """
    Builds a presentation-grade, professional multi-sheet Excel workbook.
    - Sheet 1: Late_Billing_Register (Fixed First Sheet focused on Late Billing)
    - Sheet 2: Executive_Audit_Summary (KPI scorecard & delay buckets)
    - Sheet 3: Client_Revenue_TAT_Analysis (Client-wise aggregation)
    - Sheet 4: Billing_Aging_Drilldown (Detailed bucket breakdown)
    - Sheet 5: Unbilled_Revenue_Exposure (Deliveries without billing)
    - Sheet 6: Data_Integrity_Anomalies (Data quality exception log)
    - Sheet 7: Enriched_Raw_Register (Multi-row format with Group Sr.No)
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Fonts
    FONT_TITLE = Font(name='Segoe UI', size=15, bold=True, color='1B365D')
    FONT_SUBTITLE = Font(name='Segoe UI', size=10, italic=True, color='555555')
    FONT_HEADER = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    FONT_DATA = Font(name='Segoe UI', size=9)
    FONT_BOLD = Font(name='Segoe UI', size=9, bold=True)
    FONT_KPI_VAL = Font(name='Segoe UI', size=16, bold=True, color='1B365D')
    FONT_KPI_LBL = Font(name='Segoe UI', size=9, bold=True, color='555555')

    # Fills
    FILL_NAVY = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
    FILL_STEEL = PatternFill(start_color='2C4D75', end_color='2C4D75', fill_type='solid')
    FILL_ZEBRA = PatternFill(start_color='F9FBFC', end_color='F9FBFC', fill_type='solid')
    FILL_KPI = PatternFill(start_color='EDF3F9', end_color='EDF3F9', fill_type='solid')

    # Badges
    FILL_LATE = PatternFill(start_color='FFD2D2', end_color='FFD2D2', fill_type='solid')
    FONT_LATE = Font(name='Segoe UI', size=9, bold=True, color='900C3F')
    FILL_ONTIME = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')
    FONT_ONTIME = Font(name='Segoe UI', size=9, bold=True, color='155724')
    FILL_UNBILLED = PatternFill(start_color='FFEBA8', end_color='FFEBA8', fill_type='solid')
    FONT_UNBILLED = Font(name='Segoe UI', size=9, bold=True, color='856404')
    FILL_ADVANCE = PatternFill(start_color='D1ECF1', end_color='D1ECF1', fill_type='solid')
    FONT_ADVANCE = Font(name='Segoe UI', size=9, bold=True, color='0C5460')
    FILL_ANOMALY = PatternFill(start_color='E2E3E5', end_color='E2E3E5', fill_type='solid')
    FONT_ANOMALY = Font(name='Segoe UI', size=9, bold=True, color='383D41')

    BORDER_THIN = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    BORDER_HEADER = Border(
        left=Side(style='thin', color='FFFFFF'),
        right=Side(style='thin', color='FFFFFF'),
        top=Side(style='medium', color='1B365D'),
        bottom=Side(style='medium', color='1B365D')
    )

    ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
    ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
    ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
    ALIGN_HEADER = Alignment(horizontal='center', vertical='center', wrap_text=True)

    FMT_CURRENCY = '₹#,##0.00'
    FMT_INTEGER = '#,##0'
    FMT_PERCENT = '0.0%'
    FMT_DATE = 'DD/MM/YYYY'

    # =========================================================================
    # SHEET 1: Late_Billing_Register (Fixed First Sheet focused on Late Billing)
    # =========================================================================
    ws1 = wb.create_sheet(title='Late_Billing_Register')
    ws1.views.sheetView[0].showGridLines = True

    ws1['A1'] = "SP SHAH — IMPORT JOB REGISTER AUDIT: LATE BILLING EVALUATION"
    ws1['A1'].font = FONT_TITLE
    ws1['A2'] = f"Fixed Evaluation Output | Criteria: Late Billing = Difference (Bill Date - Delivery Date) > 2 Days | Generated: {datetime.datetime.now().strftime('%d-%b-%Y %H:%M')}"
    ws1['A2'].font = FONT_SUBTITLE

    s1_cols = [
        ('Group Sr. No', 12, ALIGN_CENTER, FMT_INTEGER),
        ('Sr.No', 9, ALIGN_CENTER, FMT_INTEGER),
        ('Job No', 10, ALIGN_CENTER, '@'),
        ('Client Name', 28, ALIGN_LEFT, '@'),
        ('Bill Date', 13, ALIGN_CENTER, FMT_DATE),
        ('Delivery Date', 14, ALIGN_CENTER, FMT_DATE),
        ('Difference (Days)', 16, ALIGN_CENTER, FMT_INTEGER),
        ('Billing Status', 18, ALIGN_CENTER, '@'),
        ('Auditor Comments & Remarks', 45, ALIGN_LEFT, '@'),
        ('Duty Paid (₹)', 16, ALIGN_RIGHT, FMT_CURRENCY),
        ('Assessable Value (₹)', 20, ALIGN_RIGHT, FMT_CURRENCY),
        ('Bill No', 18, ALIGN_CENTER, '@'),
        ('B/L No', 18, ALIGN_CENTER, '@'),
        ('Invoice No & Date', 24, ALIGN_LEFT, '@'),
        ('Client Address', 35, ALIGN_LEFT, '@'),
        ('Cargo Description (CTH)', 30, ALIGN_LEFT, '@'),
        ('No.Of Pkgs / Containers', 22, ALIGN_LEFT, '@'),
        ('Job Recd Date', 13, ALIGN_CENTER, FMT_DATE),
        ('BE Date', 13, ALIGN_CENTER, FMT_DATE),
        ('Clearance TAT (Days)', 18, ALIGN_CENTER, FMT_INTEGER),
        ('Total Cycle TAT (Days)', 18, ALIGN_CENTER, FMT_INTEGER),
    ]

    for c_idx, (col_name, col_w, _, _) in enumerate(s1_cols, 1):
        cell = ws1.cell(row=4, column=c_idx, value=col_name)
        cell.font = FONT_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_HEADER
        cell.border = BORDER_HEADER
        ws1.column_dimensions[get_column_letter(c_idx)].width = col_w
    ws1.row_dimensions[4].height = 28

    for r_i, job in enumerate(parsed_jobs, start=5):
        row_fill = FILL_ZEBRA if (r_i % 2 == 0) else None
        diff = job['difference_days']
        stat = job['billing_status']

        row_vals = [
            job['group_sr_no'],
            job['sr_no'],
            job['job_no'],
            job['client_name'],
            job['bill_date'].strftime('%d/%m/%Y') if job['bill_date'] else '',
            job['delivery_date'].strftime('%d/%m/%Y') if job['delivery_date'] else '',
            diff if diff is not None else '',
            stat,
            job['comments'],
            job['duty_paid'],
            job['assessable_val'],
            job['bill_no'],
            job['bl_no'],
            job['invoice_no_full'],
            job['client_address'],
            job['cargo_desc'],
            job['pkgs_info'],
            job['recd_date'].strftime('%d/%m/%Y') if job['recd_date'] else '',
            job['be_date'].strftime('%d/%m/%Y') if job['be_date'] else '',
            job['clearance_tat'] if job['clearance_tat'] is not None else '',
            job['total_cycle_tat'] if job['total_cycle_tat'] is not None else ''
        ]

        for c_idx, (val, (_, _, align_val, num_fmt)) in enumerate(zip(row_vals, s1_cols), 1):
            cell = ws1.cell(row=r_i, column=c_idx, value=val)
            cell.font = FONT_DATA
            cell.alignment = align_val
            cell.border = BORDER_THIN
            if row_fill:
                cell.fill = row_fill
            if num_fmt and num_fmt != '@' and isinstance(val, (int, float)):
                cell.number_format = num_fmt

        status_cell = ws1.cell(row=r_i, column=8)
        diff_cell = ws1.cell(row=r_i, column=7)

        if stat == 'Late Billing':
            status_cell.fill = FILL_LATE
            status_cell.font = FONT_LATE
            diff_cell.fill = FILL_LATE
            diff_cell.font = FONT_LATE
        elif stat in ('On-Time Billing', 'Same-Day Billing'):
            status_cell.fill = FILL_ONTIME
            status_cell.font = FONT_ONTIME
        elif stat == 'Unbilled Delivery':
            status_cell.fill = FILL_UNBILLED
            status_cell.font = FONT_UNBILLED
        elif stat == 'Advance Billed':
            status_cell.fill = FILL_ADVANCE
            status_cell.font = FONT_ADVANCE
        elif stat == 'Date Anomaly':
            status_cell.fill = FILL_ANOMALY
            status_cell.font = FONT_ANOMALY

    ws1.freeze_panes = 'E5'
    ws1.auto_filter.ref = f"A4:{get_column_letter(len(s1_cols))}{len(parsed_jobs) + 4}"

    # =========================================================================
    # SHEET 2: Executive_Audit_Summary (KPI Dashboard)
    # =========================================================================
    ws2 = wb.create_sheet(title='Executive_Audit_Summary')
    ws2.views.sheetView[0].showGridLines = True

    ws2['A1'] = "SP SHAH — EXECUTIVE AUDIT SCORECARD & REVENUE TAT ANALYSIS"
    ws2['A1'].font = FONT_TITLE
    ws2['A2'] = f"Analytical procedures & internal control evaluation (Standards on Auditing SA 520, SA 265, SA 240) | Base Jobs: {stats['total_jobs']:,}"
    ws2['A2'].font = FONT_SUBTITLE

    kpis = [
        ('TOTAL IMPORT JOBS', f"{stats['total_jobs']:,}", 'A', 'B'),
        ('LATE BILLING COUNT (>2d)', f"{stats['late_billing_count']:,} ({stats['late_billing_pct']}%)", 'D', 'E'),
        ('ON-TIME BILLING (0-2d)', f"{stats['on_time_count']:,} ({stats['on_time_pct']}%)", 'G', 'H'),
        ('UNBILLED DELIVERY RISK', f"{stats['unbilled_count']:,} ({stats['unbilled_pct']}%)", 'J', 'K'),
    ]
    for lbl, val, col_start, col_end in kpis:
        c1 = f"{col_start}4"
        c2 = f"{col_end}4"
        ws2.merge_cells(f"{c1}:{c2}")
        ws2[c1] = lbl
        ws2[c1].font = FONT_KPI_LBL
        ws2[c1].alignment = ALIGN_CENTER
        ws2[c1].fill = FILL_KPI

        v1 = f"{col_start}5"
        v2 = f"{col_end}5"
        ws2.merge_cells(f"{v1}:{v2}")
        ws2[v1] = val
        ws2[v1].font = FONT_KPI_VAL
        ws2[v1].alignment = ALIGN_CENTER
        ws2[v1].fill = FILL_KPI

    fin_kpis = [
        ('TOTAL DUTY CLEARED', f"₹{stats['total_duty_paid']:,.2f}", 'A', 'B'),
        ('TOTAL ASSESSABLE VALUE', f"₹{stats['total_assessable_value']:,.2f}", 'D', 'E'),
        ('LATE BILLING DUTY EXPOSURE', f"₹{stats['late_duty_exposure']:,.2f}", 'G', 'H'),
        ('UNBILLED REVENUE EXPOSURE', f"₹{stats['unbilled_duty_exposure']:,.2f}", 'J', 'K'),
    ]
    for lbl, val, col_start, col_end in fin_kpis:
        c1 = f"{col_start}7"
        c2 = f"{col_end}7"
        ws2.merge_cells(f"{c1}:{c2}")
        ws2[c1] = lbl
        ws2[c1].font = FONT_KPI_LBL
        ws2[c1].alignment = ALIGN_CENTER
        ws2[c1].fill = FILL_KPI

        v1 = f"{col_start}8"
        v2 = f"{col_end}8"
        ws2.merge_cells(f"{v1}:{v2}")
        ws2[v1] = val
        ws2[v1].font = FONT_KPI_VAL
        ws2[v1].alignment = ALIGN_CENTER
        ws2[v1].fill = FILL_KPI

    ws2['A11'] = "BILLING DELAY AGING BRACKETS & SEVERITY BREAKDOWN"
    ws2['A11'].font = Font(name='Segoe UI', size=11, bold=True, color='1B365D')

    b_headers = ['Delay Aging Bracket', 'Job Count', '% of Total Jobs', 'Audit Severity Level', 'Recommended Internal Control Action']
    for c_i, h in enumerate(b_headers, 1):
        cell = ws2.cell(row=12, column=c_i, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_HEADER

    bucket_severity = [
        ('0-2 Days (On-Time)', 'LOW / COMPLIANT', 'Normal operational clearance. Meets company SLA policy.'),
        ('3-5 Days (Low Delay)', 'MODERATE', 'Monitor billing desk queue. Ensure bills raised within 48 hours.'),
        ('6-10 Days (Moderate Delay)', 'ELEVATED', 'Review documentation handoff between CFS delivery and finance.'),
        ('11-20 Days (Significant Delay)', 'HIGH', 'Mandatory explanation from operations manager before billing.'),
        ('21-30 Days (High Delay)', 'CRITICAL', 'Direct audit exception. Revenue cycle delay affecting cash flow.'),
        ('> 30 Days (Severe Delay)', 'SEVERE SLA BREACH', 'Escalate to TCWG/Directors. Investigate client disputed billing.'),
        ('Advance Billed (< 0d)', 'PROCESS ANOMALY', 'Verify cargo dispatch confirmation prior to invoice release.'),
        ('Unbilled Revenue Exposure', 'UNBILLED RISK', 'Immediate unbilled revenue reconciliation and invoice dispatch.'),
        ('Date Anomaly / Typo', 'DATA QUALITY', 'Audit log fix: correct manual typing of delivery/billing dates.')
    ]

    for b_idx, (b_name, sev, action) in enumerate(bucket_severity, start=13):
        cnt = stats['buckets'].get(b_name, 0)
        pct = (cnt / stats['total_jobs']) if stats['total_jobs'] else 0.0

        ws2.cell(row=b_idx, column=1, value=b_name).alignment = ALIGN_LEFT
        ws2.cell(row=b_idx, column=2, value=cnt).alignment = ALIGN_CENTER
        ws2.cell(row=b_idx, column=2).number_format = FMT_INTEGER
        ws2.cell(row=b_idx, column=3, value=pct).alignment = ALIGN_CENTER
        ws2.cell(row=b_idx, column=3).number_format = FMT_PERCENT
        ws2.cell(row=b_idx, column=4, value=sev).alignment = ALIGN_CENTER
        ws2.cell(row=b_idx, column=5, value=action).alignment = ALIGN_LEFT

        for c_k in range(1, 6):
            ws2.cell(row=b_idx, column=c_k).border = BORDER_THIN
            ws2.cell(row=b_idx, column=c_k).font = FONT_DATA

    ws2.column_dimensions['A'].width = 30
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 18
    ws2.column_dimensions['D'].width = 25
    ws2.column_dimensions['E'].width = 50
    ws2.column_dimensions['G'].width = 28
    ws2.column_dimensions['H'].width = 16
    ws2.column_dimensions['J'].width = 28
    ws2.column_dimensions['K'].width = 16

    # =========================================================================
    # SHEET 3: Client_Revenue_TAT_Analysis
    # =========================================================================
    ws3 = wb.create_sheet(title='Client_Revenue_TAT_Analysis')
    ws3.views.sheetView[0].showGridLines = True

    ws3['A1'] = "SP SHAH — CLIENT-WISE BILLING TIMELINESS & REVENUE PORTFOLIO"
    ws3['A1'].font = FONT_TITLE
    ws3['A2'] = f"Aggregation of {len(stats['client_summary'])} clients across all port job transactions | Auditor Risk Profiling"
    ws3['A2'].font = FONT_SUBTITLE

    s3_cols = [
        ('Client Name', 32, ALIGN_LEFT, '@'),
        ('Total Jobs Cleared', 18, ALIGN_CENTER, FMT_INTEGER),
        ('On-Time Jobs (<=2d)', 18, ALIGN_CENTER, FMT_INTEGER),
        ('Late Billed Jobs (>2d)', 18, ALIGN_CENTER, FMT_INTEGER),
        ('Late Billing %', 16, ALIGN_CENTER, FMT_PERCENT),
        ('Unbilled Jobs', 16, ALIGN_CENTER, FMT_INTEGER),
        ('Avg Billing Delay (Days)', 22, ALIGN_CENTER, '0.0'),
        ('Max Billing Delay (Days)', 22, ALIGN_CENTER, FMT_INTEGER),
        ('Total Duty Paid (₹)', 22, ALIGN_RIGHT, FMT_CURRENCY),
        ('Total Assessable Value (₹)', 24, ALIGN_RIGHT, FMT_CURRENCY),
        ('Auditor Risk Classification', 25, ALIGN_CENTER, '@')
    ]

    for c_i, (h, w, _, _) in enumerate(s3_cols, 1):
        cell = ws3.cell(row=4, column=c_i, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_HEADER
        cell.border = BORDER_HEADER
        ws3.column_dimensions[get_column_letter(c_i)].width = w
    ws3.row_dimensions[4].height = 28

    for r_i, c_row in enumerate(stats['client_summary'], start=5):
        row_fill = FILL_ZEBRA if (r_i % 2 == 0) else None
        vals = [
            c_row['client_name'],
            c_row['total_jobs'],
            c_row['on_time_jobs'],
            c_row['late_jobs'],
            c_row['late_pct'] / 100.0,
            c_row['unbilled_jobs'],
            c_row['avg_billing_lag'],
            c_row['max_billing_lag'],
            c_row['total_duty'],
            c_row['total_ass_val'],
            c_row['risk_tier']
        ]

        for c_idx, (v, (_, _, align_val, num_fmt)) in enumerate(zip(vals, s3_cols), 1):
            cell = ws3.cell(row=r_i, column=c_idx, value=v)
            cell.font = FONT_DATA
            cell.alignment = align_val
            cell.border = BORDER_THIN
            if row_fill:
                cell.fill = row_fill
            if num_fmt and num_fmt != '@':
                cell.number_format = num_fmt

        risk_cell = ws3.cell(row=r_i, column=11)
        if c_row['risk_tier'] == 'HIGH RISK':
            risk_cell.fill = FILL_LATE
            risk_cell.font = FONT_LATE
        elif c_row['risk_tier'] == 'COMPLIANT':
            risk_cell.fill = FILL_ONTIME
            risk_cell.font = FONT_ONTIME

    ws3.freeze_panes = 'B5'
    ws3.auto_filter.ref = f"A4:{get_column_letter(len(s3_cols))}{len(stats['client_summary']) + 4}"

    # =========================================================================
    # SHEET 4: Billing_Aging_Drilldown
    # =========================================================================
    ws4 = wb.create_sheet(title='Billing_Aging_Drilldown')
    ws4.views.sheetView[0].showGridLines = True

    ws4['A1'] = "SP SHAH — BILLING DELAY AGING DRILLDOWN: CRITICAL & SEVERE DELAYS"
    ws4['A1'].font = FONT_TITLE
    ws4['A2'] = "Drilldown of transactions delayed > 15 days or severe delay buckets requiring management justification"
    ws4['A2'].font = FONT_SUBTITLE

    severe_jobs = [j for j in parsed_jobs if j['difference_days'] is not None and j['difference_days'] > 15 and j['difference_days'] <= 100]
    severe_jobs.sort(key=lambda x: x['difference_days'], reverse=True)

    s4_cols = [
        ('Group Sr. No', 12, ALIGN_CENTER, FMT_INTEGER),
        ('Job No', 10, ALIGN_CENTER, '@'),
        ('Client Name', 28, ALIGN_LEFT, '@'),
        ('Bill Date', 13, ALIGN_CENTER, FMT_DATE),
        ('Delivery Date', 14, ALIGN_CENTER, FMT_DATE),
        ('Delay (Days)', 14, ALIGN_CENTER, FMT_INTEGER),
        ('Aging Bucket', 22, ALIGN_CENTER, '@'),
        ('Duty Paid (₹)', 18, ALIGN_RIGHT, FMT_CURRENCY),
        ('Assessable Value (₹)', 20, ALIGN_RIGHT, FMT_CURRENCY),
        ('Bill No', 18, ALIGN_CENTER, '@'),
        ('B/L No', 18, ALIGN_CENTER, '@'),
        ('Auditor Remarks / Delay Reason Required', 45, ALIGN_LEFT, '@')
    ]

    for c_i, (h, w, _, _) in enumerate(s4_cols, 1):
        cell = ws4.cell(row=4, column=c_i, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_HEADER
        cell.border = BORDER_HEADER
        ws4.column_dimensions[get_column_letter(c_i)].width = w
    ws4.row_dimensions[4].height = 28

    for r_i, job in enumerate(severe_jobs, start=5):
        row_fill = FILL_ZEBRA if (r_i % 2 == 0) else None
        vals = [
            job['group_sr_no'],
            job['job_no'],
            job['client_name'],
            job['bill_date'].strftime('%d/%m/%Y') if job['bill_date'] else '',
            job['delivery_date'].strftime('%d/%m/%Y') if job['delivery_date'] else '',
            job['difference_days'],
            job['aging_bucket'],
            job['duty_paid'],
            job['assessable_val'],
            job['bill_no'],
            job['bl_no'],
            job['comments']
        ]

        for c_idx, (v, (_, _, align_val, num_fmt)) in enumerate(zip(vals, s4_cols), 1):
            cell = ws4.cell(row=r_i, column=c_idx, value=v)
            cell.font = FONT_DATA
            cell.alignment = align_val
            cell.border = BORDER_THIN
            if row_fill:
                cell.fill = row_fill
            if num_fmt and num_fmt != '@':
                cell.number_format = num_fmt

        diff_c = ws4.cell(row=r_i, column=6)
        diff_c.fill = FILL_LATE
        diff_c.font = FONT_LATE

    ws4.freeze_panes = 'D5'
    ws4.auto_filter.ref = f"A4:{get_column_letter(len(s4_cols))}{len(severe_jobs) + 4}"

    # =========================================================================
    # SHEET 5: Unbilled_Revenue_Exposure
    # =========================================================================
    ws5 = wb.create_sheet(title='Unbilled_Revenue_Exposure')
    ws5.views.sheetView[0].showGridLines = True

    ws5['A1'] = "SP SHAH — UNBILLED REVENUE EXPOSURE: DELIVERIES WITHOUT BILL DATE"
    ws5['A1'].font = FONT_TITLE
    ws5['A2'] = "Internal Control Defect: Physical cargo delivered to client but invoice not yet booked in register | Cash leakage risk"
    ws5['A2'].font = FONT_SUBTITLE

    unbilled_list = [j for j in parsed_jobs if j['billing_status'] == 'Unbilled Delivery']
    unbilled_list.sort(key=lambda x: x['delivery_date'] or datetime.date(2000, 1, 1), reverse=True)

    s5_cols = [
        ('Group Sr. No', 12, ALIGN_CENTER, FMT_INTEGER),
        ('Job No', 10, ALIGN_CENTER, '@'),
        ('Client Name', 28, ALIGN_LEFT, '@'),
        ('Delivery Date', 14, ALIGN_CENTER, FMT_DATE),
        ('Days Since Delivery', 18, ALIGN_CENTER, FMT_INTEGER),
        ('Duty Paid (₹)', 18, ALIGN_RIGHT, FMT_CURRENCY),
        ('Assessable Value (₹)', 20, ALIGN_RIGHT, FMT_CURRENCY),
        ('B/L No', 18, ALIGN_CENTER, '@'),
        ('Client Address', 35, ALIGN_LEFT, '@'),
        ('Cargo Description', 28, ALIGN_LEFT, '@'),
        ('Auditor Risk Notice', 45, ALIGN_LEFT, '@')
    ]

    for c_i, (h, w, _, _) in enumerate(s5_cols, 1):
        cell = ws5.cell(row=4, column=c_i, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_HEADER
        cell.border = BORDER_HEADER
        ws5.column_dimensions[get_column_letter(c_i)].width = w
    ws5.row_dimensions[4].height = 28

    today = datetime.date.today()
    for r_i, job in enumerate(unbilled_list, start=5):
        row_fill = FILL_ZEBRA if (r_i % 2 == 0) else None
        days_since = (today - job['delivery_date']).days if job['delivery_date'] else ''
        vals = [
            job['group_sr_no'],
            job['job_no'],
            job['client_name'],
            job['delivery_date'].strftime('%d/%m/%Y') if job['delivery_date'] else '',
            days_since,
            job['duty_paid'],
            job['assessable_val'],
            job['bl_no'],
            job['client_address'],
            job['cargo_desc'],
            job['comments']
        ]

        for c_idx, (v, (_, _, align_val, num_fmt)) in enumerate(zip(vals, s5_cols), 1):
            cell = ws5.cell(row=r_i, column=c_idx, value=v)
            cell.font = FONT_DATA
            cell.alignment = align_val
            cell.border = BORDER_THIN
            if row_fill:
                cell.fill = row_fill
            if num_fmt and num_fmt != '@':
                cell.number_format = num_fmt

        stat_c = ws5.cell(row=r_i, column=11)
        stat_c.fill = FILL_UNBILLED
        stat_c.font = FONT_UNBILLED

    ws5.freeze_panes = 'D5'
    ws5.auto_filter.ref = f"A4:{get_column_letter(len(s5_cols))}{len(unbilled_list) + 4}"

    # =========================================================================
    # SHEET 6: Data_Integrity_Anomalies
    # =========================================================================
    ws6 = wb.create_sheet(title='Data_Integrity_Anomalies')
    ws6.views.sheetView[0].showGridLines = True

    ws6['A1'] = "SP SHAH — DATA QUALITY & INTEGRITY EXCEPTIONS"
    ws6['A1'].font = FONT_TITLE
    ws6['A2'] = "Transactions with extreme date differences (> 100 days), negative diffs (< -30 days), or year format typos"
    ws6['A2'].font = FONT_SUBTITLE

    anomaly_list = [j for j in parsed_jobs if j['billing_status'] == 'Date Anomaly' or (j['difference_days'] is not None and j['difference_days'] < -30)]

    s6_cols = [
        ('Group Sr. No', 12, ALIGN_CENTER, FMT_INTEGER),
        ('Job No', 10, ALIGN_CENTER, '@'),
        ('Client Name', 28, ALIGN_LEFT, '@'),
        ('Bill Date', 14, ALIGN_CENTER, FMT_DATE),
        ('Delivery Date', 14, ALIGN_CENTER, FMT_DATE),
        ('Calculated Diff (Days)', 20, ALIGN_CENTER, FMT_INTEGER),
        ('B/L No', 18, ALIGN_CENTER, '@'),
        ('Bill No', 18, ALIGN_CENTER, '@'),
        ('Exception Description / Audit Root Cause', 50, ALIGN_LEFT, '@')
    ]

    for c_i, (h, w, _, _) in enumerate(s6_cols, 1):
        cell = ws6.cell(row=4, column=c_i, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_HEADER
        cell.border = BORDER_HEADER
        ws6.column_dimensions[get_column_letter(c_i)].width = w
    ws6.row_dimensions[4].height = 28

    for r_i, job in enumerate(anomaly_list, start=5):
        vals = [
            job['group_sr_no'],
            job['job_no'],
            job['client_name'],
            job['bill_date'].strftime('%d/%m/%Y') if job['bill_date'] else '',
            job['delivery_date'].strftime('%d/%m/%Y') if job['delivery_date'] else '',
            job['difference_days'],
            job['bl_no'],
            job['bill_no'],
            job['comments']
        ]

        for c_idx, (v, (_, _, align_val, num_fmt)) in enumerate(zip(vals, s6_cols), 1):
            cell = ws6.cell(row=r_i, column=c_idx, value=v)
            cell.font = FONT_DATA
            cell.alignment = align_val
            cell.border = BORDER_THIN
            if num_fmt and num_fmt != '@':
                cell.number_format = num_fmt

        diff_c = ws6.cell(row=r_i, column=6)
        diff_c.fill = FILL_ANOMALY
        diff_c.font = FONT_ANOMALY

    ws6.freeze_panes = 'D5'
    ws6.auto_filter.ref = f"A4:{get_column_letter(len(s6_cols))}{len(anomaly_list) + 4}"

    # =========================================================================
    # SHEET 7: Enriched_Raw_Register
    # =========================================================================
    ws7 = wb.create_sheet(title='Enriched_Raw_Register')
    ws7.views.sheetView[0].showGridLines = True

    r_headers = enriched_rows[0]
    ws7.append(r_headers)
    for c_i in range(1, len(r_headers) + 1):
        cell = ws7.cell(row=1, column=c_i)
        cell.font = FONT_HEADER
        cell.fill = FILL_STEEL
        cell.alignment = ALIGN_HEADER
        ws7.column_dimensions[get_column_letter(c_i)].width = 16

    ws7.column_dimensions['A'].width = 14
    ws7.column_dimensions['D'].width = 30

    for row_data in enriched_rows[1:]:
        ws7.append(row_data)

    ws7.freeze_panes = 'B2'
    ws7.auto_filter.ref = f"A1:{get_column_letter(len(r_headers))}{len(enriched_rows)}"

    os.makedirs(os.path.dirname(os.path.abspath(output_filepath)), exist_ok=True)
    wb.save(output_filepath)
    wb.close()
    return output_filepath

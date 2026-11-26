#!/usr/bin/env python3
"""
SP Shah - Import Job Register Audit CLI Tool
=============================================
Evaluates Job Register batches for Late Billing (> 2 days difference between
Delivery Date and Bill Date) and generates a multi-sheet auditor workbook.

Usage:
    python3 run_audit.py [path_to_input_file.xlsx or .xls]
"""

import os
import sys
import time
import argparse
from datetime import datetime

# Add current directory to path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from job_register_processor import (
    load_file_rows,
    parse_job_register,
    generate_auditor_excel_workbook
)


def run_audit(input_file: str, output_file: str = None) -> str:
    print("=" * 80)
    print("  SP SHAH — IMPORT JOB REGISTER AUDIT & ANALYTICAL EVALUATOR")
    print("  Automated Analytical Evaluator: Late Billing & Revenue TAT")
    print("=" * 80)

    if not os.path.exists(input_file):
        print(f"\n[ERROR] File not found: {input_file}")
        sys.exit(1)

    print(f"\n[1/4] Loading input file: {os.path.basename(input_file)}")
    t0 = time.time()
    rows = load_file_rows(input_file)
    load_time = time.time() - t0
    print(f"      Loaded {len(rows):,} raw rows in {load_time:.2f} seconds.")

    print(f"\n[2/4] Parsing client groups, extracting dates, evaluating late billing...")
    t1 = time.time()
    parsed_jobs, enriched_rows, stats = parse_job_register(rows)
    parse_time = time.time() - t1
    print(f"      Identified {len(parsed_jobs):,} transaction jobs in {parse_time:.2f} seconds.")

    if not output_file:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(current_dir, "Outputs")
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f"SP_Shah_Audit_Report_{timestamp}.xlsx")

    print(f"\n[3/4] Generating 7-sheet professional auditor Excel workbook...")
    t2 = time.time()
    generate_auditor_excel_workbook(parsed_jobs, enriched_rows, stats, output_file)
    gen_time = time.time() - t2
    print(f"      Saved workbook in {gen_time:.2f} seconds to:")
    print(f"      -> {output_file}")

    print("\n" + "=" * 80)
    print("  EXECUTIVE AUDIT SUMMARY — KEY FINDINGS")
    print("=" * 80)
    print(f"  • Total Import Jobs Cleared       : {stats['total_jobs']:,}")
    print(f"  • Late Billing Jobs (> 2 Days)    : {stats['late_billing_count']:,} ({stats['late_billing_pct']}%)  [SLA BREACH]")
    print(f"  • On-Time Billed Jobs (<= 2 Days) : {stats['on_time_count']:,} ({stats['on_time_pct']}%)  [COMPLIANT]")
    print(f"  • Unbilled Cargo Deliveries       : {stats['unbilled_count']:,} ({stats['unbilled_pct']}%)  [CRITICAL RISK]")
    print(f"  • Advance Billed (Pre-delivery)   : {stats['advance_count']:,}")
    print(f"  • Average Billing Turnaround Lag  : {stats['avg_billing_lag']} days")
    print(f"  • Total Customs Duty Processed    : ₹{stats['total_duty_paid']:,.2f}")
    print(f"  • Duty Exposed to Late Billing    : ₹{stats['late_duty_exposure']:,.2f}")
    print(f"  • Duty on Unbilled Cargo Deliveries: ₹{stats['unbilled_duty_exposure']:,.2f}")

    print("\n  DELAY AGING DISTRIBUTION:")
    print("  " + "-" * 76)
    print(f"  {'Delay Bracket':<32} {'Count':<10} {'% of Total':<12} {'Severity'}")
    print("  " + "-" * 76)
    for b_name, count in stats['buckets'].items():
        pct = (count / stats['total_jobs'] * 100) if stats['total_jobs'] else 0
        sev = "COMPLIANT" if "0-2" in b_name else ("ELEVATED" if "3-5" in b_name or "6-10" in b_name else "HIGH / CRITICAL")
        print(f"  {b_name:<32} {count:<10,} {pct:>6.2f}%      {sev}")
    print("  " + "-" * 76)

    print("\n  TOP CLIENTS BY LATE BILLING COUNT:")
    print("  " + "-" * 76)
    print(f"  {'Client Name':<32} {'Total Jobs':<12} {'Late Jobs':<12} {'Late %':<10} {'Avg Delay'}")
    print("  " + "-" * 76)
    for c in stats['client_summary'][:8]:
        print(f"  {c['client_name'][:30]:<32} {c['total_jobs']:<12,} {c['late_jobs']:<12,} {c['late_pct']:>5.1f}%     {c['avg_billing_lag']}d")
    print("  " + "-" * 76)

    print(f"\n[4/4] Audit evaluation complete! Output Excel ready for export & management review.")
    print(f"File Path: {output_file}\n")
    return output_file


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SP Shah Import Job Register Auditor")
    parser.add_argument("input_file", nargs="?", default=None, help="Path to input .xlsx or .xls file")
    parser.add_argument("-o", "--output", default=None, help="Path to output Excel file")
    args = parser.parse_args()

    target_file = args.input_file
    if not target_file:
        # Default to existing files in current dir
        candidates = [
            os.path.join(current_dir, "Job Register Formated.xlsx"),
            os.path.join(current_dir, "Job Register Formated - Copy.xls"),
            "/tmp/job_reg.xlsx"
        ]
        for c in candidates:
            if os.path.exists(c):
                target_file = c
                break

    if not target_file:
        print("[ERROR] No input file provided or found. Please specify file path:")
        print("Usage: python3 run_audit.py <path_to_job_register.xlsx>")
        sys.exit(1)

    run_audit(target_file, args.output)

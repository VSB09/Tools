#!/usr/bin/env python3
"""
SP Shah - Import Job Register Auditor Web Dashboard
===================================================
A modern, responsive web application for automated evaluation of Import Job Registers:
- Upload any .xlsx or .xls batch
- Real-time Late Billing (> 2 Days) and Revenue TAT Analytics
- Interactive KPI cards & filterable tables
- Direct download of auditor-grade 7-sheet formatted Excel workbook
"""

import os
import sys
import shutil
import tempfile
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# Setup paths
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from job_register_processor import (
    load_file_rows,
    parse_job_register,
    generate_auditor_excel_workbook
)

app = FastAPI(title="SP Shah Job Register Auditor", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global in-memory cache for latest processed batch
STATE = {
    "filename": "Job Register Formated.xlsx",
    "parsed_jobs": [],
    "enriched_rows": [],
    "stats": {},
    "last_report_path": None
}


def load_default_dataset():
    candidates = [
        os.path.join(current_dir, "Job Register Formated.xlsx"),
        "/tmp/job_reg.xlsx",
        os.path.join(current_dir, "Job Register Formated - Copy.xls")
    ]
    for c in candidates:
        if os.path.exists(c):
            try:
                rows = load_file_rows(c)
                parsed_jobs, enriched_rows, stats = parse_job_register(rows)
                STATE["filename"] = os.path.basename(c)
                STATE["parsed_jobs"] = parsed_jobs
                STATE["enriched_rows"] = enriched_rows
                STATE["stats"] = stats
                
                # Check for existing report or prepare path
                out_dir = os.path.join(current_dir, "Outputs")
                os.makedirs(out_dir, exist_ok=True)
                rep_path = os.path.join(out_dir, "SP_Shah_Job_Register_Audit_Report.xlsx")
                STATE["last_report_path"] = rep_path
                return True
            except Exception as e:
                print("Error loading default dataset:", e)
    return False


# Default dataset will be loaded lazily on first request


HTML_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SP Shah — Import Job Register Audit & Late Billing Evaluator</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #1B365D;
            --primary-light: #2C4D75;
            --accent: #0066CC;
            --bg: #F4F6F9;
            --surface: #FFFFFF;
            --text-main: #1E293B;
            --text-muted: #64748B;
            --border: #E2E8F0;
            --red-bg: #FEE2E2;
            --red-text: #991B1B;
            --green-bg: #DCFCE7;
            --green-text: #166534;
            --yellow-bg: #FEF3C7;
            --yellow-text: #92400E;
            --blue-bg: #E0F2FE;
            --blue-text: #075985;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            line-height: 1.5;
            padding-bottom: 60px;
        }
        header {
            background: linear-gradient(135deg, var(--primary) 0%, #0F2537 100%);
            color: white;
            padding: 24px 36px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        .header-container {
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }
        .title-group h1 {
            font-size: 22px;
            font-weight: 700;
            letter-spacing: -0.5px;
        }
        .title-group p {
            font-size: 13px;
            opacity: 0.85;
            margin-top: 4px;
        }
        .header-actions {
            display: flex;
            gap: 12px;
            align-items: center;
        }
        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 10px 18px;
            font-size: 13px;
            font-weight: 600;
            border-radius: 6px;
            cursor: pointer;
            border: none;
            transition: all 0.2s;
            text-decoration: none;
        }
        .btn-success {
            background-color: #10B981;
            color: white;
        }
        .btn-success:hover { background-color: #059669; }
        .btn-outline {
            background-color: rgba(255,255,255,0.12);
            color: white;
            border: 1px solid rgba(255,255,255,0.25);
        }
        .btn-outline:hover { background-color: rgba(255,255,255,0.25); }
        .btn-primary {
            background-color: var(--accent);
            color: white;
        }
        .btn-primary:hover { background-color: #0052A3; }
        
        .main-container {
            max-width: 1400px;
            margin: 24px auto 0;
            padding: 0 24px;
        }

        /* Upload Banner */
        .upload-card {
            background: var(--surface);
            border: 2px dashed var(--border);
            border-radius: 10px;
            padding: 24px;
            text-align: center;
            margin-bottom: 24px;
            transition: border-color 0.2s;
        }
        .upload-card:hover { border-color: var(--accent); }
        .upload-card.dragover { border-color: #10B981; background-color: #F0FDF4; }
        .upload-btn-wrapper {
            margin-top: 12px;
            display: inline-block;
        }
        .file-input { display: none; }

        /* KPI Cards */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .kpi-card {
            background: var(--surface);
            padding: 18px;
            border-radius: 8px;
            border: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .kpi-title {
            font-size: 12px;
            font-weight: 600;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .kpi-value {
            font-size: 24px;
            font-weight: 700;
            color: var(--primary);
            margin: 8px 0 4px;
        }
        .kpi-sub {
            font-size: 12px;
            color: var(--text-muted);
        }
        .kpi-badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
            margin-top: 4px;
        }
        .badge-red { background: var(--red-bg); color: var(--red-text); }
        .badge-green { background: var(--green-bg); color: var(--green-text); }
        .badge-yellow { background: var(--yellow-bg); color: var(--yellow-text); }
        .badge-blue { background: var(--blue-bg); color: var(--blue-text); }

        /* Tabs & Content */
        .tabs-header {
            display: flex;
            gap: 4px;
            border-bottom: 2px solid var(--border);
            margin-bottom: 16px;
            background: var(--surface);
            border-radius: 8px 8px 0 0;
            padding: 8px 16px 0;
        }
        .tab-btn {
            padding: 10px 18px;
            font-size: 13px;
            font-weight: 600;
            color: var(--text-muted);
            background: none;
            border: none;
            cursor: pointer;
            border-bottom: 2px solid transparent;
            margin-bottom: -2px;
            transition: all 0.2s;
        }
        .tab-btn:hover { color: var(--primary); }
        .tab-btn.active {
            color: var(--accent);
            border-bottom-color: var(--accent);
        }
        .tab-content { display: none; }
        .tab-content.active { display: block; }

        /* Tables & Filter Bar */
        .table-card {
            background: var(--surface);
            border-radius: 8px;
            border: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            overflow: hidden;
        }
        .filter-bar {
            padding: 16px 20px;
            border-bottom: 1px solid var(--border);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
            background: #FAFAFA;
        }
        .search-input {
            padding: 8px 14px;
            font-size: 13px;
            border: 1px solid var(--border);
            border-radius: 6px;
            width: 300px;
        }
        .select-filter {
            padding: 8px 12px;
            font-size: 13px;
            border: 1px solid var(--border);
            border-radius: 6px;
            background: white;
        }
        .table-wrapper {
            overflow-x: auto;
            max-height: 600px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            text-align: left;
        }
        th {
            background: #F1F5F9;
            color: var(--text-main);
            font-weight: 600;
            padding: 12px 14px;
            position: sticky;
            top: 0;
            z-index: 2;
            border-bottom: 1px solid var(--border);
            white-space: nowrap;
        }
        td {
            padding: 10px 14px;
            border-bottom: 1px solid var(--border);
            white-space: nowrap;
        }
        tr:hover { background-color: #F8FAFC; }

        .spinner {
            display: inline-block;
            width: 16px;
            height: 16px;
            border: 2px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: #fff;
            animation: spin 0.8s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <header>
        <div class="header-container">
            <div class="title-group">
                <h1>SP SHAH & CO. — IMPORT JOB REGISTER AUDIT</h1>
                <p>Internal Audit & Automation System | Evaluator: Late Billing = Difference (Bill Date - Delivery Date) > 2 Days</p>
            </div>
            <div class="header-actions">
                <a href="/api/export" class="btn btn-success" id="exportBtn">
                    📥 Export Excel Audit Report
                </a>
                <button class="btn btn-outline" onclick="loadDefault()">
                    🔄 Reload Base File
                </button>
            </div>
        </div>
    </header>

    <div class="main-container">
        <!-- Upload Card -->
        <div class="upload-card" id="dropZone">
            <h3 style="font-size: 15px; margin-bottom: 6px;">Upload New Job Register Batch (.xlsx or .xls)</h3>
            <p style="font-size: 13px; color: var(--text-muted);">
                Drag and drop your port export file here, or click to browse. The engine will auto-detect columns, client groups, and calculate late billing.
            </p>
            <div class="upload-btn-wrapper">
                <button class="btn btn-primary" onclick="document.getElementById('fileInput').click()">
                    Browse File
                </button>
                <input type="file" id="fileInput" class="file-input" accept=".xlsx,.xls,.xlsm" onchange="handleFileUpload(this.files[0])">
            </div>
            <span id="uploadStatus" style="display: block; margin-top: 8px; font-size: 12px; font-weight: 600; color: var(--accent);"></span>
        </div>

        <!-- KPI Grid -->
        <div class="kpi-grid" id="kpiContainer">
            <div class="kpi-card">
                <div class="kpi-title">Total Import Jobs</div>
                <div class="kpi-value" id="kpiTotalJobs">--</div>
                <div class="kpi-sub" id="kpiFilename">File: --</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Late Billing (> 2 Days)</div>
                <div class="kpi-value" id="kpiLateCount" style="color: var(--red-text);">--</div>
                <div class="kpi-badge badge-red" id="kpiLatePct">--% SLA Breach</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">On-Time Billing (<= 2 Days)</div>
                <div class="kpi-value" id="kpiOnTimeCount" style="color: var(--green-text);">--</div>
                <div class="kpi-badge badge-green" id="kpiOnTimePct">--% Compliant</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Unbilled Deliveries</div>
                <div class="kpi-value" id="kpiUnbilledCount" style="color: var(--yellow-text);">--</div>
                <div class="kpi-badge badge-yellow">Unbilled Revenue Risk</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Avg Billing Turnaround</div>
                <div class="kpi-value" id="kpiAvgLag">--</div>
                <div class="kpi-sub">Days from Delivery to Bill</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Total Duty Cleared</div>
                <div class="kpi-value" id="kpiTotalDuty" style="font-size: 20px;">--</div>
                <div class="kpi-sub">Customs Duty Processed</div>
            </div>
        </div>

        <!-- Tabs Header -->
        <div class="tabs-header">
            <button class="tab-btn active" onclick="switchTab('lateBillingTab')">🚨 Late Billing Register</button>
            <button class="tab-btn" onclick="switchTab('allJobsTab')">📋 All Jobs Register</button>
            <button class="tab-btn" onclick="switchTab('agingTab')">⏳ Delay Aging Buckets</button>
            <button class="tab-btn" onclick="switchTab('clientTab')">🏢 Client Portfolio Analysis</button>
            <button class="tab-btn" onclick="switchTab('unbilledTab')">⚠️ Unbilled Deliveries Risk</button>
        </div>

        <!-- Tab 1: Late Billing Register -->
        <div class="tab-content active" id="lateBillingTab">
            <div class="table-card">
                <div class="filter-bar">
                    <input type="text" id="lateSearch" class="search-input" placeholder="Search by Client, Job No, B/L..." oninput="filterLateTable()">
                    <select id="delayFilter" class="select-filter" onchange="filterLateTable()">
                        <option value="all">All Late Billing (> 2 Days)</option>
                        <option value="5">Delayed > 5 Days</option>
                        <option value="10">Delayed > 10 Days</option>
                        <option value="15">Delayed > 15 Days</option>
                        <option value="30">Delayed > 30 Days (Critical)</option>
                    </select>
                    <span id="lateTableCount" style="font-size: 12px; color: var(--text-muted); font-weight: 500;">Showing -- jobs</span>
                </div>
                <div class="table-wrapper">
                    <table id="lateTable">
                        <thead>
                            <tr>
                                <th>Group Sr.No</th>
                                <th>Job No</th>
                                <th>Client Name</th>
                                <th>Bill Date</th>
                                <th>Delivery Date</th>
                                <th>Diff (Days)</th>
                                <th>Status</th>
                                <th>Auditor Comments</th>
                                <th>Duty Paid (₹)</th>
                                <th>Assessable Val (₹)</th>
                                <th>B/L No</th>
                                <th>Bill No</th>
                            </tr>
                        </thead>
                        <tbody id="lateTableBody">
                            <tr><td colspan="12" style="text-align: center; padding: 24px;">Loading data...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Tab 2: All Jobs Register -->
        <div class="tab-content" id="allJobsTab">
            <div class="table-card">
                <div class="filter-bar">
                    <input type="text" id="allSearch" class="search-input" placeholder="Search all jobs..." oninput="filterAllTable()">
                    <select id="statusFilter" class="select-filter" onchange="filterAllTable()">
                        <option value="all">All Billing Statuses</option>
                        <option value="Late Billing">Late Billing</option>
                        <option value="On-Time Billing">On-Time Billing</option>
                        <option value="Same-Day Billing">Same-Day Billing</option>
                        <option value="Unbilled Delivery">Unbilled Delivery</option>
                        <option value="Advance Billed">Advance Billed</option>
                    </select>
                    <span id="allTableCount" style="font-size: 12px; color: var(--text-muted); font-weight: 500;">Showing -- jobs</span>
                </div>
                <div class="table-wrapper">
                    <table id="allTable">
                        <thead>
                            <tr>
                                <th>Group Sr.No</th>
                                <th>Job No</th>
                                <th>Client Name</th>
                                <th>Bill Date</th>
                                <th>Delivery Date</th>
                                <th>Diff (Days)</th>
                                <th>Billing Status</th>
                                <th>Duty Paid (₹)</th>
                                <th>Assessable Val (₹)</th>
                                <th>Bill No</th>
                                <th>B/L No</th>
                            </tr>
                        </thead>
                        <tbody id="allTableBody">
                            <tr><td colspan="11" style="text-align: center; padding: 24px;">Loading data...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Tab 3: Aging Buckets -->
        <div class="tab-content" id="agingTab">
            <div class="table-card">
                <div class="table-wrapper">
                    <table>
                        <thead>
                            <tr>
                                <th>Delay Aging Bracket</th>
                                <th>Job Count</th>
                                <th>% of Total</th>
                                <th>Audit Severity Level</th>
                                <th>Recommended Internal Control Procedure</th>
                            </tr>
                        </thead>
                        <tbody id="agingTableBody"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Tab 4: Client Portfolio -->
        <div class="tab-content" id="clientTab">
            <div class="table-card">
                <div class="filter-bar">
                    <input type="text" id="clientSearch" class="search-input" placeholder="Search clients..." oninput="filterClientTable()">
                    <span id="clientTableCount" style="font-size: 12px; color: var(--text-muted); font-weight: 500;"></span>
                </div>
                <div class="table-wrapper">
                    <table id="clientTable">
                        <thead>
                            <tr>
                                <th>Client Name</th>
                                <th>Total Jobs</th>
                                <th>On-Time Jobs</th>
                                <th>Late Jobs</th>
                                <th>Late Billing %</th>
                                <th>Unbilled Jobs</th>
                                <th>Avg Delay (Days)</th>
                                <th>Total Duty Paid (₹)</th>
                                <th>Risk Classification</th>
                            </tr>
                        </thead>
                        <tbody id="clientTableBody"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Tab 5: Unbilled Deliveries -->
        <div class="tab-content" id="unbilledTab">
            <div class="table-card">
                <div class="table-wrapper">
                    <table>
                        <thead>
                            <tr>
                                <th>Group Sr.No</th>
                                <th>Job No</th>
                                <th>Client Name</th>
                                <th>Delivery Date</th>
                                <th>Duty Paid (₹)</th>
                                <th>Assessable Val (₹)</th>
                                <th>B/L No</th>
                                <th>Auditor Risk Notice</th>
                            </tr>
                        </thead>
                        <tbody id="unbilledTableBody"></tbody>
                    </table>
                </div>
            </div>
        </div>

    </div>

    <script>
        let GLOBAL_DATA = null;

        async function fetchAuditData() {
            try {
                const res = await fetch('/api/data');
                const data = await res.json();
                GLOBAL_DATA = data;
                renderDashboard(data);
            } catch (e) {
                console.error("Error fetching data:", e);
            }
        }

        function formatCurrency(num) {
            if (!num && num !== 0) return '₹0.00';
            return '₹' + Number(num).toLocaleString('en-IN', { maximumFractionDigits: 2 });
        }

        function renderDashboard(data) {
            const stats = data.stats;
            document.getElementById('kpiTotalJobs').innerText = stats.total_jobs ? stats.total_jobs.toLocaleString() : '--';
            document.getElementById('kpiFilename').innerText = 'File: ' + data.filename;
            document.getElementById('kpiLateCount').innerText = stats.late_billing_count ? stats.late_billing_count.toLocaleString() : '--';
            document.getElementById('kpiLatePct').innerText = stats.late_billing_pct + '% SLA Breach';
            document.getElementById('kpiOnTimeCount').innerText = stats.on_time_count ? stats.on_time_count.toLocaleString() : '--';
            document.getElementById('kpiOnTimePct').innerText = stats.on_time_pct + '% Compliant';
            document.getElementById('kpiUnbilledCount').innerText = stats.unbilled_count ? stats.unbilled_count.toLocaleString() : '0';
            document.getElementById('kpiAvgLag').innerText = stats.avg_billing_lag ? stats.avg_billing_lag + ' d' : '--';
            document.getElementById('kpiTotalDuty').innerText = formatCurrency(stats.total_duty_paid);

            renderLateTable(data.jobs.filter(j => j.billing_status === 'Late Billing'));
            renderAllTable(data.jobs);
            renderAgingTable(stats.buckets, stats.total_jobs);
            renderClientTable(stats.client_summary);
            renderUnbilledTable(data.jobs.filter(j => j.billing_status === 'Unbilled Delivery'));
        }

        function renderLateTable(jobs) {
            const tbody = document.getElementById('lateTableBody');
            document.getElementById('lateTableCount').innerText = `Showing ${jobs.length} jobs`;
            if (jobs.length === 0) {
                tbody.innerHTML = '<tr><td colspan="12" style="text-align: center; padding: 24px;">No matching late jobs found.</td></tr>';
                return;
            }
            tbody.innerHTML = jobs.slice(0, 300).map(j => `
                <tr>
                    <td style="font-weight: 600;">${j.group_sr_no}</td>
                    <td>${j.job_no}</td>
                    <td><strong>${j.client_name}</strong></td>
                    <td>${j.bill_date || ''}</td>
                    <td>${j.delivery_date || ''}</td>
                    <td><span class="kpi-badge badge-red">${j.difference_days} d</span></td>
                    <td><span class="kpi-badge badge-red">${j.billing_status}</span></td>
                    <td style="color: #991B1B; font-weight: 500;">${j.comments}</td>
                    <td style="text-align: right;">${formatCurrency(j.duty_paid)}</td>
                    <td style="text-align: right;">${formatCurrency(j.assessable_val)}</td>
                    <td>${j.bl_no}</td>
                    <td>${j.bill_no}</td>
                </tr>
            `).join('');
        }

        function filterLateTable() {
            if (!GLOBAL_DATA) return;
            const q = document.getElementById('lateSearch').value.toLowerCase();
            const delay = document.getElementById('delayFilter').value;
            let filtered = GLOBAL_DATA.jobs.filter(j => j.billing_status === 'Late Billing');

            if (delay !== 'all') {
                const dInt = parseInt(delay);
                filtered = filtered.filter(j => j.difference_days > dInt);
            }
            if (q) {
                filtered = filtered.filter(j => 
                    (j.client_name && j.client_name.toLowerCase().includes(q)) ||
                    (j.job_no && String(j.job_no).toLowerCase().includes(q)) ||
                    (j.bl_no && j.bl_no.toLowerCase().includes(q))
                );
            }
            renderLateTable(filtered);
        }

        function renderAllTable(jobs) {
            const tbody = document.getElementById('allTableBody');
            document.getElementById('allTableCount').innerText = `Showing ${jobs.length} jobs`;
            tbody.innerHTML = jobs.slice(0, 300).map(j => {
                let badge = 'badge-blue';
                if (j.billing_status === 'Late Billing') badge = 'badge-red';
                else if (j.billing_status.includes('On-Time') || j.billing_status.includes('Same-Day')) badge = 'badge-green';
                else if (j.billing_status === 'Unbilled Delivery') badge = 'badge-yellow';

                return `
                    <tr>
                        <td style="font-weight: 600;">${j.group_sr_no}</td>
                        <td>${j.job_no}</td>
                        <td>${j.client_name}</td>
                        <td>${j.bill_date || ''}</td>
                        <td>${j.delivery_date || ''}</td>
                        <td>${j.difference_days !== null ? j.difference_days + ' d' : '-'}</td>
                        <td><span class="kpi-badge ${badge}">${j.billing_status}</span></td>
                        <td style="text-align: right;">${formatCurrency(j.duty_paid)}</td>
                        <td style="text-align: right;">${formatCurrency(j.assessable_val)}</td>
                        <td>${j.bill_no}</td>
                        <td>${j.bl_no}</td>
                    </tr>
                `;
            }).join('');
        }

        function filterAllTable() {
            if (!GLOBAL_DATA) return;
            const q = document.getElementById('allSearch').value.toLowerCase();
            const stat = document.getElementById('statusFilter').value;
            let filtered = GLOBAL_DATA.jobs;

            if (stat !== 'all') {
                filtered = filtered.filter(j => j.billing_status === stat);
            }
            if (q) {
                filtered = filtered.filter(j => 
                    (j.client_name && j.client_name.toLowerCase().includes(q)) ||
                    (j.job_no && String(j.job_no).toLowerCase().includes(q)) ||
                    (j.bl_no && j.bl_no.toLowerCase().includes(q))
                );
            }
            renderAllTable(filtered);
        }

        function renderAgingTable(buckets, total) {
            const tbody = document.getElementById('agingTableBody');
            const descriptions = {
                '0-2 Days (On-Time)': ['LOW / COMPLIANT', 'Normal operational clearance. Meets company SLA policy.'],
                '3-5 Days (Low Delay)': ['MODERATE', 'Monitor billing desk queue. Ensure bills raised within 48 hours.'],
                '6-10 Days (Moderate Delay)': ['ELEVATED', 'Review documentation handoff between CFS delivery and finance.'],
                '11-20 Days (Significant Delay)': ['HIGH', 'Mandatory explanation from operations manager before billing.'],
                '21-30 Days (High Delay)': ['CRITICAL', 'Direct audit exception. Revenue cycle delay affecting cash flow.'],
                '> 30 Days (Severe Delay)': ['SEVERE SLA BREACH', 'Escalate to TCWG/Directors. Investigate client disputed billing.'],
                'Advance Billed (< 0d)': ['PROCESS ANOMALY', 'Verify cargo dispatch confirmation prior to invoice release.'],
                'Unbilled Revenue Exposure': ['UNBILLED RISK', 'Immediate unbilled revenue reconciliation and invoice dispatch.'],
                'Date Anomaly / Typo': ['DATA QUALITY', 'Audit log fix: correct manual typing of delivery/billing dates.']
            };

            tbody.innerHTML = Object.entries(buckets).map(([k, cnt]) => {
                const pct = total ? ((cnt / total) * 100).toFixed(1) + '%' : '0%';
                const desc = descriptions[k] || ['AUDIT REVIEW', 'Investigate operational workflow.'];
                return `
                    <tr>
                        <td><strong>${k}</strong></td>
                        <td style="font-weight: 600;">${cnt.toLocaleString()}</td>
                        <td>${pct}</td>
                        <td><span class="kpi-badge ${k.includes('0-2') ? 'badge-green' : (k.includes('3-5') ? 'badge-blue' : 'badge-red')}">${desc[0]}</span></td>
                        <td>${desc[1]}</td>
                    </tr>
                `;
            }).join('');
        }

        function renderClientTable(clients) {
            const tbody = document.getElementById('clientTableBody');
            document.getElementById('clientTableCount').innerText = `${clients.length} Clients Total`;
            tbody.innerHTML = clients.slice(0, 100).map(c => `
                <tr>
                    <td><strong>${c.client_name}</strong></td>
                    <td style="text-align: center;">${c.total_jobs.toLocaleString()}</td>
                    <td style="text-align: center; color: var(--green-text); font-weight: 600;">${c.on_time_jobs}</td>
                    <td style="text-align: center; color: var(--red-text); font-weight: 600;">${c.late_jobs}</td>
                    <td style="text-align: center;"><strong>${c.late_pct}%</strong></td>
                    <td style="text-align: center;">${c.unbilled_jobs}</td>
                    <td style="text-align: center;">${c.avg_billing_lag} d</td>
                    <td style="text-align: right;">${formatCurrency(c.total_duty)}</td>
                    <td><span class="kpi-badge ${c.risk_tier === 'HIGH RISK' ? 'badge-red' : (c.risk_tier === 'COMPLIANT' ? 'badge-green' : 'badge-yellow')}">${c.risk_tier}</span></td>
                </tr>
            `).join('');
        }

        function filterClientTable() {
            if (!GLOBAL_DATA) return;
            const q = document.getElementById('clientSearch').value.toLowerCase();
            const filtered = GLOBAL_DATA.stats.client_summary.filter(c => c.client_name.toLowerCase().includes(q));
            renderClientTable(filtered);
        }

        function renderUnbilledTable(jobs) {
            const tbody = document.getElementById('unbilledTableBody');
            tbody.innerHTML = jobs.map(j => `
                <tr>
                    <td style="font-weight: 600;">${j.group_sr_no}</td>
                    <td>${j.job_no}</td>
                    <td><strong>${j.client_name}</strong></td>
                    <td>${j.delivery_date}</td>
                    <td style="text-align: right;">${formatCurrency(j.duty_paid)}</td>
                    <td style="text-align: right;">${formatCurrency(j.assessable_val)}</td>
                    <td>${j.bl_no}</td>
                    <td style="color: #92400E; font-weight: 600;">${j.comments}</td>
                </tr>
            `).join('');
        }

        function switchTab(tabId) {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            event.target.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        }

        async function handleFileUpload(file) {
            if (!file) return;
            const status = document.getElementById('uploadStatus');
            status.innerHTML = '<span class="spinner"></span> Processing file... please wait.';

            const formData = new FormData();
            formData.append('file', file);

            try {
                const res = await fetch('/api/upload', {
                    method: 'POST',
                    body: formData
                });
                const data = await res.json();
                GLOBAL_DATA = data;
                renderDashboard(data);
                status.innerHTML = `✅ Successfully evaluated batch: <strong>${data.filename}</strong> (${data.stats.total_jobs.toLocaleString()} jobs parsed).`;
            } catch (e) {
                status.innerHTML = `❌ Error processing file: ${e.message}`;
            }
        }

        async function loadDefault() {
            const status = document.getElementById('uploadStatus');
            status.innerHTML = '<span class="spinner"></span> Reloading base dataset...';
            await fetchAuditData();
            status.innerHTML = `✅ Loaded base file: <strong>${GLOBAL_DATA.filename}</strong>`;
        }

        // Init Drag & Drop
        const dropZone = document.getElementById('dropZone');
        ['dragenter', 'dragover'].forEach(name => {
            dropZone.addEventListener(name, (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
        });
        ['dragleave', 'drop'].forEach(name => {
            dropZone.addEventListener(name, (e) => { e.preventDefault(); dropZone.classList.remove('dragover'); });
        });
        dropZone.addEventListener('drop', (e) => {
            if (e.dataTransfer.files.length) {
                handleFileUpload(e.dataTransfer.files[0]);
            }
        });

        // Load initially
        fetchAuditData();
    </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


@app.get("/api/data")
async def get_data():
    if not STATE["parsed_jobs"]:
        load_default_dataset()
    
    # Format dates as strings for JSON serialization
    serialized_jobs = []
    for j in STATE["parsed_jobs"]:
        sj = dict(j)
        for k in ("bill_date", "delivery_date", "recd_date", "bl_date", "igm_date", "be_date"):
            if sj.get(k):
                sj[k] = sj[k].strftime("%d/%m/%Y")
        serialized_jobs.append(sj)

    return {
        "filename": STATE["filename"],
        "stats": STATE["stats"],
        "jobs": serialized_jobs
    }


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".xlsx", ".xls", ".xlsm"):
        raise HTTPException(status_code=400, detail="Only .xlsx or .xls files supported.")

    # Save to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    try:
        rows = load_file_rows(tmp_path)
        parsed_jobs, enriched_rows, stats = parse_job_register(rows)

        # Generate fresh report in Outputs/
        out_dir = os.path.join(current_dir, "Outputs")
        os.makedirs(out_dir, exist_ok=True)
        report_path = os.path.join(out_dir, f"Audit_Report_{filename}.xlsx")
        generate_auditor_excel_workbook(parsed_jobs, enriched_rows, stats, report_path)

        # Update state
        STATE["filename"] = filename
        STATE["parsed_jobs"] = parsed_jobs
        STATE["enriched_rows"] = enriched_rows
        STATE["stats"] = stats
        STATE["last_report_path"] = report_path

        # Return serialized data
        serialized_jobs = []
        for j in parsed_jobs:
            sj = dict(j)
            for k in ("bill_date", "delivery_date", "recd_date", "bl_date", "igm_date", "be_date"):
                if sj.get(k):
                    sj[k] = sj[k].strftime("%d/%m/%Y")
            serialized_jobs.append(sj)

        return {
            "filename": filename,
            "stats": stats,
            "jobs": serialized_jobs
        }
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.get("/api/export")
async def export_excel():
    rep_path = STATE.get("last_report_path")
    if not rep_path or not os.path.exists(rep_path):
        # Generate on the fly
        if not STATE["parsed_jobs"]:
            load_default_dataset()
        out_dir = os.path.join(current_dir, "Outputs")
        os.makedirs(out_dir, exist_ok=True)
        rep_path = os.path.join(out_dir, "SP_Shah_Job_Register_Audit_Report.xlsx")
        generate_auditor_excel_workbook(STATE["parsed_jobs"], STATE["enriched_rows"], STATE["stats"], rep_path)
        STATE["last_report_path"] = rep_path

    return FileResponse(
        rep_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(rep_path)
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8055))
    print(f"Starting SP Shah Job Register Audit Web Tool on http://localhost:{port}")
    uvicorn.run(app, host="0.0.0.0", port=port)

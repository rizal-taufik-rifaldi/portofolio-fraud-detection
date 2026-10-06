#!/usr/bin/env python3
"""
generate_excel_report.py

This script processes the sample transaction data and fraud detection findings,
generating a highly professional, beautifully styled, and executive-ready Excel report
named 'report/Laporan_Audit_Fraud_PaySim.xlsx' using openpyxl.

The Excel file contains 4 sheets:
1. Data Transaksi (Max 1,000 rows with Indonesian headers and conditional formatting)
2. Temuan Fraud (All findings, rule-based color coding, summary box)
3. Ringkasan per Tipe (Aggregated pivot-like table with nominal bar chart and fraud pie chart)
4. Dashboard Eksekutif (5 KPI cards, dynamic main insights, fraud by type bar chart, daily volume line chart)
"""

import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference

# --- Global Style Constants ---
FONT_NAME = "Segoe UI"
COLOR_NAVY = "1F497D"      # Dark Navy for main headers
COLOR_ACCENT = "E8EEF5"    # Very light navy/blue for total/accent rows
COLOR_LIGHT_GRAY = "F2F4F7" # For KPI card backgrounds
COLOR_BORDER = "D9D9D9"    # Border color

# Fonts
font_title = Font(name=FONT_NAME, size=16, bold=True, color=COLOR_NAVY)
font_section = Font(name=FONT_NAME, size=12, bold=True, color=COLOR_NAVY)
font_header = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
font_data = Font(name=FONT_NAME, size=10)
font_bold = Font(name=FONT_NAME, size=10, bold=True)
font_card_label = Font(name=FONT_NAME, size=9, bold=True, color="595959")
font_card_value = Font(name=FONT_NAME, size=16, bold=True, color=COLOR_NAVY)

# Fills
fill_header = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
fill_accent = PatternFill(start_color=COLOR_ACCENT, end_color=COLOR_ACCENT, fill_type="solid")
fill_card = PatternFill(start_color=COLOR_LIGHT_GRAY, end_color=COLOR_LIGHT_GRAY, fill_type="solid")
fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

# Borders
border_thin = Border(
    left=Side(style="thin", color=COLOR_BORDER),
    right=Side(style="thin", color=COLOR_BORDER),
    top=Side(style="thin", color=COLOR_BORDER),
    bottom=Side(style="thin", color=COLOR_BORDER)
)
border_total = Border(
    left=Side(style="thin", color=COLOR_BORDER),
    right=Side(style="thin", color=COLOR_BORDER),
    top=Side(style="thin", color=COLOR_BORDER),
    bottom=Side(style="double", color=COLOR_NAVY)
)

# Alignments
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

# Function to auto-fit columns
def autofit_cols(ws, min_col=1, max_col=None, padding=3, min_width=12):
    max_col = max_col or ws.max_column
    for col_idx in range(min_col, max_col + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[col_letter]:
            if cell.row == 1 and cell.coordinate in ws.merged_cells:
                continue # Skip title headers in row 1 for calculation
            val_str = str(cell.value or "")
            # If cell has formula, use a default length
            if val_str.startswith("="):
                max_len = max(max_len, 10)
            else:
                max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + padding, min_width)

# --- SHEET 1: Data Transaksi ---
def create_sheet_data_transaksi(wb, df):
    print("Creating sheet: Data Transaksi...")
    ws = wb.create_sheet(title="Data Transaksi")

    # Grid lines enabled
    try:
        ws.sheet_view.showGridLines = True
    except AttributeError:
        pass

    # Take max 1000 rows
    df_sub = df.head(1000).copy()

    # Translate headers
    indonesian_headers = [
        "Langkah Waktu", "Tipe", "Nominal", "Akun Pengirim",
        "Saldo Awal Pengirim", "Saldo Akhir Pengirim", "Akun Penerima",
        "Saldo Awal Penerima", "Saldo Akhir Penerima", "Status Fraud", "Terflag Sistem"
    ]

    # Write headers
    for col_idx, header in enumerate(indonesian_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_thin

    # Format rules for values
    red_fill = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid") # soft red-orange
    red_font = Font(name=FONT_NAME, size=10, color="9C0006", bold=True)

    # Write Data
    for row_idx, row_val in enumerate(df_sub.values, 2):
        is_fraud = row_val[9] == 1  # isFraud column is index 9
        row_fill = red_fill if is_fraud else fill_white
        row_font = red_font if is_fraud else font_data

        for col_idx, val in enumerate(row_val, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = row_font
            cell.fill = row_fill
            cell.border = border_thin

            # Formats based on data types
            if col_idx in [1, 10, 11]:  # step, isFraud, isFlaggedFraud
                cell.alignment = align_center
                cell.number_format = "#,##0"
            elif col_idx in [2, 4, 7]:  # type, nameOrig, nameDest
                cell.alignment = align_left
            elif col_idx in [3, 5, 6, 8, 9]:  # amount and balances
                cell.alignment = align_right
                cell.number_format = "#,##0.00"

    # Freeze first row
    ws.freeze_panes = "A2"
    autofit_cols(ws, min_width=12)

# --- SHEET 2: Temuan Fraud ---
def create_sheet_temuan_fraud(wb, findings_df, df_len):
    print("Creating sheet: Temuan Fraud...")
    ws = wb.create_sheet(title="Temuan Fraud")

    try:
        ws.sheet_view.showGridLines = True
    except AttributeError:
        pass

    # Translate rules
    rule_translations = {
        'balance_mismatch': "Ketidaksesuaian Saldo",
        'large_transaction': "Transaksi Nominal Besar",
        'zero_balance_origin': "Saldo Pengirim Terkuras",
        'suspicious_dest_account': "Akun Penerima Mencurigakan (Mule)",
        'missed_by_system': "Fraud Lolos Flag Sistem"
    }

    # Calculate rule counts for summary
    rule_counts = findings_df['rule_triggered'].value_counts()

    # Title Summary Box
    ws.cell(row=2, column=2, value="RINGKASAN TEMUAN DETEKSI FRAUD").font = font_section

    # Summary Box Headers
    summary_headers = ["Kategori Aturan Deteksi", "Jumlah Temuan", "% dari Total Data"]
    for col_offset, h in enumerate(summary_headers, 2):
        cell = ws.cell(row=4, column=col_offset, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_thin

    # Write Rule Summary Data
    summary_rules = [
        ('balance_mismatch', "Ketidaksesuaian Saldo"),
        ('large_transaction', "Transaksi Nominal Besar"),
        ('zero_balance_origin', "Saldo Pengirim Terkuras"),
        ('suspicious_dest_account', "Akun Penerima Mencurigakan (Mule)"),
        ('missed_by_system', "Fraud Lolos Flag Sistem")
    ]

    for idx, (rule_id, rule_name) in enumerate(summary_rules, 5):
        count = rule_counts.get(rule_id, 0)
        pct = (count / df_len)

        c1 = ws.cell(row=idx, column=2, value=rule_name)
        c2 = ws.cell(row=idx, column=3, value=count)
        c3 = ws.cell(row=idx, column=4, value=pct)

        for c in [c1, c2, c3]:
            c.font = font_data
            c.border = border_thin
        c1.alignment = align_left
        c2.alignment = align_right
        c2.number_format = "#,##0"
        c3.alignment = align_right
        c3.number_format = "0.00%"

    # Summary Total Row
    total_idx = 10
    c_tot_lbl = ws.cell(row=total_idx, column=2, value="Total Pemicu Aturan (Triggers)")
    c_tot_val = ws.cell(row=total_idx, column=3, value=f"=SUM(C5:C9)")
    c_tot_pct = ws.cell(row=total_idx, column=4, value=f"=SUM(D5:D9)")

    for c in [c_tot_lbl, c_tot_val, c_tot_pct]:
        c.font = font_bold
        c.fill = fill_accent
        c.border = border_total
    c_tot_lbl.alignment = align_left
    c_tot_val.alignment = align_right
    c_tot_val.number_format = "#,##0"
    c_tot_pct.alignment = align_right
    c_tot_pct.number_format = "0.00%"

    # ----------------------------------------------------
    # Table of Findings
    # ----------------------------------------------------
    start_row = 12
    ws.cell(row=start_row, column=1, value="DAFTAR TRANSAKSI MENCURIGAKAN (TEMUAN)").font = font_section

    headers = [
        "Langkah Waktu", "Tipe", "Nominal", "Akun Pengirim",
        "Saldo Awal Pengirim", "Saldo Akhir Pengirim", "Akun Penerima",
        "Saldo Awal Penerima", "Saldo Akhir Penerima", "Status Fraud", "Terflag Sistem", "Aturan Terpicu"
    ]

    header_row = start_row + 1
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_thin

    # Cell fills per rule type to visually segment
    fill_rules = {
        'balance_mismatch': PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"), # yellow
        'large_transaction': PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"), # orange
        'zero_balance_origin': PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"), # blue
        'suspicious_dest_account': PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"), # green
        'missed_by_system': PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid") # dark orange/red-fill
    }

    fill_is_fraud = PatternFill(start_color="FFD8D8", end_color="FFD8D8", fill_type="solid")
    font_is_fraud = Font(name=FONT_NAME, size=10, color="9C0006", bold=True)

    # Write findings data row-by-row
    current_row = header_row + 1
    for row_val in findings_df.values:
        rule_id = row_val[11]  # rule_triggered is at column index 11
        translated_rule = rule_translations.get(rule_id, rule_id)
        is_fraud = row_val[9] == 1

        for col_idx in range(11):
            val = row_val[col_idx]
            cell = ws.cell(row=current_row, column=col_idx+1, value=val)
            cell.font = font_is_fraud if is_fraud and col_idx == 9 else font_data
            cell.fill = fill_is_fraud if is_fraud and col_idx == 9 else fill_white
            cell.border = border_thin

            # Alignments & formats
            if col_idx in [0, 9, 10]: # step, isFraud, isFlaggedFraud
                cell.alignment = align_center
                cell.number_format = "#,##0"
            elif col_idx in [1, 3, 6]: # type, nameOrig, nameDest
                cell.alignment = align_left
            elif col_idx in [2, 4, 5, 7, 8]: # amount, balances
                cell.alignment = align_right
                cell.number_format = "#,##0.00"

        # Rule Triggered Column (Col 12)
        cell_rule = ws.cell(row=current_row, column=12, value=translated_rule)
        cell_rule.font = font_bold
        cell_rule.fill = fill_rules.get(rule_id, fill_white)
        cell_rule.border = border_thin
        cell_rule.alignment = align_left

        current_row += 1

    # Freeze rows
    ws.freeze_panes = f"A{header_row + 1}"
    autofit_cols(ws, min_width=12)

# --- SHEET 3: Ringkasan per Tipe ---
def create_sheet_ringkasan_per_tipe(wb, df):
    print("Creating sheet: Ringkasan per Tipe...")
    ws = wb.create_sheet(title="Ringkasan per Tipe")

    try:
        ws.sheet_view.showGridLines = True
    except AttributeError:
        pass

    ws.cell(row=2, column=1, value="RINGKASAN STATISTIK DAN KASUS FRAUD").font = font_title

    # ----------------------------------------------------
    # Table 1: Pie Chart Helper (Fraud vs Non-Fraud)
    # ----------------------------------------------------
    ws.cell(row=4, column=1, value="Status").font = font_bold
    ws.cell(row=4, column=2, value="Jumlah Transaksi").font = font_bold
    for c in [ws.cell(row=4, column=1), ws.cell(row=4, column=2)]:
        c.fill = fill_accent
        c.border = border_thin

    total_fraud = df['isFraud'].sum()
    total_non_fraud = len(df) - total_fraud

    ws.cell(row=5, column=1, value="Non-Fraud").border = border_thin
    ws.cell(row=5, column=2, value=total_non_fraud).border = border_thin
    ws.cell(row=5, column=2).number_format = "#,##0"

    ws.cell(row=6, column=1, value="Fraud").border = border_thin
    ws.cell(row=6, column=2, value=total_fraud).border = border_thin
    ws.cell(row=6, column=2).number_format = "#,##0"

    # ----------------------------------------------------
    # Table 2: Pivot Table per Tipe Transaksi
    # ----------------------------------------------------
    ws.cell(row=8, column=1, value="Statistik Berdasarkan Tipe Transaksi").font = font_section

    pivot_headers = [
        "Tipe Transaksi", "Total Transaksi", "Total Nominal (IDR)",
        "Jumlah Fraud", "% Rasio Fraud", "Jumlah Terflag"
    ]

    for col_idx, h in enumerate(pivot_headers, 1):
        cell = ws.cell(row=9, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_thin

    # Group and aggregate data
    grouped = df.groupby('type').agg(
        total_tx=('amount', 'count'),
        total_amt=('amount', 'sum'),
        fraud_cnt=('isFraud', 'sum'),
        flagged_cnt=('isFlaggedFraud', 'sum')
    ).reset_index()

    # Write aggregated data
    for idx, row in grouped.iterrows():
        r_idx = 10 + idx
        c1 = ws.cell(row=r_idx, column=1, value=row['type'])
        c2 = ws.cell(row=r_idx, column=2, value=int(row['total_tx']))
        c3 = ws.cell(row=r_idx, column=3, value=float(row['total_amt']))
        c4 = ws.cell(row=r_idx, column=4, value=int(row['fraud_cnt']))

        # rasio fraud formula: Jumlah Fraud / Total Transaksi
        c5 = ws.cell(row=r_idx, column=5, value=f"=D{r_idx}/B{r_idx}")
        c6 = ws.cell(row=r_idx, column=6, value=int(row['flagged_cnt']))

        for col_c in [c1, c2, c3, c4, c5, c6]:
            col_c.font = font_data
            col_c.border = border_thin

        c1.alignment = align_left
        c2.alignment = align_right
        c2.number_format = "#,##0"
        c3.alignment = align_right
        c3.number_format = "#,##0.00"
        c4.alignment = align_right
        c4.number_format = "#,##0"
        c5.alignment = align_right
        c5.number_format = "0.00%"
        c6.alignment = align_right
        c6.number_format = "#,##0"

    # Totals Row
    total_row_idx = 10 + len(grouped)
    c_tot_lbl = ws.cell(row=total_row_idx, column=1, value="Total Keseluruhan")
    c_tot_tx = ws.cell(row=total_row_idx, column=2, value=f"=SUM(B10:B{total_row_idx-1})")
    c_tot_amt = ws.cell(row=total_row_idx, column=3, value=f"=SUM(C10:C{total_row_idx-1})")
    c_tot_fr = ws.cell(row=total_row_idx, column=4, value=f"=SUM(D10:D{total_row_idx-1})")
    c_tot_rt = ws.cell(row=total_row_idx, column=5, value=f"=D{total_row_idx}/B{total_row_idx}")
    c_tot_fl = ws.cell(row=total_row_idx, column=6, value=f"=SUM(F10:F{total_row_idx-1})")

    for c in [c_tot_lbl, c_tot_tx, c_tot_amt, c_tot_fr, c_tot_rt, c_tot_fl]:
        c.font = font_bold
        c.fill = fill_accent
        c.border = border_total

    c_tot_lbl.alignment = align_left
    c_tot_tx.alignment = align_right
    c_tot_tx.number_format = "#,##0"
    c_tot_amt.alignment = align_right
    c_tot_amt.number_format = "#,##0.00"
    c_tot_fr.alignment = align_right
    c_tot_fr.number_format = "#,##0"
    c_tot_rt.alignment = align_right
    c_tot_rt.number_format = "0.00%"
    c_tot_fl.alignment = align_right
    c_tot_fl.number_format = "#,##0"

    # ----------------------------------------------------
    # Charts Layout
    # ----------------------------------------------------
    # 1. Pie Chart for Fraud vs Non-Fraud
    pie = PieChart()
    pie.title = "Proporsi Kasus Fraud"
    pie.style = 10

    data_pie = Reference(ws, min_col=2, min_row=4, max_row=6) # Col B (Jumlah) including header
    labels_pie = Reference(ws, min_col=1, min_row=5, max_row=6) # Col A (Status)
    pie.add_data(data_pie, titles_from_data=True)
    pie.set_categories(labels_pie)
    ws.add_chart(pie, "H4")

    # 2. Bar Chart for Total Nominal per Type
    bar = BarChart()
    bar.type = "col"
    bar.style = 11
    bar.title = "Total Nominal Transaksi per Tipe"
    bar.y_axis.title = "Nominal (IDR)"
    bar.x_axis.title = "Tipe Transaksi"
    bar.legend = None

    data_bar = Reference(ws, min_col=3, min_row=9, max_row=total_row_idx-1) # Col C (Nominal) including header
    cats_bar = Reference(ws, min_col=1, min_row=10, max_row=total_row_idx-1) # Col A (Tipe)
    bar.add_data(data_bar, titles_from_data=True)
    bar.set_categories(cats_bar)
    ws.add_chart(bar, "H19")

    autofit_cols(ws, min_width=12)

# --- SHEET 4: Dashboard Eksekutif ---
def create_sheet_dashboard_eksekutif(wb, df, findings_df):
    print("Creating sheet: Dashboard Eksekutif...")
    ws = wb.create_sheet(title="Dashboard Eksekutif")

    # Hide gridlines for a clean dashboard presentation
    try:
        ws.sheet_view.showGridLines = False
    except AttributeError:
        pass

    # Title
    ws.cell(row=2, column=2, value="DASHBOARD AUDIT EKSEKUTIF - PAYSIM FRAUD ANALYSIS").font = font_title

    # Calculate stats
    total_tx = len(df)
    total_amt = df['amount'].sum()
    total_fraud = df['isFraud'].sum()
    pct_fraud = (total_fraud / total_tx)

    # Fraud missed by system
    missed_fraud = len(df[(df['isFraud'] == 1) & (df['isFlaggedFraud'] == 0)])

    # ----------------------------------------------------
    # 5 KPI Cards in Rows 4 to 6
    # ----------------------------------------------------
    kpis = [
        ("TOTAL TRANSAKSI", total_tx, "#,##0", 2),
        ("TOTAL NOMINAL (IDR)", total_amt, "#,##0.00", 4),
        ("FRAUD TERDETEKSI", total_fraud, "#,##0", 6),
        ("RASIO FRAUD (%)", pct_fraud, "0.00%", 8),
        ("FRAUD LOLOS SISTEM", missed_fraud, "#,##0", 10)
    ]

    for label, val, num_fmt, col_idx in kpis:
        # Merge label (Row 4)
        ws.merge_cells(start_row=4, start_column=col_idx, end_row=4, end_column=col_idx+1)
        ws.cell(row=4, column=col_idx, value=label).font = font_card_label
        ws.cell(row=4, column=col_idx).alignment = align_center
        ws.cell(row=4, column=col_idx).fill = fill_card

        # Merge value (Row 5-6)
        ws.merge_cells(start_row=5, start_column=col_idx, end_row=6, end_column=col_idx+1)
        ws.cell(row=5, column=col_idx, value=val).font = font_card_value
        ws.cell(row=5, column=col_idx).alignment = align_center
        ws.cell(row=5, column=col_idx).fill = fill_card
        ws.cell(row=5, column=col_idx).number_format = num_fmt

        # Apply border around card cells
        for r in range(4, 7):
            for c in range(col_idx, col_idx + 2):
                ws.cell(row=r, column=c).border = border_thin

    # ----------------------------------------------------
    # Key Insights Section (Rows 8 to 14, Columns B to G)
    # ----------------------------------------------------
    # Headers
    ws.merge_cells(start_row=8, start_column=2, end_row=8, end_column=7)
    cell_ins_title = ws.cell(row=8, column=2, value="TEMUAN UTAMA & INSIGHT AUDIT")
    cell_ins_title.font = font_header
    cell_ins_title.fill = fill_header
    cell_ins_title.alignment = align_center
    cell_ins_title.border = border_thin

    # Border wrapper for findings body
    for r in range(8, 15):
        for c in range(2, 8):
            ws.cell(row=r, column=c).border = border_thin
            if r > 8:
                ws.cell(row=r, column=c).fill = fill_white

    # Calculate values dynamically for insights
    mismatch_count = len(findings_df[findings_df['rule_triggered'] == 'balance_mismatch'])
    mismatch_pct = (mismatch_count / total_tx) * 100
    mule_count = len(findings_df[findings_df['rule_triggered'] == 'suspicious_dest_account']['nameDest'].unique())
    missed_pct = (missed_fraud / total_fraud * 100) if total_fraud > 0 else 100.0

    insight_1 = f"• Kerentanan Saldo: Aturan 'Balance Mismatch' mengidentifikasi {mismatch_count:,} transaksi anomali ({mismatch_pct:.2f}% dari total data), memperkuat adanya celah sistematis di mana saldo pengirim tidak terpotong dengan benar pada jenis transaksi TRANSFER & CASH_OUT."
    insight_2 = f"• Deteksi Sistem Tidak Efektif: Sistem deteksi bawaan PaySim melewatkan {missed_fraud:,} dari {total_fraud:,} total kasus fraud nyata (tingkat kegagalan {missed_pct:.1f}%), mengindikasikan bahwa sistem flagging bawaan tidak berfungsi secara memadai dan memerlukan revisi aturan darurat."
    insight_3 = f"• Jaringan Mule Account Terdeteksi: Terdapat {mule_count:,} rekening penerima mencurigakan (mule accounts) yang menerima beberapa transfer dana dari berbagai akun pengirim berbeda dalam rentang waktu singkat kurang dari 24 jam."

    # Write insights inside the card layout
    ws.merge_cells(start_row=9, start_column=2, end_row=10, end_column=7)
    c1 = ws.cell(row=9, column=2, value=insight_1)
    c1.font = font_data
    c1.alignment = align_wrap_left

    ws.merge_cells(start_row=11, start_column=2, end_row=12, end_column=7)
    c2 = ws.cell(row=11, column=2, value=insight_2)
    c2.font = font_data
    c2.alignment = align_wrap_left

    ws.merge_cells(start_row=13, start_column=2, end_row=14, end_column=7)
    c3 = ws.cell(row=13, column=2, value=insight_3)
    c3.font = font_data
    c3.alignment = align_wrap_left

    # ----------------------------------------------------
    # Helper Data Tables (Columns L to P - will be HIDDEN)
    # ----------------------------------------------------
    # Table 1: Fraud by Type (L8:M13)
    ws.cell(row=8, column=12, value="Tipe Transaksi")
    ws.cell(row=8, column=13, value="Jumlah Fraud")

    fraud_by_type = df.groupby('type')['isFraud'].sum().reset_index()
    for idx, row in fraud_by_type.iterrows():
        ws.cell(row=9+idx, column=12, value=row['type'])
        ws.cell(row=9+idx, column=13, value=int(row['isFraud']))

    # Table 2: Daily Transaction Volume (O8:P39)
    ws.cell(row=8, column=15, value="Hari")
    ws.cell(row=8, column=16, value="Jumlah Transaksi")

    df['day'] = (df['step'] - 1) // 24 + 1
    daily_vol = df.groupby('day')['amount'].count().reset_index()
    for idx, row in daily_vol.iterrows():
        ws.cell(row=9+idx, column=15, value=int(row['day']))
        ws.cell(row=9+idx, column=16, value=int(row['amount']))

    # ----------------------------------------------------
    # Charts Rendering
    # ----------------------------------------------------
    # Chart 1: Fraud Counts per Type Bar Chart
    bar = BarChart()
    bar.type = "col"
    bar.style = 13
    bar.title = "Kasus Fraud per Tipe Transaksi"
    bar.y_axis.title = "Kasus Fraud"
    bar.x_axis.title = "Tipe Transaksi"
    bar.legend = None
    bar.width = 11.5
    bar.height = 7.5

    data_bar = Reference(ws, min_col=13, min_row=8, max_row=8+len(fraud_by_type))
    cats_bar = Reference(ws, min_col=12, min_row=9, max_row=8+len(fraud_by_type))
    bar.add_data(data_bar, titles_from_data=True)
    bar.set_categories(cats_bar)
    ws.add_chart(bar, "B16")

    # Chart 2: Daily Volume Line Chart
    line = LineChart()
    line.style = 13
    line.title = "Volume Transaksi Harian (30 Hari)"
    line.y_axis.title = "Jumlah Transaksi"
    line.x_axis.title = "Hari ke-"
    line.legend = None
    line.width = 12.5
    line.height = 7.5

    data_line = Reference(ws, min_col=16, min_row=8, max_row=8+len(daily_vol))
    cats_line = Reference(ws, min_col=15, min_row=9, max_row=8+len(daily_vol))
    line.add_data(data_line, titles_from_data=True)
    line.set_categories(cats_line)
    ws.add_chart(line, "H16")

    # Hide reference columns (L to P) to maintain neat dashboard aesthetics
    for c_letter in ["L", "M", "N", "O", "P"]:
        ws.column_dimensions[c_letter].hidden = True

    # Set explicit column dimensions for dashboard layout spacing
    ws.column_dimensions["A"].width = 3
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
        ws.column_dimensions[col].width = 12

def main():
    # Setup Paths
    sample_file = "data/transactions_sample.csv"
    findings_file = "report/fraud_findings.csv"
    output_dir = "report"
    output_excel = os.path.join(output_dir, "Laporan_Audit_Fraud_PaySim.xlsx")

    # Ensure input files exist
    if not os.path.exists(sample_file):
        print(f"[ERROR] Sample file '{sample_file}' not found. Please run setup_dataset.py first.")
        sys.exit(1)
    if not os.path.exists(findings_file):
        print(f"[ERROR] Findings file '{findings_file}' not found. Please run fraud_detector.py first.")
        sys.exit(1)

    # Load Data
    print(f"Reading input CSV files...")
    try:
        df = pd.read_csv(sample_file)
        findings_df = pd.read_csv(findings_file)
        print(f"Loaded {len(df):,} sample records and {len(findings_df):,} findings records.")
    except Exception as e:
        print(f"[ERROR] Failed to read CSVs: {e}")
        sys.exit(1)

    # Initialize openpyxl workbook
    print("Initializing Excel Workbook...")
    wb = openpyxl.Workbook()

    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Create sheets (Dashboard first so it's the opening tab)
    create_sheet_dashboard_eksekutif(wb, df, findings_df)
    create_sheet_ringkasan_per_tipe(wb, df)
    create_sheet_temuan_fraud(wb, findings_df, len(df))
    create_sheet_data_transaksi(wb, df)

    # Ensure report folder exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory '{output_dir}/'")

    # Save Excel file
    print(f"Saving Excel report to '{output_excel}'...")
    try:
        wb.save(output_excel)
        print(f"[SUCCESS] Excel report generated successfully at: {output_excel}")
    except Exception as e:
        print(f"[ERROR] Failed to save Excel file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

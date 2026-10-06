#!/usr/bin/env python3
"""
fraud_detector.py

This script processes the transaction sample and applies 5 rule-based detection logic
to identify suspicious activities. It prints a summary of the findings, outputs
the detailed findings to 'report/fraud_findings.csv', and creates an audit summary
in 'report/summary.txt'.
"""

import os
import sys
import time
import pandas as pd

def detect_balance_mismatch(df):
    """
    Rule 1: Balance Mismatch
    Logika: oldbalanceOrg - amount != newbalanceOrig (selisih > 1 karena float)
    Hanya cek tipe TRANSFER dan CASH_OUT.

    Pola fraud klasik di PaySim di mana saldo pengirim tidak berkurang secara presisi
    atau tersisa saldo yang tidak sesuai setelah dana ditransfer keluar.
    """
    # Filter only TRANSFER and CASH_OUT transaction types
    target_types = ['TRANSFER', 'CASH_OUT']
    df_filtered = df[df['type'].isin(target_types)].copy()

    if df_filtered.empty:
        return pd.DataFrame(columns=df.columns)

    # Check mismatch where the difference is greater than 1.0 (to handle float precision issues)
    expected_new_balance = df_filtered['oldbalanceOrg'] - df_filtered['amount']
    mismatch_mask = (expected_new_balance - df_filtered['newbalanceOrig']).abs() > 1.0

    return df_filtered[mismatch_mask].copy()

def detect_large_transactions(df):
    """
    Rule 2: Large Transactions (Outliers)
    Logika: amount > mean + 3*std per tipe transaksi

    Menandai transaksi yang nilainya sangat menyimpang (outliers) dibandingkan
    dengan rata-rata transaksi lain dengan tipe yang sama.
    """
    if df.empty:
        return pd.DataFrame(columns=df.columns)

    # Calculate mean and standard deviation per transaction type
    stats = df.groupby('type')['amount'].agg(['mean', 'std']).reset_index()
    stats['std'] = stats['std'].fillna(0.0) # Fill NaN standard deviation (for single-item groups)

    # Merge statistics back into the dataframe
    merged = df.merge(stats, on='type', how='left')

    # Outlier threshold condition
    threshold = merged['mean'] + 3 * merged['std']
    outlier_mask = merged['amount'] > threshold

    return df[outlier_mask].copy()

def detect_zero_balance_origin(df):
    """
    Rule 3: Zero Balance Origin (Drained Account)
    Logika: newbalanceOrig == 0 AND oldbalanceOrg > 0 AND isFraud == 0

    Akun pengirim yang saldonya terkuras habis sampai tepat nol, tetapi transaksi
    tersebut tidak diberi label fraud. Ini sering menjadi indikator fraud yang terlewat.
    """
    if df.empty:
        return pd.DataFrame(columns=df.columns)

    mask = (df['newbalanceOrig'] == 0) & (df['oldbalanceOrg'] > 0) & (df['isFraud'] == 0)
    return df[mask].copy()

def detect_suspicious_dest_accounts(df):
    """
    Rule 4: Suspicious Destination Accounts (Mule Accounts)
    Logika: nameDest yang menerima lebih dari 3 transaksi berbeda dalam window 24 step (24 jam)

    Akun penerima yang menerima banyak kiriman dana dari berbagai pihak dalam rentang
    waktu singkat (24 jam), mengindikasikan pola "mule account" penampung uang.
    """
    if len(df) < 4:
        return pd.DataFrame(columns=df.columns)

    # Sort by destination and time step for sequential analysis
    sorted_df = df.sort_values(['nameDest', 'step']).copy()

    # Shift step by 3 rows within the same nameDest group (covers 4 consecutive transactions)
    sorted_df['prev_step_3'] = sorted_df.groupby('nameDest')['step'].shift(3)

    # Check if the time difference is 24 hours or less (<= 24 steps)
    sorted_df['is_suspicious'] = (sorted_df['step'] - sorted_df['prev_step_3']) <= 24

    # Extract unique destination accounts that triggered this rule
    flagged_dests = sorted_df.loc[sorted_df['is_suspicious'] == True, 'nameDest'].unique()

    # Return all transactions related to these flagged destination accounts
    return df[df['nameDest'].isin(flagged_dests)].copy()

def detect_missed_by_system(df):
    """
    Rule 5: Missed by System
    Logika: isFraud == 1 AND isFlaggedFraud == 0

    Kasus fraud nyata yang berhasil lolos dari sistem flagging bawaan PaySim.
    Sangat penting untuk dianalisis oleh auditor guna memperbaiki sistem deteksi.
    """
    if df.empty:
        return pd.DataFrame(columns=df.columns)

    mask = (df['isFraud'] == 1) & (df['isFlaggedFraud'] == 0)
    return df[mask].copy()

def main():
    # Setup paths
    input_file = "data/transactions_sample.csv"
    output_dir = "report"
    findings_file = os.path.join(output_dir, "fraud_findings.csv")
    summary_file = os.path.join(output_dir, "summary.txt")

    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"[ERROR] Input file '{input_file}' not found.")
        print("Please run 'setup_dataset.py' first to prepare the dataset.")
        sys.exit(1)

    # Load dataset
    print(f"Reading transaction data from '{input_file}'...")
    try:
        df = pd.read_csv(input_file)
        print(f"Loaded {len(df):,} transactions successfully.")
    except Exception as e:
        print(f"[ERROR] Failed to read CSV file: {e}")
        sys.exit(1)

    # Define the rules to run
    rules = [
        ("Balance Mismatch", detect_balance_mismatch, "balance_mismatch",
         "oldbalanceOrg - amount != newbalanceOrig (TRANSFER & CASH_OUT only)"),
        ("Large Transactions", detect_large_transactions, "large_transaction",
         "amount > mean + 3*std per transaction type"),
        ("Zero Balance Origin", detect_zero_balance_origin, "zero_balance_origin",
         "newbalanceOrig == 0 AND oldbalanceOrg > 0 AND isFraud == 0"),
        ("Suspicious Destination Accounts", detect_suspicious_dest_accounts, "suspicious_dest_account",
         "nameDest receiving > 3 transactions in 24 hours"),
        ("Missed by System", detect_missed_by_system, "missed_by_system",
         "isFraud == 1 AND isFlaggedFraud == 0")
    ]

    findings_dfs = []
    rule_summaries = []

    print("\n" + "=" * 80)
    print("RUNNING AUDIT RULE-BASED DETECTIONS")
    print("=" * 80)

    for rule_name, rule_func, rule_id, rule_desc in rules:
        print(f"\nRunning: {rule_name}...")
        print(f"Logic  : {rule_desc}")

        start_time = time.time()
        findings_df = rule_func(df)
        execution_time = time.time() - start_time

        num_findings = len(findings_df)
        pct_findings = (num_findings / len(df)) * 100

        print(f"Time   : {execution_time:.6f} seconds")
        print(f"Found  : {num_findings:,} suspicious transactions ({pct_findings:.2f}%)")

        # Display top 3 examples if found
        if num_findings > 0:
            print("Top 3 Examples:")
            preview_cols = ['step', 'type', 'amount', 'nameOrig', 'oldbalanceOrg', 'newbalanceOrig', 'nameDest', 'isFraud']
            # Select preview columns that actually exist in the dataframe
            actual_preview_cols = [c for c in preview_cols if c in findings_df.columns]
            print(findings_df[actual_preview_cols].head(3).to_string(index=False))

            # Add rule identifier column
            findings_tagged = findings_df.copy()
            findings_tagged['rule_triggered'] = rule_id
            findings_dfs.append(findings_tagged)
        else:
            print("No anomalies detected for this rule.")

        rule_summaries.append({
            'name': rule_name,
            'desc': rule_desc,
            'count': num_findings,
            'percentage': pct_findings,
            'time_sec': execution_time
        })
        print("-" * 80)

    # Combine all findings into a single DataFrame
    if findings_dfs:
        all_findings = pd.concat(findings_dfs, ignore_index=True)
        # Calculate unique transaction overlap
        # Since a transaction could be flagged by multiple rules, it might appear multiple times
        # We can identify duplicate rows (excluding the 'rule_triggered' column)
        duplicate_cols = [c for c in all_findings.columns if c != 'rule_triggered']
        unique_flagged_count = all_findings.duplicated(subset=duplicate_cols).sum()
        total_unique_flagged = len(all_findings.drop_duplicates(subset=duplicate_cols))
    else:
        all_findings = pd.DataFrame(columns=list(df.columns) + ['rule_triggered'])
        total_unique_flagged = 0
        unique_flagged_count = 0

    # Create output directory
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"\nCreated output directory '{output_dir}/'")

    # Save detailed findings to CSV
    try:
        all_findings.to_csv(findings_file, index=False)
        print(f"[SUCCESS] Detailed findings saved to '{findings_file}' ({len(all_findings):,} total triggers).")
    except Exception as e:
        print(f"[ERROR] Failed to save findings to CSV: {e}")

    # Prepare summary report text
    summary_text = []
    summary_text.append("=" * 80)
    summary_text.append("AUDIT DETECTION SUMMARY REPORT")
    summary_text.append("=" * 80)
    summary_text.append(f"Source File             : {input_file}")
    summary_text.append(f"Total Transactions      : {len(df):,}")
    summary_text.append(f"Total Unique Flagged    : {total_unique_flagged:,} ({ (total_unique_flagged / len(df)) * 100:.2f}%)")
    summary_text.append(f"Multi-Rule Overlaps     : {unique_flagged_count:,}")
    summary_text.append("-" * 80)
    summary_text.append(f"{'Rule Name':<35} | {'Count':<8} | {'% of Data':<10} | {'Time (sec)':<12}")
    summary_text.append("-" * 80)
    for summary in rule_summaries:
        summary_text.append(
            f"{summary['name']:<35} | {summary['count']:<8,} | {summary['percentage']:<9.2f}% | {summary['time_sec']:<12.6f}"
        )
    summary_text.append("-" * 80)
    summary_text.append("\nRule Explanations:")
    for summary in rule_summaries:
        summary_text.append(f"- {summary['name']}: {summary['desc']}")
    summary_text.append("=" * 80)

    summary_string = "\n".join(summary_text)

    # Save summary report to txt
    try:
        with open(summary_file, 'w') as f:
            f.write(summary_string)
        print(f"[SUCCESS] Audit summary report saved to '{summary_file}'.")
    except Exception as e:
        print(f"[ERROR] Failed to save summary report: {e}")

    print("\n" + "=" * 80)
    print("DETECTION SUMMARY STATUS")
    print("=" * 80)
    print(summary_string)
    print("=" * 80)

if __name__ == "__main__":
    main()

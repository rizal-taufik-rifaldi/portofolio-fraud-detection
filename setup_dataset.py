#!/usr/bin/env python3
"""
setup_dataset.py

This script downloads the PaySim Synthetic Financial Dataset for Fraud Detection
from Hugging Face, takes a stratified sample of 50,000 rows to keep it lightweight,
saves it to the 'data' directory, and prints basic descriptive statistics.
"""

import os
import sys
import subprocess

def install_and_import(package_name, import_name=None):
    """Checks if a package is installed. If not, installs it using pip."""
    if import_name is None:
        import_name = package_name
    try:
        __import__(import_name)
        print(f"[INFO] Package '{package_name}' is already installed.")
    except ImportError:
        print(f"[INFO] Package '{package_name}' not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"[SUCCESS] Package '{package_name}' installed successfully.")
        except Exception as e:
            print(f"[ERROR] Failed to install package '{package_name}': {e}")
            print("[INFO] Please install it manually using: pip install " + package_name)
            sys.exit(1)

# Step 1: Ensure dependencies are installed
print("=" * 60)
print("Step 1: Checking and Installing Dependencies")
print("=" * 60)
install_and_import("pandas")
install_and_import("datasets")
install_and_import("openpyxl")
print("-" * 60)

# Import the libraries after confirming they are installed
import pandas as pd
from datasets import load_dataset

# Step 2: Load the dataset from Hugging Face
print("\n" + "=" * 60)
print("Step 2: Loading Dataset from Hugging Face")
print("=" * 60)
dataset_name = "purulalwani/Synthetic-Financial-Datasets-For-Fraud-Detection"
print(f"Loading '{dataset_name}'...")

try:
    ds = load_dataset(dataset_name)
    print("[SUCCESS] Dataset loaded successfully from Hugging Face.")
except Exception as e:
    print(f"[ERROR] Failed to load dataset from Hugging Face: {e}")
    sys.exit(1)

# Step 3: Convert to pandas DataFrame
print("\n" + "=" * 60)
print("Step 3: Converting to pandas DataFrame")
print("=" * 60)

try:
    if hasattr(ds, "keys"):
        print(f"Dataset splits found: {list(ds.keys())}")
        # Concatenate all splits if there are multiple, or just take the main one
        dfs = []
        for split in ds.keys():
            print(f"Converting split '{split}' to DataFrame...")
            dfs.append(ds[split].to_pandas())
        df = pd.concat(dfs, ignore_index=True)
    else:
        df = ds.to_pandas()

    print(f"[SUCCESS] Converted to DataFrame. Shape: {df.shape}")
except Exception as e:
    print(f"[ERROR] Failed to convert dataset to pandas DataFrame: {e}")
    sys.exit(1)

# Validate columns
required_columns = ['type', 'isFraud', 'amount']
missing_cols = [col for col in required_columns if col not in df.columns]
if missing_cols:
    print(f"[WARNING] Missing expected columns for stratification/statistics: {missing_cols}")

# Step 4: Stratified Sampling
print("\n" + "=" * 60)
print("Step 4: Performing Stratified Sampling")
print("=" * 60)
target_samples = 50000

if 'type' in df.columns:
    try:
        print(f"Creating a stratified sample of {target_samples} rows based on 'type'...")

        # Calculate class proportions
        proportions = df['type'].value_counts(normalize=True)
        print("Original 'type' distribution proportion:")
        for k, v in proportions.items():
            print(f"  - {k}: {v:.4f}")

        # Determine sample sizes per class
        samples_per_class = (proportions * target_samples).round().astype(int)

        # Handle rounding errors to ensure total is exactly target_samples
        diff = target_samples - samples_per_class.sum()
        if diff != 0:
            largest_class = proportions.idxmax()
            samples_per_class[largest_class] += diff

        # Sample from each class
        sampled_dfs = []
        for category, count in samples_per_class.items():
            sub_df = df[df['type'] == category]
            actual_count = min(len(sub_df), count)
            if actual_count > 0:
                sampled_dfs.append(sub_df.sample(n=actual_count, random_state=42))
            else:
                print(f"[WARNING] Category '{category}' has no elements to sample.")

        df_sample = pd.concat(sampled_dfs).sample(frac=1, random_state=42).reset_index(drop=True)
        print(f"[SUCCESS] Stratified sampling completed. Sample shape: {df_sample.shape}")
    except Exception as e:
        print(f"[WARNING] Stratified sampling failed: {e}. Falling back to random sampling.")
        df_sample = df.sample(n=min(len(df), target_samples), random_state=42).reset_index(drop=True)
else:
    print("[INFO] 'type' column not found. Performing random sampling instead.")
    df_sample = df.sample(n=min(len(df), target_samples), random_state=42).reset_index(drop=True)

# Step 5: Save to data/transactions_sample.csv
print("\n" + "=" * 60)
print("Step 5: Saving Sample to CSV")
print("=" * 60)
output_dir = "data"
output_file = os.path.join(output_dir, "transactions_sample.csv")

try:
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory '{output_dir}/'")

    print(f"Saving to '{output_file}'...")
    df_sample.to_csv(output_file, index=False)
    print(f"[SUCCESS] Sample saved to '{output_file}'.")
except Exception as e:
    print(f"[ERROR] Failed to save sample: {e}")
    sys.exit(1)

# Step 6: Print initial statistics
print("\n" + "=" * 60)
print("Step 6: Sample Dataset Statistics")
print("=" * 60)

print(f"Total Rows: {len(df_sample)}")

if 'type' in df_sample.columns:
    print("\nTransaction Type Distribution:")
    type_counts = df_sample['type'].value_counts()
    type_pcts = df_sample['type'].value_counts(normalize=True) * 100
    for name in type_counts.index:
        print(f"  - {name:<12}: {type_counts[name]:>6} ({type_pcts[name]:.2f}%)")

if 'isFraud' in df_sample.columns:
    print("\nFraud vs Non-Fraud Distribution:")
    fraud_counts = df_sample['isFraud'].value_counts()
    fraud_pcts = df_sample['isFraud'].value_counts(normalize=True) * 100
    for name in fraud_counts.index:
        label = "Fraud" if name == 1 else "Non-Fraud"
        print(f"  - {label:<10}: {fraud_counts[name]:>6} ({fraud_pcts[name]:.2f}%)")

if 'amount' in df_sample.columns:
    total_amount = df_sample['amount'].sum()
    mean_amount = df_sample['amount'].mean()
    max_amount = df_sample['amount'].max()
    print(f"\nTransaction Nominal Statistics:")
    print(f"  - Total Amount : {total_amount:,.2f}")
    print(f"  - Mean Amount  : {mean_amount:,.2f}")
    print(f"  - Max Amount   : {max_amount:,.2f}")

print("\n" + "=" * 60)
print("Script execution completed successfully!")
print("=" * 60)

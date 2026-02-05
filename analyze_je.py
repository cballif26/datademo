import pandas as pd
import os
from datetime import datetime

def analyze():
    print("Starting analysis...")
    input_file = 'je_samples.xlsx'
    output_dir = 'output'
    output_file = os.path.join(output_dir, 'report.md')

    # Ensure output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Read the file
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading {input_file}: {e}")
        return

    # Clean string columns to remove trailing whitespace
    # Some object columns might contain mixed types (int and str), so we only strip actual strings
    for col in df.select_dtypes(['object']).columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

    # Calculate statistics
    row_count = len(df)
    min_date = df['EffectiveDate'].min()
    max_date = df['EffectiveDate'].max()
    total_amount = df['Amount'].sum()
    unique_accounts = df['GLAccountNumber'].nunique()

    # Categorical counts
    source_counts = df['Source'].value_counts()
    bu_counts = df['BusinessUnit'].value_counts()
    account_type_counts = df['AccountType'].value_counts()

    # Generate Report
    with open(output_file, 'w') as f:
        f.write("# GL Data Analysis Report\n\n")
        f.write(f"**Generated on:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        f.write("## Summary Statistics\n")
        f.write("| Metric | Value |\n")
        f.write("|---|---|\n")
        f.write(f"| Total Rows | {row_count:,} |\n")
        try:
             f.write(f"| Date Range | {min_date.date()} to {max_date.date()} |\n")
        except AttributeError:
             f.write(f"| Date Range | {min_date} to {max_date} |\n")
        f.write(f"| Total Amount | {total_amount:,.2f} |\n")
        f.write(f"| Unique GL Accounts | {unique_accounts:,} |\n\n")

        f.write("## Entries by Source\n")
        f.write("| Source | Count |\n")
        f.write("|---|---|\n")
        for source, count in source_counts.items():
            f.write(f"| {source} | {count:,} |\n")
        f.write("\n")

        f.write("## Entries by Business Unit\n")
        f.write("| Business Unit | Count |\n")
        f.write("|---|---|\n")
        for bu, count in bu_counts.items():
            f.write(f"| {bu} | {count:,} |\n")
        f.write("\n")

        f.write("## Entries by Account Type\n")
        f.write("| Account Type | Count |\n")
        f.write("|---|---|\n")
        for at, count in account_type_counts.items():
            f.write(f"| {at} | {count:,} |\n")
        f.write("\n")

    print(f"Analysis complete. Report written to {output_file}")

if __name__ == "__main__":
    analyze()

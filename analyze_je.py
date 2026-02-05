import pandas as pd
import os
import math
import matplotlib.pyplot as plt
from datetime import datetime

def get_leading_digit(n):
    """Returns the leading digit of a number."""
    try:
        s = str(abs(n))
        for char in s:
            if char.isdigit() and char != '0':
                return int(char)
        return None
    except (ValueError, TypeError):
        return None

def process_file(filepath, report_file, output_dir):
    filename = os.path.basename(filepath)
    print(f"Processing {filename}...")

    # Read the file
    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        report_file.write(f"## Error processing {filename}\n\n")
        report_file.write(f"Error: {e}\n\n")
        return

    # Clean string columns
    for col in df.select_dtypes(['object']).columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

    # Calculate statistics
    row_count = len(df)
    try:
        min_date = df['EffectiveDate'].min()
        max_date = df['EffectiveDate'].max()
    except KeyError:
        min_date = "N/A"
        max_date = "N/A"

    try:
        total_amount = df['Amount'].sum()
    except KeyError:
        total_amount = 0

    try:
        unique_accounts = df['GLAccountNumber'].nunique()
    except KeyError:
        unique_accounts = 0

    # Write file header to report
    report_file.write(f"# Analysis for {filename}\n\n")

    # Summary Statistics
    report_file.write("## Summary Statistics\n")
    report_file.write("| Metric | Value |\n")
    report_file.write("|---|---|\n")
    report_file.write(f"| Total Rows | {row_count:,} |\n")

    date_range_str = "N/A"
    if hasattr(min_date, 'date'):
        date_range_str = f"{min_date.date()} to {max_date.date()}"
    elif min_date != "N/A":
        date_range_str = f"{min_date} to {max_date}"

    report_file.write(f"| Date Range | {date_range_str} |\n")
    report_file.write(f"| Total Amount | {total_amount:,.2f} |\n")
    report_file.write(f"| Unique GL Accounts | {unique_accounts:,} |\n\n")

    # Categorical counts
    if 'Source' in df.columns:
        source_counts = df['Source'].value_counts()
        report_file.write("## Entries by Source\n")
        report_file.write("| Source | Count |\n")
        report_file.write("|---|---|\n")
        for source, count in source_counts.items():
            report_file.write(f"| {source} | {count:,} |\n")
        report_file.write("\n")

    if 'BusinessUnit' in df.columns:
        bu_counts = df['BusinessUnit'].value_counts()
        report_file.write("## Entries by Business Unit\n")
        report_file.write("| Business Unit | Count |\n")
        report_file.write("|---|---|\n")
        for bu, count in bu_counts.items():
            report_file.write(f"| {bu} | {count:,} |\n")
        report_file.write("\n")

    if 'AccountType' in df.columns:
        account_type_counts = df['AccountType'].value_counts()
        report_file.write("## Entries by Account Type\n")
        report_file.write("| Account Type | Count |\n")
        report_file.write("|---|---|\n")
        for at, count in account_type_counts.items():
            report_file.write(f"| {at} | {count:,} |\n")
        report_file.write("\n")

    # Benford's Law Analysis
    if 'Amount' in df.columns:
        report_file.write("## Benford's Law Analysis\n\n")

        # Filter non-zero amounts
        amounts = df['Amount'][df['Amount'] != 0].dropna()

        # Get leading digits
        leading_digits = amounts.apply(lambda x: get_leading_digit(x)).dropna()

        if leading_digits.empty:
             report_file.write("Insufficient data for Benford's Law analysis.\n\n")
        else:
            # Calculate observed frequencies
            observed_counts = leading_digits.value_counts().sort_index()
            total_digits = len(leading_digits)

            # Benford's expected frequencies
            digits = range(1, 10)
            expected_probs = [math.log10(1 + 1/d) for d in digits]
            expected_counts = [p * total_digits for p in expected_probs]

            # Prepare data for table and plot
            observed_freqs = []
            for d in digits:
                count = observed_counts.get(d, 0)
                observed_freqs.append(count / total_digits)

            # Table
            report_file.write("| Digit | Observed Count | Observed % | Expected % |\n")
            report_file.write("|---|---|---|---|\n")
            for i, d in enumerate(digits):
                 report_file.write(f"| {d} | {observed_counts.get(d, 0):,} | {observed_freqs[i]*100:.2f}% | {expected_probs[i]*100:.2f}% |\n")
            report_file.write("\n")

            # Plot
            plt.figure(figsize=(10, 6))
            width = 0.35
            x = range(1, 10)

            plt.bar([p - width/2 for p in x], [f * 100 for f in observed_freqs], width, label='Observed')
            plt.bar([p + width/2 for p in x], [p * 100 for p in expected_probs], width, label='Expected (Benford)', alpha=0.7)

            plt.xlabel('Leading Digit')
            plt.ylabel('Frequency (%)')
            plt.title(f"Benford's Law Analysis - {filename}")
            plt.xticks(x)
            plt.legend()
            plt.grid(axis='y', linestyle='--', alpha=0.7)

            # Save plot
            plot_filename = f"benford_{os.path.splitext(filename)[0]}.png"
            plot_path = os.path.join(output_dir, plot_filename)
            plt.savefig(plot_path)
            plt.close()

            report_file.write(f"![Benford Analysis]({plot_filename})\n\n")

    report_file.write("---\n\n")

def main():
    print("Starting analysis...")
    output_dir = 'output'
    report_path = os.path.join(output_dir, 'report.md')

    # Ensure output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Open report file
    with open(report_path, 'w') as f:
        f.write("# GL Data Analysis Report\n\n")
        f.write(f"**Generated on:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        # Find all xlsx files
        files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')]

        if not files:
            f.write("No Excel files found for analysis.\n")
            print("No Excel files found.")
            return

        for filename in files:
             process_file(filename, f, output_dir)

    print(f"Analysis complete. Report written to {report_path}")

if __name__ == "__main__":
    main()

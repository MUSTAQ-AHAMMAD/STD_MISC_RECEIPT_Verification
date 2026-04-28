# Payment Verification Automation

Comprehensive automation script for payment reconciliation and verification across all store folders.

## Features

This script automates the complete verification process with the following checks:

1. **Amount Verification**: Compares amounts between payment files and standard consolidated receipts
2. **Transaction Number Verification**: Matches Oracle transaction numbers to standard receipt IDs (handles prefixes: Visa-, Cash-, Mada-, Master-, AMEX-)
3. **Organization ID Validation**: Verifies that all records have the correct Organization ID (`AlQurashi-KSA`)
4. **Bank Account Verification**: Validates bank account details against the reference file (`Subledger Bank Account.xlsx`)
5. **Comprehensive Reporting**: Generates detailed reports for each folder with all verification results

## Prerequisites

- Python 3.7 or higher
- Required Python packages (automatically installed):
  - pandas
  - openpyxl

## File Structure

```
Repository Root/
├── RUN_VERIFICATION.bat         # ⭐ MAIN SCRIPT - Double-click this on Windows!
├── master_automation.py         # Master script (extraction + verification)
├── extract_zip_files.py         # Extracts zip files (for TABOUK onwards)
├── payment_verification.py      # Main verification script
├── run_verification.bat         # Alternative batch script (verification only)
├── requirements.txt             # Python dependencies
├── VERIFICATION_README.md       # This file
├── Subledger Bank Account.xlsx  # Bank account reference file
├── STORE_FOLDER_1/
│   ├── STORE_FOLDER_1.csv      # Oracle payment file
│   ├── Receipt_ALL_CONSOLIDATED.csv  # Standard receipts
│   ├── oracle_fusion_output (XX).zip  # Zip file (if not extracted yet)
│   ├── ORACLE_FUSION_OUTPUT/   # Extracted folder
│   │   └── Receipts/
│   │       └── Receipt_ALL_CONSOLIDATED.csv  # Oracle receipts
│   └── [Generated Reports]
├── STORE_FOLDER_2/
│   └── ...
└── ...
```

## Usage

### On Windows (Recommended - Easy One-Click Solution)

Simply double-click `RUN_VERIFICATION.bat` in Windows Explorer.

The script will automatically:
1. Check if Python is installed
2. Install required dependencies if needed
3. Extract any remaining zip files (from TABOUK onwards)
4. Run the comprehensive verification process
5. Generate detailed reports for all folders

### Manual Execution

#### Complete Automation (Recommended)
```bash
# Install dependencies
pip install -r requirements.txt

# Run the master automation (extraction + verification)
python master_automation.py
```

#### Verification Only (if zip files already extracted)
```bash
# Install dependencies
pip install -r requirements.txt

# Run only the verification
python payment_verification.py
```

### On Linux/Mac

```bash
# Make script executable
chmod +x payment_verification.py

# Install dependencies
pip install -r requirements.txt

# Run verification
./payment_verification.py
```

## Output

For each store folder, the script generates:

1. **Detailed Text Report** (`Verification_Report_YYYYMMDD_HHMMSS.txt`)
   - Complete verification results
   - Summary statistics
   - Detailed lists of mismatches and errors
   - Easy to read and share

2. **JSON Summary** (`Verification_Summary_YYYYMMDD_HHMMSS.json`)
   - Machine-readable format
   - Contains all verification data
   - Can be used for further analysis

3. **Overall Summary** (`Overall_Verification_Summary_YYYYMMDD_HHMMSS.json`)
   - Summary across all folders
   - Generated in the root directory
   - Shows overall statistics

## Verification Details

### 1. Amount Verification
- Compares transaction amounts between Oracle data and standard receipts
- Identifies matched, mismatched, and missing transactions
- Calculates total amounts and differences
- Allows for rounding differences (< 0.01)

### 2. Transaction Number Matching
- Strips payment prefixes (Visa-, Cash-, Mada-, Master-, AMEX-) for comparison
- Matches Oracle transaction numbers to receipt numbers
- Identifies missing transactions in either file
- Example: "Cash-BLKU-0001687" matches with "BLKU-0001687"

### 3. Organization ID Validation
- Validates that all records have Organization ID: `AlQurashi-KSA`
- Reports any invalid Organization IDs found
- Checks both Oracle receipts and Misc receipts

### 4. Bank Account Verification
- Matches bank account names in receipts against reference file
- Identifies accounts not found in reference
- Shows which receipts use unrecognized accounts

## Configuration

The script uses the following constants (can be modified in `payment_verification.py`):

```python
ORGANIZATION_ID = "AlQurashi-KSA"
PAYMENT_PREFIXES = ["Visa-", "Cash-", "Mada-", "Master-", "AMEX-"]
BANK_ACCOUNT_FILE = "Subledger Bank Account.xlsx"
```

## Troubleshooting

### Python not found
- Install Python from https://www.python.org/
- Make sure to check "Add Python to PATH" during installation

### Permission denied
- On Linux/Mac, make the script executable: `chmod +x payment_verification.py`
- On Windows, run Command Prompt as Administrator

### Missing files
- Ensure all store folders have the required CSV files
- Check that `Subledger Bank Account.xlsx` exists in the root directory

### Excel file errors
- Make sure Excel files are closed before running the script
- Verify that Excel files are not corrupted

## Report Interpretation

### Success Indicators (Green ✓)
- Matched transactions
- Valid Organization IDs
- Matched bank accounts

### Warnings (Yellow ⚠)
- Missing transactions (might be expected in some cases)
- Bank accounts not in reference file (might need to update reference)

### Errors (Red ✗)
- Mismatched amounts (requires investigation)
- Invalid Organization IDs (needs correction)
- Failed file loads (check file paths)

## Notes

- The script processes all folders with uppercase names that contain required files
- Processing time depends on the number of folders and transactions
- Reports are timestamped to avoid overwriting previous runs
- The script is safe to run multiple times

## Support

For issues or questions:
1. Check the error messages in the console output
2. Review the generated reports for details
3. Ensure all input files are properly formatted CSV files
4. Verify that the bank account reference file is up to date

## Examples

### Successful Run Output
```
================================================================================
Processing Folder: TABOUK
================================================================================

ℹ Loading Oracle file: TABOUK.csv
ℹ Loading standard receipts: Receipt_ALL_CONSOLIDATED.csv
ℹ Loading Oracle receipts: Receipt_ALL_CONSOLIDATED.csv

--------------------------------------------------------------------------------
1. Amount Verification
--------------------------------------------------------------------------------
ℹ Total Oracle Amount: 150000.00
ℹ Total Standard Amount: 150000.00
ℹ Difference: 0.00
✓ Matched transactions: 245

--------------------------------------------------------------------------------
2. Transaction Number Verification
--------------------------------------------------------------------------------
✓ Matched transactions: 245

--------------------------------------------------------------------------------
3. Organization ID Verification
--------------------------------------------------------------------------------
ℹ Expected Organization ID: AlQurashi-KSA
ℹ Found Organization IDs: AlQurashi-KSA
✓ Valid records: 245

--------------------------------------------------------------------------------
4. Bank Account Details Verification
--------------------------------------------------------------------------------
ℹ Unique bank accounts found: 5
✓ Matched accounts: 245

✓ Report saved: Verification_Report_20260428_010000.txt
✓ Summary saved: Verification_Summary_20260428_010000.json
```

## License

Internal use only for payment verification automation.

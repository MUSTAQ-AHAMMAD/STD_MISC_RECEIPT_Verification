# QUICK START GUIDE

## For Windows Users (Easiest Method)

### Step 1: Install Python (One-time setup)
If you don't have Python installed:
1. Go to https://www.python.org/downloads/
2. Download Python 3.7 or higher
3. Run the installer
4. ⚠️ **IMPORTANT**: Check the box "Add Python to PATH" during installation

### Step 2: Run the Verification
1. Open Windows Explorer
2. Navigate to the repository folder
3. Double-click `RUN_VERIFICATION.bat`
4. Wait for the process to complete
5. Check the generated reports in each store folder

That's it! The script will:
- ✓ Check Python installation
- ✓ Install required packages automatically
- ✓ Extract zip files (TABOUK onwards)
- ✓ Run all verifications
- ✓ Generate detailed reports

## What Gets Verified?

### 1. Amount Verification ✓
- Compares amounts from Oracle payment files vs standard consolidated receipts
- Shows matched, mismatched, and missing transactions
- Reports: Total amounts, differences, and mismatches

### 2. Transaction Number Matching ✓
- Matches transaction numbers with automatic prefix removal
- Handles: Visa-, Cash-, Mada-, Master-, AMEX-
- Example: "Cash-BLKU-0001687" ↔ "BLKU-0001687"

### 3. Organization ID Validation ✓
- Verifies all records have: `AlQurashi-KSA`
- Reports any invalid Organization IDs

### 4. Bank Account Verification ✓
- Validates bank accounts against `Subledger Bank Account.xlsx`
- Shows which accounts are not in the reference file

## Where Are the Reports?

After running the script, you'll find:

### In Each Store Folder:
- `Verification_Report_YYYYMMDD_HHMMSS.txt` - Detailed human-readable report
- `Verification_Summary_YYYYMMDD_HHMMSS.json` - Machine-readable summary

### In the Root Folder:
- `Overall_Verification_Summary_YYYYMMDD_HHMMSS.json` - Summary of all folders

## Understanding the Output

### Console Colors:
- 🟢 **Green (✓)**: Success - Everything matched
- 🟡 **Yellow (⚠)**: Warning - Needs attention (missing items, accounts not in reference)
- 🔴 **Red (✗)**: Error - Critical issue (mismatched amounts, invalid org IDs)
- 🔵 **Blue (ℹ)**: Information - Status updates

### Report Sections:

#### SUMMARY Section
- Total amounts from Oracle and Standard receipts
- Number of matched/mismatched transactions
- Organization ID validation counts
- Bank account verification counts

#### MISMATCHED AMOUNTS Section
- Lists all transactions where amounts don't match
- Shows Oracle amount, Standard amount, and difference

#### MISSING IN STANDARD RECEIPTS Section
- Transactions in Oracle but not in standard receipts
- Shows transaction IDs and amounts

#### INVALID ORGANIZATION IDs Section
- Records with incorrect Organization ID
- Shows what was found instead of expected `AlQurashi-KSA`

#### BANK ACCOUNTS NOT FOUND Section
- Bank accounts used but not in reference file
- Shows how many receipts use each unrecognized account

## Troubleshooting

### "Python is not installed"
- Install Python from https://www.python.org/
- Make sure to check "Add Python to PATH"

### "Failed to load bank account reference file"
- Make sure `Subledger Bank Account.xlsx` exists in the root folder
- Check that the file is not open in Excel
- Verification will continue without bank account checking

### "Oracle receipt file not found"
- Some folders might not have extracted their zip files yet
- The script automatically extracts zip files for TABOUK onwards
- Check if the folder has `oracle_fusion_output (XX).zip` file

### Permission Errors
- Close any Excel files in the folders
- Run Command Prompt as Administrator (if needed)

## Advanced Usage

### Run Extraction Only
```bash
python extract_zip_files.py
```

### Run Verification Only (no extraction)
```bash
python payment_verification.py
```

### Run Complete Automation (extraction + verification)
```bash
python master_automation.py
```

## What to Do With Results?

1. **Review Reports**: Check the text reports for each folder
2. **Focus on Red Errors**: Mismatched amounts need investigation
3. **Check Yellow Warnings**: Missing transactions might be expected
4. **Update Bank Reference**: If many "not found" accounts, update reference file
5. **Share Reports**: Text reports are easy to share with team

## Tips

- Run the script regularly after new data arrives
- Keep the bank account reference file updated
- Reports are timestamped, so you can track changes over time
- The script is safe to run multiple times

## Need Help?

Check the detailed documentation in `VERIFICATION_README.md`

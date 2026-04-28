#!/usr/bin/env python3
"""
Comprehensive Payment Verification Automation Script

This script automates the verification process for payment reconciliation:
1. Amount verification from payment file and standard consolidated receipt
2. Transaction number verification from Oracle file to standard receipt ID
3. Organization ID validation in Misc script
4. Bank account details matching with standard file
5. Generates full report for each folder with all verifications
"""

import os
import sys
import csv
import json
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Set
import openpyxl
import re


# Base directory
BASE_DIR = Path(__file__).parent

# Bank account reference file
BANK_ACCOUNT_FILE = BASE_DIR / "Subledger Bank Account.xlsx"

# Organization ID constant
ORGANIZATION_ID = "AlQurashi-KSA"

# Payment prefixes to remove for transaction matching
PAYMENT_PREFIXES = ["Visa-", "Cash-", "Mada-", "Master-", "AMEX-"]


class Color:
    """ANSI color codes for terminal output"""
    RESET = '\033[0m'
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'


def print_header(text: str):
    """Print a formatted header"""
    print(f"\n{Color.BOLD}{Color.CYAN}{'=' * 80}")
    print(f"{text}")
    print(f"{'=' * 80}{Color.RESET}\n")


def print_section(text: str):
    """Print a formatted section"""
    print(f"\n{Color.BOLD}{Color.BLUE}{'-' * 80}")
    print(f"{text}")
    print(f"{'-' * 80}{Color.RESET}")


def print_success(text: str):
    """Print success message"""
    print(f"{Color.GREEN}✓ {text}{Color.RESET}")


def print_error(text: str):
    """Print error message"""
    print(f"{Color.RED}✗ {text}{Color.RESET}")


def print_warning(text: str):
    """Print warning message"""
    print(f"{Color.YELLOW}⚠ {text}{Color.RESET}")


def print_info(text: str):
    """Print info message"""
    print(f"{Color.CYAN}ℹ {text}{Color.RESET}")


def load_bank_accounts() -> Dict[str, Dict]:
    """Load bank account reference data from Excel file"""
    print_info("Loading bank account reference data...")

    try:
        df = pd.read_excel(BANK_ACCOUNT_FILE)

        bank_accounts = {}
        for _, row in df.iterrows():
            account_name = str(row.get('Bank Account Name', '')).strip()
            account_number = str(row.get('Bank Account Number', '')).strip()

            if account_name and account_number:
                bank_accounts[account_name] = {
                    'number': account_number,
                    'bank': row.get('Bank Name', ''),
                    'branch': row.get('Branch', ''),
                }

        print_success(f"Loaded {len(bank_accounts)} bank accounts")
        return bank_accounts

    except Exception as e:
        print_error(f"Error loading bank accounts: {e}")
        return {}


def strip_payment_prefix(transaction_number: str) -> str:
    """Remove payment prefix from transaction number for matching"""
    transaction_number = str(transaction_number).strip()

    for prefix in PAYMENT_PREFIXES:
        if transaction_number.startswith(prefix):
            return transaction_number[len(prefix):]

    return transaction_number


def load_csv_file(file_path: Path) -> List[Dict]:
    """Load CSV file and return list of dictionaries"""
    try:
        df = pd.read_csv(file_path)
        return df.to_dict('records')
    except Exception as e:
        print_error(f"Error loading {file_path}: {e}")
        return []


def load_oracle_file(folder_path: Path) -> List[Dict]:
    """Load Oracle CSV file (main payment file)"""
    # Look for the main CSV file (FOLDERNAME.csv)
    folder_name = folder_path.name
    oracle_file = folder_path / f"{folder_name}.csv"

    if not oracle_file.exists():
        print_error(f"Oracle file not found: {oracle_file}")
        return []

    print_info(f"Loading Oracle file: {oracle_file.name}")
    return load_csv_file(oracle_file)


def load_standard_receipts(folder_path: Path) -> List[Dict]:
    """Load standard consolidated receipt file"""
    receipt_file = folder_path / "Receipt_ALL_CONSOLIDATED.csv"

    if not receipt_file.exists():
        print_error(f"Standard receipt file not found: {receipt_file}")
        return []

    print_info(f"Loading standard receipts: {receipt_file.name}")
    return load_csv_file(receipt_file)


def load_oracle_receipts(folder_path: Path) -> List[Dict]:
    """Load Oracle fusion output receipts"""
    oracle_receipt_file = folder_path / "ORACLE_FUSION_OUTPUT" / "Receipts" / "Receipt_ALL_CONSOLIDATED.csv"

    if not oracle_receipt_file.exists():
        print_error(f"Oracle receipt file not found: {oracle_receipt_file}")
        return []

    print_info(f"Loading Oracle receipts: {oracle_receipt_file.name}")
    return load_csv_file(oracle_receipt_file)


def load_payment_file(folder_path: Path) -> pd.DataFrame:
    """Load payment Excel file"""
    # Look for payment file (usually ends with .xlsx and contains "payment")
    payment_files = list(folder_path.glob("*payment*.xlsx"))

    if not payment_files:
        print_warning("No payment Excel file found")
        return None

    payment_file = payment_files[0]
    print_info(f"Loading payment file: {payment_file.name}")

    try:
        df = pd.read_excel(payment_file)
        return df
    except Exception as e:
        print_error(f"Error loading payment file: {e}")
        return None


def verify_amounts(oracle_data: List[Dict], standard_receipts: List[Dict]) -> Dict:
    """Verify amounts between payment file and standard consolidated receipt"""
    print_section("1. Amount Verification")

    results = {
        'total_oracle': 0,
        'total_standard': 0,
        'matched': [],
        'mismatched': [],
        'missing_in_standard': [],
        'missing_in_oracle': [],
    }

    # Create lookup dictionaries
    oracle_amounts = {}
    for record in oracle_data:
        txn_number = strip_payment_prefix(str(record.get('Transaction Number', '')))
        if txn_number:
            amount = float(record.get('Transaction Line Amount', 0))
            if txn_number in oracle_amounts:
                oracle_amounts[txn_number] += amount
            else:
                oracle_amounts[txn_number] = amount

    standard_amounts = {}
    for record in standard_receipts:
        receipt_number = strip_payment_prefix(str(record.get('ReceiptNumber', '')))
        if receipt_number:
            amount = float(record.get('Amount', 0))
            standard_amounts[receipt_number] = amount

    results['total_oracle'] = sum(oracle_amounts.values())
    results['total_standard'] = sum(standard_amounts.values())

    # Compare amounts
    all_txns = set(oracle_amounts.keys()) | set(standard_amounts.keys())

    for txn in all_txns:
        oracle_amt = oracle_amounts.get(txn, 0)
        standard_amt = standard_amounts.get(txn, 0)

        if txn in oracle_amounts and txn in standard_amounts:
            if abs(oracle_amt - standard_amt) < 0.01:  # Allow for rounding differences
                results['matched'].append({
                    'transaction': txn,
                    'oracle_amount': oracle_amt,
                    'standard_amount': standard_amt
                })
            else:
                results['mismatched'].append({
                    'transaction': txn,
                    'oracle_amount': oracle_amt,
                    'standard_amount': standard_amt,
                    'difference': oracle_amt - standard_amt
                })
        elif txn in oracle_amounts:
            results['missing_in_standard'].append({
                'transaction': txn,
                'oracle_amount': oracle_amt
            })
        else:
            results['missing_in_oracle'].append({
                'transaction': txn,
                'standard_amount': standard_amt
            })

    # Print summary
    print_info(f"Total Oracle Amount: {results['total_oracle']:.2f}")
    print_info(f"Total Standard Amount: {results['total_standard']:.2f}")
    print_info(f"Difference: {abs(results['total_oracle'] - results['total_standard']):.2f}")
    print_success(f"Matched transactions: {len(results['matched'])}")

    if results['mismatched']:
        print_error(f"Mismatched amounts: {len(results['mismatched'])}")
    if results['missing_in_standard']:
        print_warning(f"Missing in standard: {len(results['missing_in_standard'])}")
    if results['missing_in_oracle']:
        print_warning(f"Missing in oracle: {len(results['missing_in_oracle'])}")

    return results


def verify_transactions(oracle_data: List[Dict], oracle_receipts: List[Dict]) -> Dict:
    """Verify transaction numbers from Oracle file to standard receipt ID"""
    print_section("2. Transaction Number Verification")

    results = {
        'matched': [],
        'mismatched': [],
        'missing_in_receipts': [],
        'missing_in_oracle': [],
    }

    # Extract transaction numbers
    oracle_txns = set()
    for record in oracle_data:
        txn = strip_payment_prefix(str(record.get('Transaction Number', '')))
        if txn:
            oracle_txns.add(txn)

    receipt_txns = set()
    for record in oracle_receipts:
        receipt_num = strip_payment_prefix(str(record.get('ReceiptNumber', '')))
        if receipt_num:
            receipt_txns.add(receipt_num)

    # Compare
    matched = oracle_txns & receipt_txns
    missing_in_receipts = oracle_txns - receipt_txns
    missing_in_oracle = receipt_txns - oracle_txns

    results['matched'] = list(matched)
    results['missing_in_receipts'] = list(missing_in_receipts)
    results['missing_in_oracle'] = list(missing_in_oracle)

    # Print summary
    print_success(f"Matched transactions: {len(matched)}")

    if missing_in_receipts:
        print_error(f"Transactions in Oracle but missing in Receipts: {len(missing_in_receipts)}")
    if missing_in_oracle:
        print_warning(f"Receipts not found in Oracle: {len(missing_in_oracle)}")

    return results


def verify_organization_id(oracle_receipts: List[Dict], misc_receipts: List[Dict] = None) -> Dict:
    """Verify Organization ID in receipt files"""
    print_section("3. Organization ID Verification")

    results = {
        'oracle_org_ids': set(),
        'misc_org_ids': set(),
        'valid_count': 0,
        'invalid_count': 0,
        'invalid_records': []
    }

    # Check Oracle receipts
    for record in oracle_receipts:
        org_id = str(record.get('BusinessUnit', '')).strip()
        results['oracle_org_ids'].add(org_id)

        if org_id == ORGANIZATION_ID:
            results['valid_count'] += 1
        else:
            results['invalid_count'] += 1
            results['invalid_records'].append({
                'receipt': record.get('ReceiptNumber', ''),
                'org_id': org_id,
                'source': 'Oracle Receipts'
            })

    # Check Misc receipts if provided
    if misc_receipts:
        for record in misc_receipts:
            org_id = str(record.get('BusinessUnit', '')).strip()
            results['misc_org_ids'].add(org_id)

            if org_id != ORGANIZATION_ID:
                results['invalid_count'] += 1
                results['invalid_records'].append({
                    'receipt': record.get('ReceiptNumber', ''),
                    'org_id': org_id,
                    'source': 'Misc Receipts'
                })

    # Print summary
    print_info(f"Expected Organization ID: {ORGANIZATION_ID}")
    print_info(f"Found Organization IDs: {', '.join(results['oracle_org_ids'] | results['misc_org_ids'])}")
    print_success(f"Valid records: {results['valid_count']}")

    if results['invalid_count'] > 0:
        print_error(f"Invalid Organization ID records: {results['invalid_count']}")

    return results


def verify_bank_accounts(oracle_receipts: List[Dict], bank_accounts: Dict) -> Dict:
    """Verify bank account details with standard file"""
    print_section("4. Bank Account Details Verification")

    results = {
        'matched': [],
        'mismatched': [],
        'not_found': [],
        'unique_accounts': set()
    }

    for record in oracle_receipts:
        bank_account_name = str(record.get('RemittanceBankAccountNumber', '')).strip()
        results['unique_accounts'].add(bank_account_name)

        if bank_account_name in bank_accounts:
            results['matched'].append({
                'receipt': record.get('ReceiptNumber', ''),
                'account': bank_account_name,
                'details': bank_accounts[bank_account_name]
            })
        else:
            results['not_found'].append({
                'receipt': record.get('ReceiptNumber', ''),
                'account': bank_account_name
            })

    # Print summary
    print_info(f"Unique bank accounts found: {len(results['unique_accounts'])}")
    print_success(f"Matched accounts: {len(results['matched'])}")

    if results['not_found']:
        print_warning(f"Bank accounts not found in reference: {len(results['not_found'])}")
        unique_not_found = set(r['account'] for r in results['not_found'])
        for account in unique_not_found:
            print_warning(f"  - {account}")

    return results


def verify_oracle_invoice_folder(folder_path: Path) -> Dict:
    """Verify Oracle invoice folder exists and files have correct naming"""
    print_section("5. Oracle Invoice Folder Verification")

    results = {
        'folder_exists': False,
        'folder_path': None,
        'total_files': 0,
        'valid_files': [],
        'invalid_files': [],
        'valid_count': 0,
        'invalid_count': 0
    }

    # Look for Oracle invoice folder (case-insensitive)
    oracle_invoice_patterns = [
        'Oracle invoice',
        'Oracle invoices',
        'oracle_invoice',
        'oracle_invoices',
        'Oracle Invoice',
        'Oracle Invoices',
        'ORACLE_INVOICE',
        'ORACLE INVOICE'
    ]

    oracle_invoice_folder = None
    for item in folder_path.iterdir():
        if item.is_dir():
            # Check if the folder name matches any pattern (case-insensitive)
            for pattern in oracle_invoice_patterns:
                if item.name.lower() == pattern.lower():
                    oracle_invoice_folder = item
                    results['folder_exists'] = True
                    results['folder_path'] = str(item)
                    break
            if oracle_invoice_folder:
                break

    if not oracle_invoice_folder:
        print_warning("Oracle invoice folder not found")
        return results

    print_success(f"Oracle invoice folder found: {oracle_invoice_folder.name}")

    # Check all CSV and Excel files in the folder
    file_extensions = ['.csv', '.xlsx', '.xls', '.CSV', '.XLSX', '.XLS']
    files_to_check = []

    for ext in file_extensions:
        files_to_check.extend(oracle_invoice_folder.glob(f'*{ext}'))

    results['total_files'] = len(files_to_check)

    if results['total_files'] == 0:
        print_warning("No CSV or Excel files found in Oracle invoice folder")
        return results

    print_info(f"Found {results['total_files']} file(s) to verify")

    # Verify each file has "Oracle invoice" in its name (case-insensitive)
    for file_path in files_to_check:
        file_name = file_path.name
        # Check if filename contains "oracle" and "invoice" (case-insensitive)
        if 'oracle' in file_name.lower() and 'invoice' in file_name.lower():
            results['valid_files'].append(file_name)
            results['valid_count'] += 1
        else:
            results['invalid_files'].append(file_name)
            results['invalid_count'] += 1

    # Print summary
    print_success(f"Files with correct naming: {results['valid_count']}")

    if results['invalid_count'] > 0:
        print_error(f"Files without 'Oracle invoice' in name: {results['invalid_count']}")
        for file_name in results['invalid_files']:
            print_error(f"  - {file_name}")

    return results


def generate_detailed_report(folder_name: str, verification_results: Dict) -> str:
    """Generate detailed verification report"""
    report_lines = []

    # Header
    report_lines.append("=" * 100)
    report_lines.append(f"PAYMENT VERIFICATION REPORT - {folder_name}")
    report_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("=" * 100)
    report_lines.append("")

    # Summary
    report_lines.append("SUMMARY")
    report_lines.append("-" * 100)

    amount_results = verification_results.get('amount_verification', {})
    report_lines.append(f"Total Oracle Amount: {amount_results.get('total_oracle', 0):.2f}")
    report_lines.append(f"Total Standard Amount: {amount_results.get('total_standard', 0):.2f}")
    report_lines.append(f"Amount Difference: {abs(amount_results.get('total_oracle', 0) - amount_results.get('total_standard', 0)):.2f}")
    report_lines.append(f"Matched Transactions: {len(amount_results.get('matched', []))}")
    report_lines.append(f"Mismatched Amounts: {len(amount_results.get('mismatched', []))}")
    report_lines.append("")

    txn_results = verification_results.get('transaction_verification', {})
    report_lines.append(f"Transaction Matches: {len(txn_results.get('matched', []))}")
    report_lines.append(f"Missing in Receipts: {len(txn_results.get('missing_in_receipts', []))}")
    report_lines.append(f"Missing in Oracle: {len(txn_results.get('missing_in_oracle', []))}")
    report_lines.append("")

    org_results = verification_results.get('organization_verification', {})
    report_lines.append(f"Valid Organization IDs: {org_results.get('valid_count', 0)}")
    report_lines.append(f"Invalid Organization IDs: {org_results.get('invalid_count', 0)}")
    report_lines.append("")

    bank_results = verification_results.get('bank_verification', {})
    report_lines.append(f"Matched Bank Accounts: {len(bank_results.get('matched', []))}")
    report_lines.append(f"Bank Accounts Not Found: {len(bank_results.get('not_found', []))}")
    report_lines.append("")

    oracle_invoice_results = verification_results.get('oracle_invoice_verification', {})
    report_lines.append(f"Oracle Invoice Folder Exists: {'Yes' if oracle_invoice_results.get('folder_exists') else 'No'}")
    if oracle_invoice_results.get('folder_exists'):
        report_lines.append(f"Oracle Invoice Files with Correct Naming: {oracle_invoice_results.get('valid_count', 0)}")
        report_lines.append(f"Oracle Invoice Files with Incorrect Naming: {oracle_invoice_results.get('invalid_count', 0)}")
    report_lines.append("")

    # Detailed sections
    if amount_results.get('mismatched'):
        report_lines.append("MISMATCHED AMOUNTS")
        report_lines.append("-" * 100)
        report_lines.append(f"{'Transaction':<40} {'Oracle Amount':>15} {'Standard Amount':>15} {'Difference':>15}")
        report_lines.append("-" * 100)
        for item in amount_results['mismatched']:
            report_lines.append(
                f"{item['transaction']:<40} "
                f"{item['oracle_amount']:>15.2f} "
                f"{item['standard_amount']:>15.2f} "
                f"{item['difference']:>15.2f}"
            )
        report_lines.append("")

    if amount_results.get('missing_in_standard'):
        report_lines.append("MISSING IN STANDARD RECEIPTS")
        report_lines.append("-" * 100)
        for item in amount_results['missing_in_standard'][:20]:  # Limit to first 20
            report_lines.append(f"  {item['transaction']}: {item['oracle_amount']:.2f}")
        if len(amount_results['missing_in_standard']) > 20:
            report_lines.append(f"  ... and {len(amount_results['missing_in_standard']) - 20} more")
        report_lines.append("")

    if txn_results.get('missing_in_receipts'):
        report_lines.append("TRANSACTIONS MISSING IN ORACLE RECEIPTS")
        report_lines.append("-" * 100)
        for txn in txn_results['missing_in_receipts'][:20]:  # Limit to first 20
            report_lines.append(f"  {txn}")
        if len(txn_results['missing_in_receipts']) > 20:
            report_lines.append(f"  ... and {len(txn_results['missing_in_receipts']) - 20} more")
        report_lines.append("")

    if org_results.get('invalid_records'):
        report_lines.append("INVALID ORGANIZATION IDs")
        report_lines.append("-" * 100)
        for record in org_results['invalid_records'][:20]:  # Limit to first 20
            report_lines.append(f"  Receipt: {record['receipt']}, Found: {record['org_id']}, Source: {record['source']}")
        if len(org_results['invalid_records']) > 20:
            report_lines.append(f"  ... and {len(org_results['invalid_records']) - 20} more")
        report_lines.append("")

    if bank_results.get('not_found'):
        report_lines.append("BANK ACCOUNTS NOT FOUND IN REFERENCE")
        report_lines.append("-" * 100)
        unique_not_found = {}
        for record in bank_results['not_found']:
            account = record['account']
            if account not in unique_not_found:
                unique_not_found[account] = []
            unique_not_found[account].append(record['receipt'])

        for account, receipts in unique_not_found.items():
            report_lines.append(f"  Account: {account}")
            report_lines.append(f"    Used in {len(receipts)} receipts")
        report_lines.append("")

    if oracle_invoice_results.get('invalid_files'):
        report_lines.append("ORACLE INVOICE FILES WITH INCORRECT NAMING")
        report_lines.append("-" * 100)
        report_lines.append("The following files do not have 'Oracle invoice' in their filename:")
        for file_name in oracle_invoice_results['invalid_files']:
            report_lines.append(f"  - {file_name}")
        report_lines.append("")

    report_lines.append("=" * 100)
    report_lines.append("END OF REPORT")
    report_lines.append("=" * 100)

    return "\n".join(report_lines)


def process_folder(folder_path: Path, bank_accounts: Dict) -> Dict:
    """Process a single folder and perform all verifications"""
    folder_name = folder_path.name
    print_header(f"Processing Folder: {folder_name}")

    results = {
        'folder_name': folder_name,
        'status': 'success',
        'errors': [],
        'amount_verification': {},
        'transaction_verification': {},
        'organization_verification': {},
        'bank_verification': {},
        'oracle_invoice_verification': {},
    }

    try:
        # Load data files
        oracle_data = load_oracle_file(folder_path)
        standard_receipts = load_standard_receipts(folder_path)
        oracle_receipts = load_oracle_receipts(folder_path)

        if not oracle_data:
            results['errors'].append("Failed to load Oracle data")
            results['status'] = 'failed'
            return results

        if not standard_receipts:
            results['errors'].append("Failed to load standard receipts")

        if not oracle_receipts:
            results['errors'].append("Failed to load Oracle receipts")
            results['status'] = 'failed'
            return results

        # Perform verifications
        if standard_receipts:
            results['amount_verification'] = verify_amounts(oracle_data, standard_receipts)

        results['transaction_verification'] = verify_transactions(oracle_data, oracle_receipts)
        results['organization_verification'] = verify_organization_id(oracle_receipts)
        results['bank_verification'] = verify_bank_accounts(oracle_receipts, bank_accounts)
        results['oracle_invoice_verification'] = verify_oracle_invoice_folder(folder_path)

        # Generate report
        report_content = generate_detailed_report(folder_name, results)

        # Save report
        report_file = folder_path / f"Verification_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(report_content)

        print_success(f"Report saved: {report_file.name}")

        # Save JSON summary
        summary_file = folder_path / f"Verification_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(summary_file, 'w', encoding='utf-8') as f:
            # Convert sets to lists for JSON serialization
            json_results = results.copy()
            if 'organization_verification' in json_results:
                org_v = json_results['organization_verification']
                if 'oracle_org_ids' in org_v:
                    org_v['oracle_org_ids'] = list(org_v['oracle_org_ids'])
                if 'misc_org_ids' in org_v:
                    org_v['misc_org_ids'] = list(org_v['misc_org_ids'])
            if 'bank_verification' in json_results:
                bank_v = json_results['bank_verification']
                if 'unique_accounts' in bank_v:
                    bank_v['unique_accounts'] = list(bank_v['unique_accounts'])

            json.dump(json_results, f, indent=2, default=str)

        print_success(f"Summary saved: {summary_file.name}")

    except Exception as e:
        print_error(f"Error processing folder {folder_name}: {e}")
        results['status'] = 'error'
        results['errors'].append(str(e))
        import traceback
        traceback.print_exc()

    return results


def main():
    """Main function to process all folders"""
    print_header("Payment Verification Automation")
    print_info(f"Base Directory: {BASE_DIR}")
    print_info(f"Organization ID: {ORGANIZATION_ID}")
    print("")

    # Load bank accounts
    bank_accounts = load_bank_accounts()

    if not bank_accounts:
        print_error("Failed to load bank account reference file. Continuing without bank verification.")

    # Get all store folders
    store_folders = []
    for item in BASE_DIR.iterdir():
        if item.is_dir() and not item.name.startswith('.') and item.name.isupper():
            # Check if it has required files
            has_oracle = (item / f"{item.name}.csv").exists()
            has_oracle_output = (item / "ORACLE_FUSION_OUTPUT").exists()

            if has_oracle or has_oracle_output:
                store_folders.append(item)

    store_folders.sort()

    print_info(f"Found {len(store_folders)} store folders to process")
    print("")

    if not store_folders:
        print_error("No store folders found!")
        return

    # Process each folder
    all_results = []
    success_count = 0
    failed_count = 0

    for folder in store_folders:
        result = process_folder(folder, bank_accounts)
        all_results.append(result)

        if result['status'] == 'success':
            success_count += 1
        else:
            failed_count += 1

        print("")  # Empty line between folders

    # Generate overall summary
    print_header("OVERALL SUMMARY")
    print_success(f"Successfully processed: {success_count} folders")

    if failed_count > 0:
        print_error(f"Failed to process: {failed_count} folders")

    print_info(f"Total folders: {len(store_folders)}")

    # Save overall summary
    summary_file = BASE_DIR / f"Overall_Verification_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(summary_file, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, indent=2, default=str)

    print_success(f"Overall summary saved: {summary_file.name}")

    print_header("Verification Complete!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user")
        sys.exit(1)
    except Exception as e:
        print_error(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

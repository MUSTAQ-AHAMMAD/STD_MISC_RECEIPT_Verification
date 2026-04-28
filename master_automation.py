#!/usr/bin/env python3
"""
Master Automation Script for Payment Verification

This script:
1. First runs the extraction script to extract any remaining zip files
2. Then runs the payment verification script

Usage: python master_automation.py
"""

import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent


def print_header(text):
    """Print a formatted header"""
    print("\n" + "=" * 80)
    print(text)
    print("=" * 80 + "\n")


def run_extraction():
    """Run the zip extraction script"""
    print_header("STEP 1: Extracting ZIP Files")

    extract_script = BASE_DIR / "extract_zip_files.py"

    if not extract_script.exists():
        print("⚠ Extraction script not found. Skipping extraction step.")
        return True

    try:
        result = subprocess.run(
            [sys.executable, str(extract_script)],
            capture_output=False,
            text=True
        )

        if result.returncode == 0:
            print("\n✓ Extraction completed successfully")
            return True
        else:
            print(f"\n⚠ Extraction completed with warnings (exit code: {result.returncode})")
            return True  # Continue anyway

    except Exception as e:
        print(f"✗ Error running extraction: {e}")
        print("Continuing with verification...")
        return True


def run_verification():
    """Run the payment verification script"""
    print_header("STEP 2: Payment Verification")

    verify_script = BASE_DIR / "payment_verification.py"

    if not verify_script.exists():
        print("✗ Verification script not found!")
        return False

    try:
        result = subprocess.run(
            [sys.executable, str(verify_script)],
            capture_output=False,
            text=True
        )

        if result.returncode == 0:
            print("\n✓ Verification completed successfully")
            return True
        else:
            print(f"\n✗ Verification failed with exit code: {result.returncode}")
            return False

    except Exception as e:
        print(f"✗ Error running verification: {e}")
        return False


def main():
    """Main function"""
    print_header("Payment Verification Master Automation")

    # Step 1: Extract zip files
    if not run_extraction():
        print("\n✗ Extraction failed. Aborting.")
        return 1

    # Step 2: Run verification
    if not run_verification():
        print("\n✗ Verification failed.")
        return 1

    print_header("✓ All Steps Completed Successfully!")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

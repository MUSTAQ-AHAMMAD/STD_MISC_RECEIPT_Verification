#!/usr/bin/env python3
"""
Script to extract zip files in folders after TABOUK and copy contents to subfolder root.
Each folder contains oracle_fusion_output (XX).zip that needs to be extracted and
the contents moved to the parent folder.
"""

import os
import zipfile
import shutil
from pathlib import Path

# Base directory
BASE_DIR = Path("/home/runner/work/STD_MISC_RECEIPT_Verification/STD_MISC_RECEIPT_Verification")

# Folders to process (TABOUK and after, alphabetically)
FOLDERS_TO_PROCESS = [
    "TABOUK",
    "TAIBAMED",
    "TAIFSHARST",
    "TAIFTERAML",
    "TALAMALL",
    "TAWAREN",
    "THEGATERYD",
    "TIFZHRNCTR",
    "TOWNSQJED",
    "VILLAGEHMT",
    "WADIDAWSER",
    "WADILABAN",
    "WESTAVENUE",
    "YASMEEN",
    "YASMEENPLZ",
    "ZAHRAN"
]

def extract_and_copy_zip(folder_path):
    """
    Extract zip file in the folder and copy contents to the folder root.

    Args:
        folder_path: Path object pointing to the folder
    """
    folder_name = folder_path.name

    # Find zip file
    zip_files = list(folder_path.glob("oracle_fusion_output*.zip"))

    if not zip_files:
        print(f"  ⚠️  No oracle_fusion_output zip file found in {folder_name}")
        return False

    if len(zip_files) > 1:
        print(f"  ⚠️  Multiple zip files found in {folder_name}, using first one")

    zip_file = zip_files[0]
    print(f"  📦 Found: {zip_file.name}")

    # Create temporary extraction directory
    extract_dir = folder_path / "temp_extract"
    extract_dir.mkdir(exist_ok=True)

    try:
        # Extract zip file
        print(f"  📂 Extracting {zip_file.name}...")
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # Find the extracted content (usually in ORACLE_FUSION_OUTPUT folder)
        extracted_items = list(extract_dir.iterdir())

        if not extracted_items:
            print(f"  ❌ No content found after extraction in {folder_name}")
            return False

        # Copy all extracted content to parent folder
        print(f"  📋 Copying contents to {folder_name} root...")
        for item in extracted_items:
            dest = folder_path / item.name

            # Skip if destination already exists
            if dest.exists():
                if dest.is_dir():
                    print(f"     Removing existing directory: {item.name}")
                    shutil.rmtree(dest)
                else:
                    print(f"     Removing existing file: {item.name}")
                    dest.unlink()

            # Copy item to destination
            if item.is_dir():
                shutil.copytree(item, dest)
                print(f"     ✅ Copied directory: {item.name}")
            else:
                shutil.copy2(item, dest)
                print(f"     ✅ Copied file: {item.name}")

        # Clean up temporary extraction directory
        shutil.rmtree(extract_dir)
        print(f"  🧹 Cleaned up temporary files")

        return True

    except Exception as e:
        print(f"  ❌ Error processing {folder_name}: {e}")
        # Clean up on error
        if extract_dir.exists():
            shutil.rmtree(extract_dir)
        return False

def main():
    """Main function to process all folders."""
    print("=" * 70)
    print("Starting extraction of zip files from TABOUK onwards...")
    print("=" * 70)
    print()

    success_count = 0
    fail_count = 0

    for folder_name in FOLDERS_TO_PROCESS:
        folder_path = BASE_DIR / folder_name

        print(f"Processing: {folder_name}")
        print("-" * 70)

        if not folder_path.exists():
            print(f"  ⚠️  Folder does not exist: {folder_name}")
            fail_count += 1
            print()
            continue

        if extract_and_copy_zip(folder_path):
            success_count += 1
            print(f"  ✅ Successfully processed {folder_name}")
        else:
            fail_count += 1
            print(f"  ❌ Failed to process {folder_name}")

        print()

    print("=" * 70)
    print("Extraction Summary")
    print("=" * 70)
    print(f"✅ Successful: {success_count}")
    print(f"❌ Failed: {fail_count}")
    print(f"📊 Total: {success_count + fail_count}")
    print("=" * 70)

if __name__ == "__main__":
    main()

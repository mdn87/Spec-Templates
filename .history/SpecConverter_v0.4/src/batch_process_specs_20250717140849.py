#!/usr/bin/env python3
"""
Batch Processing Script for Specification Documents

This script processes all specification documents in the examples/Specs folder,
extracting content and regenerating documents with proper formatting.

Usage:
    python batch_process_specs.py
"""

import os
import sys
import subprocess
import time
from pathlib import Path

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"\n{'='*60}")
    print(f"Running: {description}")
    print(f"Command: {command}")
    print(f"{'='*60}")
    
    start_time = time.time()
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        end_time = time.time()
        
        if result.returncode == 0:
            print(f"✅ SUCCESS: {description}")
            print(f"⏱️  Time taken: {end_time - start_time:.2f} seconds")
            if result.stdout:
                print("Output:")
                print(result.stdout)
        else:
            print(f"❌ FAILED: {description}")
            print(f"⏱️  Time taken: {end_time - start_time:.2f} seconds")
            print("Error output:")
            print(result.stderr)
            return False
            
    except Exception as e:
        print(f"❌ EXCEPTION: {description}")
        print(f"Error: {e}")
        return False
    
    return True

def main():
    """Main batch processing function"""
    print("🚀 Starting Batch Processing of Specification Documents")
    print("=" * 80)
    
    # Define paths
    examples_dir = Path("../examples/Specs")
    output_dir = Path("../output/Specs")
    template_path = "../templates/test_template_cleaned.docx"
    
    # Check if directories exist
    if not examples_dir.exists():
        print(f"❌ Error: Examples directory not found: {examples_dir}")
        return
    
    if not Path(template_path).exists():
        print(f"❌ Error: Template file not found: {template_path}")
        return
    
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get all .docx files in the examples directory
    docx_files = list(examples_dir.glob("*.docx"))
    
    if not docx_files:
        print(f"❌ No .docx files found in {examples_dir}")
        return
    
    print(f"📁 Found {len(docx_files)} specification documents to process:")
    for docx_file in docx_files:
        print(f"   - {docx_file.name}")
    
    # Process each document
    successful = 0
    failed = 0
    
    for i, docx_file in enumerate(docx_files, 1):
        print(f"\n{'#'*80}")
        print(f"📄 Processing document {i}/{len(docx_files)}: {docx_file.name}")
        print(f"{'#'*80}")
        
        # Extract content
        extract_command = f'python extract_spec_content_v3.py "{docx_file}" . "{template_path}"'
        if run_command(extract_command, f"Extracting content from {docx_file.name}"):
            # Regenerate document
            # Update the test_generator.py to use the correct output path
            base_name = docx_file.stem
            output_filename = f"{base_name}_regenerated.docx"
            output_path = output_dir / output_filename
            
            # Temporarily modify test_generator.py to use the correct output
            with open("test_generator.py", "r", encoding="utf-8") as f:
                content = f.read()
            
            # Replace the output path
            modified_content = content.replace(
                'OUTPUT_PATH   = \'../output/generated_spec_v3_fixed_new2.docx\'',
                f'OUTPUT_PATH   = \'{output_path}\''
            )
            
            with open("test_generator.py", "w", encoding="utf-8") as f:
                f.write(modified_content)
            
            # Run regeneration
            regenerate_command = "python test_generator.py"
            if run_command(regenerate_command, f"Regenerating document for {docx_file.name}"):
                successful += 1
                print(f"✅ Successfully processed: {docx_file.name} → {output_filename}")
            else:
                failed += 1
                print(f"❌ Failed to regenerate: {docx_file.name}")
        else:
            failed += 1
            print(f"❌ Failed to extract: {docx_file.name}")
    
    # Restore original test_generator.py
    with open("test_generator.py", "w", encoding="utf-8") as f:
        f.write(content)
    
    # Summary
    print(f"\n{'='*80}")
    print("📊 BATCH PROCESSING SUMMARY")
    print(f"{'='*80}")
    print(f"✅ Successful: {successful}/{len(docx_files)}")
    print(f"❌ Failed: {failed}/{len(docx_files)}")
    print(f"📁 Output location: {output_dir}")
    
    if successful > 0:
        print(f"\n📋 Successfully processed documents:")
        for docx_file in docx_files:
            base_name = docx_file.stem
            output_filename = f"{base_name}_regenerated.docx"
            output_path = output_dir / output_filename
            if output_path.exists():
                print(f"   ✅ {docx_file.name} → {output_filename}")
    
    if failed > 0:
        print(f"\n⚠️  Failed documents:")
        for docx_file in docx_files:
            base_name = docx_file.stem
            output_filename = f"{base_name}_regenerated.docx"
            output_path = output_dir / output_filename
            if not output_path.exists():
                print(f"   ❌ {docx_file.name}")
    
    print(f"\n🎉 Batch processing complete!")

if __name__ == "__main__":
    main() 
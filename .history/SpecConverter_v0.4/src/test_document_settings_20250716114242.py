#!/usr/bin/env python3
"""
Test script to verify document-level settings extraction and regeneration
"""

from docx import Document
import json
import os
from header_footer_extractor import HeaderFooterExtractor

def test_document_settings_extraction():
    """Test extraction of document-level settings"""
    print("Testing document-level settings extraction...")
    
    # Test extraction from original document
    original_doc_path = "../examples/SECTION 26 05 00.docx"
    
    if not os.path.exists(original_doc_path):
        print(f"Original document not found: {original_doc_path}")
        return
    
    # Extract settings from original
    extractor = HeaderFooterExtractor()
    original_data = extractor.extract_header_footer_margins(original_doc_path)
    
    print("Original document settings:")
    if original_data.get('document_settings'):
        for key, value in original_data['document_settings'].items():
            print(f"  {key}: {value}")
    else:
        print("  No document settings found")
    
    print("\nOriginal margins:")
    if original_data.get('margins'):
        for key, value in original_data['margins'].items():
            print(f"  {key}: {value} inches")
    else:
        print("  No margins found")
    
    return original_data

def test_document_settings_regeneration():
    """Test regeneration of document-level settings"""
    print("\nTesting document-level settings regeneration...")
    
    # Test extraction from regenerated document
    regenerated_doc_path = "../output/generated_spec_v3.docx"
    
    if not os.path.exists(regenerated_doc_path):
        print(f"Regenerated document not found: {regenerated_doc_path}")
        return
    
    # Extract settings from regenerated
    extractor = HeaderFooterExtractor()
    regenerated_data = extractor.extract_header_footer_margins(regenerated_doc_path)
    
    print("Regenerated document settings:")
    if regenerated_data.get('document_settings'):
        for key, value in regenerated_data['document_settings'].items():
            print(f"  {key}: {value}")
    else:
        print("  No document settings found")
    
    print("\nRegenerated margins:")
    if regenerated_data.get('margins'):
        for key, value in regenerated_data['margins'].items():
            print(f"  {key}: {value} inches")
    else:
        print("  No margins found")
    
    return regenerated_data

def compare_document_settings(original_data, regenerated_data):
    """Compare original and regenerated document settings"""
    print("\nComparing document settings...")
    
    original_settings = original_data.get('document_settings', {})
    regenerated_settings = regenerated_data.get('document_settings', {})
    
    original_margins = original_data.get('margins', {})
    regenerated_margins = regenerated_data.get('margins', {})
    
    # Compare key settings
    key_settings = ['page_width', 'page_height', 'page_orientation', 'different_first_page_header_footer']
    
    print("Document settings comparison:")
    for key in key_settings:
        orig_val = original_settings.get(key)
        reg_val = regenerated_settings.get(key)
        if orig_val == reg_val:
            print(f"  ✓ {key}: {orig_val}")
        else:
            print(f"  ✗ {key}: {orig_val} vs {reg_val}")
    
    print("\nMargin settings comparison:")
    margin_keys = ['top_margin', 'bottom_margin', 'left_margin', 'right_margin', 'header_distance', 'footer_distance']
    
    for key in margin_keys:
        orig_val = original_margins.get(key)
        reg_val = regenerated_margins.get(key)
        if orig_val == reg_val:
            print(f"  ✓ {key}: {orig_val} inches")
        else:
            print(f"  ✗ {key}: {orig_val} vs {reg_val} inches")

def test_json_document_settings():
    """Test that document settings are included in JSON output"""
    print("\nTesting JSON document settings...")
    
    json_path = "../output/SECTION 26 05 00_v3.json"
    
    if not os.path.exists(json_path):
        print(f"JSON file not found: {json_path}")
        return
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print("JSON document settings:")
    if data.get('document_settings'):
        for key, value in data['document_settings'].items():
            print(f"  {key}: {value}")
    else:
        print("  No document settings found in JSON")
    
    print("\nJSON margins:")
    if data.get('margins'):
        for key, value in data['margins'].items():
            print(f"  {key}: {value} inches")
    else:
        print("  No margins found in JSON")

def main():
    """Main test function"""
    print("Document-Level Settings Test")
    print("=" * 50)
    
    # Test extraction
    original_data = test_document_settings_extraction()
    
    # Test regeneration
    regenerated_data = test_document_settings_regeneration()
    
    # Compare settings
    if original_data and regenerated_data:
        compare_document_settings(original_data, regenerated_data)
    
    # Test JSON
    test_json_document_settings()
    
    print("\nTest completed!")

if __name__ == "__main__":
    main() 
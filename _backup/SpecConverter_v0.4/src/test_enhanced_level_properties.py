#!/usr/bin/env python3
"""
Test script to verify the enhanced level list properties extraction and regeneration
"""

import json
import os
from extract_spec_content_v3 import SpecContentExtractorV3

def test_enhanced_properties_extraction():
    """Test extraction of enhanced level list properties"""
    print("Testing enhanced level list properties extraction...")
    
    # Use the existing extraction result
    json_path = "../output/SECTION 26 05 00_v3.json"
    
    if not os.path.exists(json_path):
        print("Extraction result not found, please run extraction first")
        return
    
    # Load the JSON data
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Check for enhanced properties in content blocks
    content_blocks = data.get('content_blocks', [])
    enhanced_properties_found = 0
    
    for i, block in enumerate(content_blocks):
        # Check if any enhanced properties are set
        if (block.get('number_alignment') or 
            block.get('aligned_at') or 
            block.get('text_indent_at') or 
            block.get('follow_number_with') or 
            block.get('add_tab_stop_at') or 
            block.get('link_level_to_style')):
            
            enhanced_properties_found += 1
            print(f"Block {i+1} ({block.get('level_type', 'unknown')}):")
            print(f"  number_alignment: {block.get('number_alignment')}")
            print(f"  aligned_at: {block.get('aligned_at')}")
            print(f"  text_indent_at: {block.get('text_indent_at')}")
            print(f"  follow_number_with: {block.get('follow_number_with')}")
            print(f"  add_tab_stop_at: {block.get('add_tab_stop_at')}")
            print(f"  link_level_to_style: {block.get('link_level_to_style')}")
            print()
    
    print(f"Found {enhanced_properties_found} blocks with enhanced level list properties")
    
    # Check template analysis for numbering definitions
    template_analysis = data.get('template_analysis', {})
    numbering_definitions = template_analysis.get('template_numbering', {})
    
    print(f"\nTemplate numbering definitions: {len(numbering_definitions)} found")
    
    # Show some example numbering definitions
    count = 0
    for key, value in numbering_definitions.items():
        if count < 3:  # Show first 3
            print(f"  {key}: {type(value)}")
            if isinstance(value, dict) and 'levels' in value:
                print(f"    Levels: {len(value['levels'])}")
                for level_key, level_info in value['levels'].items():
                    if count < 3:
                        print(f"      Level {level_key}:")
                        print(f"        lvlJc: {level_info.get('lvlJc')}")
                        print(f"        suff: {level_info.get('suff')}")
                        print(f"        pStyle: {level_info.get('pStyle')}")
                        if 'pPr' in level_info and 'indent' in level_info['pPr']:
                            print(f"        indent: {level_info['pPr']['indent']}")
        count += 1

def test_regeneration_with_enhanced_properties():
    """Test regeneration with enhanced properties"""
    print("\nTesting regeneration with enhanced properties...")
    
    # Import the test generator
    from test_generator import parse_spec_json, generate_content_from_v3_json
    from docx import Document
    
    # Use the existing extraction result
    json_path = "../output/SECTION 26 05 00_v3.json"
    
    if not os.path.exists(json_path):
        print("Extraction result not found, please run extraction first")
        return
    
    # Parse the JSON
    json_data = parse_spec_json(json_path)
    if not json_data:
        print("Failed to parse JSON data")
        return
    
    # Create a new document
    doc = Document()
    
    # Generate content
    generate_content_from_v3_json(doc, json_data)
    
    # Save the regenerated document
    output_path = "../output/test_enhanced_properties_regeneration.docx"
    doc.save(output_path)
    
    print(f"Regeneration completed successfully: {output_path}")
    print("Note: Enhanced level list properties are preserved in the JSON and")
    print("can be used for advanced formatting during regeneration.")

if __name__ == "__main__":
    test_enhanced_properties_extraction()
    test_regeneration_with_enhanced_properties() 
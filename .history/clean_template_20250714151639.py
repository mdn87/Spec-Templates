#!/usr/bin/env python3
"""
Template Cleaner Script

This script cleans a Word template by removing numbering definitions that don't have
a "BWA" list name, and outputs a cleaned .docx file with "_cleaned" suffix.

Usage:
    python clean_template.py <template_file.docx>

Example:
    python clean_template.py test_template.docx
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import sys
import shutil
from typing import Set, Dict, Any, Optional

def extract_numbering_info(docx_path: str) -> Optional[Dict[str, Any]]:
    """Extract numbering information from a .docx file"""
    numbering_info = {
        "abstract_nums": {},
        "nums": {},
        "list_names": {},
        "to_remove": set()
    }
    
    try:
        with zipfile.ZipFile(docx_path) as zf:
            # Extract numbering.xml
            if "word/numbering.xml" in zf.namelist():
                num_xml = zf.read("word/numbering.xml")
                root = ET.fromstring(num_xml)
                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                
                # Find all abstract numbering definitions
                for abstract_num in root.findall(".//w:abstractNum", ns):
                    abstract_num_id = abstract_num.get(f"{{{ns['w']}}}abstractNumId")
                    
                    # Check for list name
                    list_name_elem = abstract_num.find("w:lvl[1]/w:lvlText", ns)
                    list_name = ""
                    if list_name_elem is not None:
                        list_name = list_name_elem.get(f"{{{ns['w']}}}val", "")
                    
                    numbering_info["abstract_nums"][abstract_num_id] = {
                        "element": abstract_num,
                        "list_name": list_name,
                        "has_bwa": "BWA" in list_name.upper()
                    }
                    
                    # Mark for removal if no BWA in list name
                    if not numbering_info["abstract_nums"][abstract_num_id]["has_bwa"]:
                        numbering_info["to_remove"].add(abstract_num_id)
                        print(f"Marking abstractNum {abstract_num_id} for removal (list name: '{list_name}')")
                    else:
                        print(f"Keeping abstractNum {abstract_num_id} (BWA list name: '{list_name}')")
                
                # Find all num elements that reference abstract numbers
                for num_elem in root.findall(".//w:num", ns):
                    num_id = num_elem.get(f"{{{ns['w']}}}numId")
                    abstract_num_ref = num_elem.find("w:abstractNumId", ns)
                    
                    if abstract_num_ref is not None:
                        abstract_num_id_ref = abstract_num_ref.get(f"{{{ns['w']}}}val")
                        numbering_info["nums"][num_id] = {
                            "element": num_elem,
                            "abstract_num_id": abstract_num_id_ref,
                            "to_remove": abstract_num_id_ref in numbering_info["to_remove"]
                        }
                        
                        if numbering_info["nums"][num_id]["to_remove"]:
                            print(f"Marking num {num_id} for removal (references abstractNum {abstract_num_id_ref})")
                        else:
                            print(f"Keeping num {num_id} (references BWA abstractNum {abstract_num_id_ref})")
            
            # Extract document.xml to check for paragraph references
            if "word/document.xml" in zf.namelist():
                doc_xml = zf.read("word/document.xml")
                doc_root = ET.fromstring(doc_xml)
                
                # Find paragraphs that reference numbering
                for para in doc_root.findall(".//w:p", ns):
                    num_pr = para.find("w:pPr/w:numPr", ns)
                    if num_pr is not None:
                        num_id_elem = num_pr.find("w:numId", ns)
                        if num_id_elem is not None:
                            num_id = num_id_elem.get(f"{{{ns['w']}}}val")
                            if num_id in numbering_info["nums"]:
                                if numbering_info["nums"][num_id]["to_remove"]:
                                    print(f"Warning: Paragraph references numId {num_id} that will be removed")
            
    except Exception as e:
        print(f"Error extracting numbering info: {e}")
        return None
    
    return numbering_info

def clean_numbering_xml(numbering_info: Dict[str, Any]) -> Optional[str]:
    """Create cleaned numbering.xml content"""
    try:
        # Create a new numbering.xml structure
        root = ET.Element("w:numbering")
        root.set("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        
        # Add only the BWA abstract numbering definitions
        for abstract_num_id, info in numbering_info["abstract_nums"].items():
            if info["has_bwa"]:
                # Deep copy the element
                new_abstract_num = ET.SubElement(root, "w:abstractNum")
                new_abstract_num.set(f"{{{ns['w']}}}abstractNumId", abstract_num_id)
                
                # Copy all child elements
                for child in info["element"]:
                    new_abstract_num.append(child)
        
        # Add only the num elements that reference BWA abstract numbers
        for num_id, info in numbering_info["nums"].items():
            if not info["to_remove"]:
                # Deep copy the element
                new_num = ET.SubElement(root, "w:num")
                new_num.set(f"{{{ns['w']}}}numId", num_id)
                
                # Copy all child elements
                for child in info["element"]:
                    new_num.append(child)
        
        # Convert to string
        return ET.tostring(root, encoding='unicode', xml_declaration=True)
        
    except Exception as e:
        print(f"Error creating cleaned numbering.xml: {e}")
        return None

def clean_document_xml(docx_path: str, numbering_info: Dict[str, Any]) -> Optional[str]:
    """Create cleaned document.xml content by removing references to deleted numbering"""
    try:
        with zipfile.ZipFile(docx_path) as zf:
            if "word/document.xml" in zf.namelist():
                doc_xml = zf.read("word/document.xml")
                root = ET.fromstring(doc_xml)
                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                
                # Find and remove numbering references that point to deleted numIds
                for para in root.findall(".//w:p", ns):
                    p_pr = para.find("w:pPr", ns)
                    if p_pr is not None:
                        num_pr = p_pr.find("w:numPr", ns)
                        if num_pr is not None:
                            num_id_elem = num_pr.find("w:numId", ns)
                            if num_id_elem is not None:
                                num_id = num_id_elem.get(f"{{{ns['w']}}}val")
                                if num_id in numbering_info["nums"] and numbering_info["nums"][num_id]["to_remove"]:
                                    # Remove the entire numPr element
                                    p_pr.remove(num_pr)
                                    print(f"Removed numbering reference numId {num_id} from paragraph")
                
                # Convert to string
                return ET.tostring(root, encoding='unicode', xml_declaration=True)
        
    except Exception as e:
        print(f"Error cleaning document.xml: {e}")
        return None

def create_cleaned_docx(original_path: str, numbering_info: Dict[str, Any]) -> str:
    """Create a cleaned .docx file"""
    try:
        # Generate output filename
        base_name = os.path.splitext(original_path)[0]
        output_path = f"{base_name}_cleaned.docx"
        
        # Create a copy of the original file
        shutil.copy2(original_path, output_path)
        
        # Update the copy with cleaned content
        with zipfile.ZipFile(output_path, 'a') as zf:
            # Remove original numbering.xml
            if "word/numbering.xml" in zf.namelist():
                zf.remove("word/numbering.xml")
            
            # Add cleaned numbering.xml
            cleaned_numbering = clean_numbering_xml(numbering_info)
            if cleaned_numbering:
                zf.writestr("word/numbering.xml", cleaned_numbering)
                print(f"Added cleaned numbering.xml")
            
            # Update document.xml
            cleaned_document = clean_document_xml(original_path, numbering_info)
            if cleaned_document:
                zf.remove("word/document.xml")
                zf.writestr("word/document.xml", cleaned_document)
                print(f"Updated document.xml")
        
        return output_path
        
    except Exception as e:
        print(f"Error creating cleaned .docx: {e}")
        return None

def main():
    """Main function"""
    if len(sys.argv) != 2:
        print("Usage: python clean_template.py <template_file.docx>")
        print("Example: python clean_template.py test_template.docx")
        sys.exit(1)
    
    template_path = sys.argv[1]
    
    if not os.path.exists(template_path):
        print(f"Error: File '{template_path}' not found.")
        sys.exit(1)
    
    print(f"Cleaning template: {template_path}")
    print("=" * 50)
    
    # Extract numbering information
    numbering_info = extract_numbering_info(template_path)
    if not numbering_info:
        print("Failed to extract numbering information.")
        sys.exit(1)
    
    print(f"\nSummary:")
    print(f"- Total abstract numbering definitions: {len(numbering_info['abstract_nums'])}")
    print(f"- Total num mappings: {len(numbering_info['nums'])}")
    print(f"- To be removed: {len(numbering_info['to_remove'])}")
    
    if not numbering_info['to_remove']:
        print("\nNo non-BWA numbering definitions found. Template is already clean.")
        return
    
    # Create cleaned .docx
    print(f"\nCreating cleaned template...")
    output_path = create_cleaned_docx(template_path, numbering_info)
    
    if output_path:
        print(f"\n✅ Successfully created cleaned template: {output_path}")
        
        # Verify the cleaned template
        print(f"\nVerifying cleaned template...")
        cleaned_info = extract_numbering_info(output_path)
        if cleaned_info:
            print(f"Cleaned template has {len(cleaned_info['abstract_nums'])} BWA numbering definitions")
    else:
        print(f"\n❌ Failed to create cleaned template")

if __name__ == "__main__":
    main() 
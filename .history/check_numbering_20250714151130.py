#!/usr/bin/env python3
import zipfile
import xml.etree.ElementTree as ET

def check_numbering_xml(docx_path):
    """Extract and display numbering.xml content"""
    try:
        with zipfile.ZipFile(docx_path) as zf:
            if "word/numbering.xml" in zf.namelist():
                num_xml = zf.read("word/numbering.xml")
                root = ET.fromstring(num_xml)
                
                print(f"Numbering definitions in {docx_path}:")
                print("=" * 50)
                
                for abstract_num in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNum"):
                    abstract_num_id = abstract_num.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId")
                    print(f"\nAbstract Numbering Definition {abstract_num_id}:")
                    
                    for lvl in abstract_num.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvl"):
                        ilvl = lvl.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl")
                        
                        # Get lvlText
                        lvl_text_elem = lvl.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlText")
                        lvl_text = lvl_text_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if lvl_text_elem is not None else "None"
                        
                        # Get numFmt
                        num_fmt_elem = lvl.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numFmt")
                        num_fmt = num_fmt_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if num_fmt_elem is not None else "None"
                        
                        print(f"  Level {ilvl}: pattern='{lvl_text}', format='{num_fmt}'")
            else:
                print("No numbering.xml found in the document")
                
    except Exception as e:
        print(f"Error reading numbering.xml: {e}")

if __name__ == "__main__":
    check_numbering_xml("test_template.docx") 
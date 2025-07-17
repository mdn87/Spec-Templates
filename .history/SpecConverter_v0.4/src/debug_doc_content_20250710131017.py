#!/usr/bin/env python3
"""
Debug script to examine the actual content structure of a Word document
"""

from docx import Document
import sys

def debug_document_content(docx_path):
    """Print the raw content of the document to understand its structure"""
    doc = Document(docx_path)
    
    print(f"Document: {docx_path}")
    print("=" * 60)
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            style = para.style.name if para.style else "No Style"
            print(f"Line {i+1:2d} [{style:15s}]: {text}")
            
            # Check for numbering
            try:
                if para._element.pPr and para._element.pPr.numPr:
                    numPr = para._element.pPr.numPr
                    if numPr.numId and numPr.ilvl:
                        print(f"           Numbering: Level {numPr.ilvl.val}")
            except:
                pass

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python debug_doc_content.py <docx_file>")
        sys.exit(1)
    
    debug_document_content(sys.argv[1]) 
from docx import Document
import json
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def extract_text_from_element(element, nsmap):
    """Extract all text from an element and its children"""
    texts = []
    for text_elem in element.findall('.//w:t', namespaces=nsmap):
        if text_elem.text:
            texts.append(text_elem.text)
    return ''.join(texts).strip()

def extract_content_from_section(section_element, nsmap):
    """Extract content from header or footer section"""
    content = {
        "paragraphs": [],
        "tables": [],
        "text_boxes": []
    }
    
    # Extract paragraphs
    for p in section_element.findall('.//w:p', namespaces=nsmap):
        text = extract_text_from_element(p, nsmap)
        if text:
            content["paragraphs"].append(text)
    
    # Extract tables
    for tbl in section_element.findall('.//w:tbl', namespaces=nsmap):
        table_data = []
        for row in tbl.findall('.//w:tr', namespaces=nsmap):
            row_data = []
            for cell in row.findall('.//w:tc', namespaces=nsmap):
                cell_text = extract_text_from_element(cell, nsmap)
                row_data.append(cell_text)
            if row_data:
                table_data.append(row_data)
        if table_data:
            content["tables"].append(table_data)
    
    # Extract text boxes
    for drawing in section_element.findall('.//w:txbxContent', namespaces=nsmap):
        textbox_data = []
        for p in drawing.findall('.//w:p', namespaces=nsmap):
            text = extract_text_from_element(p, nsmap)
            if text:
                textbox_data.append(text)
        if textbox_data:
            content["text_boxes"].append(textbox_data)
    
    return content

def extract_header_and_footer(docx_path):
    """Extract header and footer information and margin settings from Word document"""
    doc = Document(docx_path)
    sec = doc.sections[0]
    
    # Extract margin settings with null checks
    margins = {}
    try:
        if sec.top_margin:
            margins["top_margin"] = sec.top_margin.inches
        if sec.bottom_margin:
            margins["bottom_margin"] = sec.bottom_margin.inches
        if sec.left_margin:
            margins["left_margin"] = sec.left_margin.inches
        if sec.right_margin:
            margins["right_margin"] = sec.right_margin.inches
        if sec.header_distance:
            margins["header_distance"] = sec.header_distance.inches
        if sec.footer_distance:
            margins["footer_distance"] = sec.footer_distance.inches
    except Exception as e:
        print(f"Warning: Could not extract margin settings: {e}")

    # Extract header content
    header_content = {}
    try:
        if sec.header:
            header_element = sec.header._element
            header_content = extract_content_from_section(header_element, header_element.nsmap)
    except Exception as e:
        print(f"Warning: Could not extract header content: {e}")

    # Extract footer content
    footer_content = {}
    try:
        if sec.footer:
            footer_element = sec.footer._element
            footer_content = extract_content_from_section(footer_element, footer_element.nsmap)
    except Exception as e:
        print(f"Warning: Could not extract footer content: {e}")

    return {
        "header": header_content,
        "footer": footer_content,
        "margins": margins
    }

def save_header_footer_to_json(data, output_path):
    """Save header/footer data to JSON file"""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)
    print(f"Header/footer data saved to JSON: {output_path}")

def save_header_footer_to_txt(data, output_path):
    """Save header/footer data to TXT file"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("HEADER, FOOTER, AND MARGIN DATA FROM DOCUMENT\n")
        f.write("=" * 60 + "\n\n")
        
        # Header information
        f.write("HEADER CONTENT:\n")
        f.write("-" * 20 + "\n")
        if data.get('header'):
            header = data['header']
            if header.get('paragraphs'):
                f.write("Paragraphs:\n")
                for i, para in enumerate(header['paragraphs'], 1):
                    f.write(f"  {i}. {para}\n")
            
            if header.get('tables'):
                f.write("\nTables:\n")
                for i, table in enumerate(header['tables'], 1):
                    f.write(f"  Table {i}:\n")
                    for j, row in enumerate(table, 1):
                        f.write(f"    Row {j}: {row}\n")
            
            if header.get('text_boxes'):
                f.write("\nText Boxes:\n")
                for i, textbox in enumerate(header['text_boxes'], 1):
                    f.write(f"  Text Box {i}:\n")
                    for j, para in enumerate(textbox, 1):
                        f.write(f"    {j}. {para}\n")
        else:
            f.write("No header content found\n")
        
        # Footer information
        f.write("\n\nFOOTER CONTENT:\n")
        f.write("-" * 20 + "\n")
        if data.get('footer'):
            footer = data['footer']
            if footer.get('paragraphs'):
                f.write("Paragraphs:\n")
                for i, para in enumerate(footer['paragraphs'], 1):
                    f.write(f"  {i}. {para}\n")
            
            if footer.get('tables'):
                f.write("\nTables:\n")
                for i, table in enumerate(footer['tables'], 1):
                    f.write(f"  Table {i}:\n")
                    for j, row in enumerate(table, 1):
                        f.write(f"    Row {j}: {row}\n")
            
            if footer.get('text_boxes'):
                f.write("\nText Boxes:\n")
                for i, textbox in enumerate(footer['text_boxes'], 1):
                    f.write(f"  Text Box {i}:\n")
                    for j, para in enumerate(textbox, 1):
                        f.write(f"    {j}. {para}\n")
        else:
            f.write("No footer content found\n")
        
        # Margin settings
        f.write("\n\nMARGIN SETTINGS:\n")
        f.write("-" * 20 + "\n")
        if data.get('margins'):
            for key, value in data['margins'].items():
                f.write(f"{key.replace('_', ' ').title()}: {value} inches\n")
        else:
            f.write("No margin information found\n")
    
    print(f"Header/footer data saved to TXT: {output_path}")

if __name__ == "__main__":
    try:
        docx_file = "SECTION 26 05 00.docx"
        data = extract_header_and_footer(docx_file)
        
        if data:
            # Save to JSON
            json_file = "SECTION 26 05 00_header_footer.json"
            save_header_footer_to_json(data, json_file)
            
            # Save to TXT
            txt_file = "SECTION 26 05 00_header_footer.txt"
            save_header_footer_to_txt(data, txt_file)
            
            print(f"\nExtracted header, footer, and margin data from {docx_file}")
            print("Files created:")
            print(f"  - {json_file}")
            print(f"  - {txt_file}")
            
            # Show a preview of the extracted data
            print("\nData Preview:")
            
            # Header preview
            if data.get('header'):
                header = data['header']
                print("  Header Content:")
                if header.get('paragraphs'):
                    print(f"    Paragraphs: {len(header['paragraphs'])} found")
                if header.get('tables'):
                    print(f"    Tables: {len(header['tables'])} found")
                if header.get('text_boxes'):
                    print(f"    Text Boxes: {len(header['text_boxes'])} found")
            else:
                print("  Header Content: None found")
            
            # Footer preview
            if data.get('footer'):
                footer = data['footer']
                print("  Footer Content:")
                if footer.get('paragraphs'):
                    print(f"    Paragraphs: {len(footer['paragraphs'])} found")
                if footer.get('tables'):
                    print(f"    Tables: {len(footer['tables'])} found")
                if footer.get('text_boxes'):
                    print(f"    Text Boxes: {len(footer['text_boxes'])} found")
            else:
                print("  Footer Content: None found")
            
            # Margins preview
            if data.get('margins'):
                print("  Margin Settings:")
                for key, value in list(data['margins'].items())[:3]:  # Show first 3 items
                    print(f"    {key.replace('_', ' ').title()}: {value} inches")
        else:
            print("No header/footer data found to save")
            
    except FileNotFoundError:
        print("File 'SECTION 26 05 00.docx' not found")
    except Exception as e:
        print(f"Error processing document: {e}")

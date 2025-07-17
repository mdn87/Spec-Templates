#!/usr/bin/env python3
"""
Simple Template Style Analyzer

This script analyzes paragraph styles in a Word document template to extract
key formatting information including spacing, indentation, and level list properties.

Usage:
    python simple_style_analyzer.py <template_file.docx>
"""

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import json
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

class SimpleStyleAnalyzer:
    """Analyzes styles in a Word document template"""
    
    def __init__(self):
        self.styles = {}
        self.namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    
    def analyze_template(self, template_path):
        """Analyze all styles in the template document"""
        try:
            print(f"Analyzing styles in template: {template_path}")
            doc = Document(template_path)
            
            # Extract numbering definitions first
            numbering_definitions = self.extract_numbering_definitions(template_path)
            
            # Analyze all styles
            for style in doc.styles:
                style_info = self.extract_style_info(style, numbering_definitions)
                if style_info:
                    self.styles[style.name] = style_info
            
            # Generate summary
            summary = self.generate_summary()
            
            return {
                "template_path": template_path,
                "styles": self.styles,
                "numbering_definitions": numbering_definitions,
                "summary": summary
            }
            
        except Exception as e:
            print(f"Error analyzing template styles: {e}")
            return {"error": str(e)}
    
    def extract_numbering_definitions(self, template_path):
        """Extract numbering definitions from template's numbering.xml"""
        numbering_definitions = {}
        
        try:
            with zipfile.ZipFile(template_path) as zf:
                if "word/numbering.xml" in zf.namelist():
                    num_xml = zf.read("word/numbering.xml")
                    root = ET.fromstring(num_xml)
                    
                    # Extract abstract numbering definitions
                    for abstract_num in root.findall(".//w:abstractNum", self.namespace):
                        abstract_num_id = abstract_num.get(f"{{{self.namespace['w']}}}abstractNumId")
                        numbering_definitions[abstract_num_id] = {
                            "levels": {},
                            "bwa_label": None
                        }
                        
                        # Extract all levels
                        for lvl in abstract_num.findall("w:lvl", self.namespace):
                            ilvl = lvl.get(f"{{{self.namespace['w']}}}ilvl")
                            level_data = self.extract_level_data(lvl)
                            numbering_definitions[abstract_num_id]["levels"][ilvl] = level_data
                    
                    # Extract num mappings
                    for num_elem in root.findall(".//w:num", self.namespace):
                        num_id = num_elem.get(f"{{{self.namespace['w']}}}numId")
                        abstract_num_ref = num_elem.find("w:abstractNumId", self.namespace)
                        if abstract_num_ref is not None:
                            abstract_num_id = abstract_num_ref.get(f"{{{self.namespace['w']}}}val")
                            numbering_definitions[f"num_{num_id}"] = {
                                "abstract_num_id": abstract_num_id
                            }
                            
        except Exception as e:
            print(f"Error extracting numbering definitions: {e}")
        
        return numbering_definitions
    
    def extract_level_data(self, level_element):
        """Extract detailed data from a level element"""
        level_data = {
            "ilvl": level_element.get(f"{{{self.namespace['w']}}}ilvl"),
            "lvlText": None,
            "numFmt": None,
            "start": None,
            "suff": None,
            "lvlJc": None,
            "pStyle": None,
            "lvlPicBulletId": None,
            "legacy": None,
            "lvlRestart": None,
            "isLgl": None,
            "startOverride": None,
            "lvlOverride": None,
            "pPr": {},
            "rPr": {}
        }
        
        # Extract lvlText (the pattern like "%1.0", "%1.%2")
        lvl_text_elem = level_element.find("w:lvlText", self.namespace)
        if lvl_text_elem is not None:
            level_data["lvlText"] = lvl_text_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract numFmt (decimal, lowerLetter, upperLetter, etc.)
        num_fmt_elem = level_element.find("w:numFmt", self.namespace)
        if num_fmt_elem is not None:
            level_data["numFmt"] = num_fmt_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract start value
        start_elem = level_element.find("w:start", self.namespace)
        if start_elem is not None:
            level_data["start"] = start_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract suffix (tab, space, nothing)
        suff_elem = level_element.find("w:suff", self.namespace)
        if suff_elem is not None:
            level_data["suff"] = suff_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract lvlJc (justification: left, center, right)
        lvl_jc_elem = level_element.find("w:lvlJc", self.namespace)
        if lvl_jc_elem is not None:
            level_data["lvlJc"] = lvl_jc_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract pStyle (paragraph style)
        p_style_elem = level_element.find("w:pStyle", self.namespace)
        if p_style_elem is not None:
            level_data["pStyle"] = p_style_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract paragraph properties (pPr)
        p_pr_elem = level_element.find("w:pPr", self.namespace)
        if p_pr_elem is not None:
            level_data["pPr"] = self.extract_paragraph_properties(p_pr_elem)
        
        # Extract run properties (rPr) - font info
        r_pr_elem = level_element.find("w:rPr", self.namespace)
        if r_pr_elem is not None:
            level_data["rPr"] = self.extract_run_properties(r_pr_elem)
        
        return level_data
    
    def extract_paragraph_properties(self, p_pr_elem):
        """Extract paragraph properties from element"""
        p_pr = {}
        
        # Indentation
        indent_elem = p_pr_elem.find("w:ind", self.namespace)
        if indent_elem is not None:
            p_pr["indent"] = {
                "left": indent_elem.get(f"{{{self.namespace['w']}}}left"),
                "right": indent_elem.get(f"{{{self.namespace['w']}}}right"),
                "hanging": indent_elem.get(f"{{{self.namespace['w']}}}hanging"),
                "firstLine": indent_elem.get(f"{{{self.namespace['w']}}}firstLine")
            }
        
        # Spacing
        spacing_elem = p_pr_elem.find("w:spacing", self.namespace)
        if spacing_elem is not None:
            p_pr["spacing"] = {
                "before": spacing_elem.get(f"{{{self.namespace['w']}}}before"),
                "after": spacing_elem.get(f"{{{self.namespace['w']}}}after"),
                "line": spacing_elem.get(f"{{{self.namespace['w']}}}line"),
                "lineRule": spacing_elem.get(f"{{{self.namespace['w']}}}lineRule")
            }
        
        return p_pr
    
    def extract_run_properties(self, r_pr_elem):
        """Extract run properties from element"""
        r_pr = {}
        
        # Font family
        r_fonts_elem = r_pr_elem.find("w:rFonts", self.namespace)
        if r_fonts_elem is not None:
            r_pr["rFonts"] = {
                "ascii": r_fonts_elem.get(f"{{{self.namespace['w']}}}ascii"),
                "hAnsi": r_fonts_elem.get(f"{{{self.namespace['w']}}}hAnsi"),
                "eastAsia": r_fonts_elem.get(f"{{{self.namespace['w']}}}eastAsia"),
                "cs": r_fonts_elem.get(f"{{{self.namespace['w']}}}cs")
            }
        
        # Font size
        sz_elem = r_pr_elem.find("w:sz", self.namespace)
        if sz_elem is not None:
            r_pr["sz"] = sz_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Bold
        b_elem = r_pr_elem.find("w:b", self.namespace)
        if b_elem is not None:
            r_pr["bold"] = b_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Italic
        i_elem = r_pr_elem.find("w:i", self.namespace)
        if i_elem is not None:
            r_pr["italic"] = i_elem.get(f"{{{self.namespace['w']}}}val")
        
        return r_pr
    
    def extract_style_info(self, style, numbering_definitions):
        """Extract detailed information from a style"""
        try:
            style_info = {
                "name": style.name,
                "type": style.type,
                "base_style": None,
                "next_style": None,
                "is_bwa_style": "bwa" in style.name.lower(),
                "bwa_level_name": None
            }
            
            # Get base style and next style
            if hasattr(style, 'base_style') and style.base_style:
                style_info["base_style"] = style.base_style.name
            if hasattr(style, 'next_style') and style.next_style:
                style_info["next_style"] = style.next_style.name
            
            if style_info["is_bwa_style"]:
                style_info["bwa_level_name"] = style.name
            
            # Extract paragraph formatting
            if hasattr(style, 'paragraph_format'):
                pf = style.paragraph_format
                
                # Alignment
                if pf.alignment:
                    alignment_map = {
                        WD_ALIGN_PARAGRAPH.LEFT: "left",
                        WD_ALIGN_PARAGRAPH.CENTER: "center", 
                        WD_ALIGN_PARAGRAPH.RIGHT: "right",
                        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify"
                    }
                    style_info["alignment"] = alignment_map.get(pf.alignment, str(pf.alignment))
                
                # Indentation (convert to points)
                if pf.left_indent:
                    style_info["left_indent"] = pf.left_indent.pt
                if pf.right_indent:
                    style_info["right_indent"] = pf.right_indent.pt
                if pf.first_line_indent:
                    style_info["first_line_indent"] = pf.first_line_indent.pt
                
                # Spacing
                if pf.space_before:
                    style_info["space_before"] = pf.space_before.pt
                if pf.space_after:
                    style_info["space_after"] = pf.space_after.pt
                if pf.line_spacing:
                    style_info["line_spacing"] = pf.line_spacing
                
                # Line spacing rule
                if hasattr(pf, 'line_spacing_rule'):
                    style_info["line_spacing_rule"] = str(pf.line_spacing_rule)
                
                # Other paragraph properties
                if hasattr(pf, 'keep_with_next'):
                    style_info["keep_with_next"] = pf.keep_with_next
                if hasattr(pf, 'keep_lines_together'):
                    style_info["keep_lines_together"] = pf.keep_lines_together
                if hasattr(pf, 'page_break_before'):
                    style_info["page_break_before"] = pf.page_break_before
                if hasattr(pf, 'widow_control'):
                    style_info["widow_control"] = pf.widow_control
            
            # Extract character formatting
            if hasattr(style, 'font'):
                font = style.font
                
                if font.name:
                    style_info["font_name"] = font.name
                if font.size:
                    style_info["font_size"] = font.size.pt
                if font.bold is not None:
                    style_info["font_bold"] = font.bold
                if font.italic is not None:
                    style_info["font_italic"] = font.italic
                if font.underline:
                    style_info["font_underline"] = str(font.underline)
                if font.color.rgb:
                    style_info["font_color"] = str(font.color.rgb)
                if font.strike is not None:
                    style_info["font_strike"] = font.strike
                if font.small_caps is not None:
                    style_info["font_small_caps"] = font.small_caps
                if font.all_caps is not None:
                    style_info["font_all_caps"] = font.all_caps
            
            # Extract numbering information
            if hasattr(style, '_element'):
                style_element = style._element
                if style_element is not None:
                    pPr = style_element.find(qn('w:pPr'))
                    if pPr is not None:
                        numPr = pPr.find(qn('w:numPr'))
                        if numPr is not None:
                            numId = numPr.find(qn('w:numId'))
                            if numId is not None:
                                style_info["numbering_id"] = numId.get(qn('w:val'))
                            ilvl = numPr.find(qn('w:ilvl'))
                            if ilvl is not None:
                                style_info["numbering_level"] = int(ilvl.get(qn('w:val')))
            
            # Extract level list properties from numbering definitions
            if style_info.get("numbering_id"):
                self.extract_level_list_properties(style_info, numbering_definitions)
            
            return style_info
            
        except Exception as e:
            print(f"Error extracting style info for {style.name}: {e}")
            return None
    
    def extract_level_list_properties(self, style_info, numbering_definitions):
        """Extract level list position values from numbering definitions"""
        try:
            numbering_id = style_info.get("numbering_id")
            numbering_level = style_info.get("numbering_level", 0)
            
            # Find the numbering definition
            num_key = f"num_{numbering_id}"
            if num_key in numbering_definitions:
                abstract_num_id = numbering_definitions[num_key].get("abstract_num_id")
                if abstract_num_id in numbering_definitions:
                    abstract_info = numbering_definitions[abstract_num_id]
                    level_str = str(numbering_level)
                    
                    if level_str in abstract_info.get("levels", {}):
                        level_info = abstract_info["levels"][level_str]
                        
                        # Extract level list position values
                        style_info["number_alignment"] = level_info.get("lvlJc")  # left, center, right
                        style_info["follow_number_with"] = level_info.get("suff")  # tab, space, nothing
                        
                        # Extract position values from paragraph properties
                        p_pr = level_info.get("pPr", {})
                        if "indent" in p_pr:
                            indent = p_pr["indent"]
                            if indent.get("left"):
                                style_info["aligned_at"] = float(indent["left"]) / 20.0  # Convert twips to points
                            if indent.get("firstLine"):
                                style_info["text_indent_at"] = float(indent["firstLine"]) / 20.0
                        
                        # Extract tab stop information
                        # This would need to be extracted from the tab stops in the level definition
                        # For now, we'll note that this information is available in the XML
                        style_info["has_tab_stops"] = "tab_stops_available_in_xml"
                        
        except Exception as e:
            print(f"Error extracting level list properties: {e}")
    
    def generate_summary(self):
        """Generate summary statistics"""
        total_styles = len(self.styles)
        paragraph_styles = sum(1 for s in self.styles.values() if s.get("type") == 1)
        character_styles = sum(1 for s in self.styles.values() if s.get("type") == 2)
        table_styles = sum(1 for s in self.styles.values() if s.get("type") == 3)
        bwa_styles = sum(1 for s in self.styles.values() if s.get("is_bwa_style"))
        
        # Count styles with specific properties
        styles_with_numbering = sum(1 for s in self.styles.values() if s.get("numbering_id"))
        styles_with_spacing = sum(1 for s in self.styles.values() if s.get("space_before") or s.get("space_after"))
        styles_with_indentation = sum(1 for s in self.styles.values() if s.get("left_indent") or s.get("right_indent"))
        
        return {
            "total_styles": total_styles,
            "paragraph_styles": paragraph_styles,
            "character_styles": character_styles,
            "table_styles": table_styles,
            "bwa_styles": bwa_styles,
            "styles_with_numbering": styles_with_numbering,
            "styles_with_spacing": styles_with_spacing,
            "styles_with_indentation": styles_with_indentation
        }
    
    def save_analysis_to_json(self, analysis, output_path):
        """Save style analysis to JSON file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(analysis, f, indent=2, ensure_ascii=False)
        print(f"Style analysis saved to: {output_path}")
    
    def print_style_details(self, style_name=None):
        """Print detailed information about styles"""
        if style_name:
            if style_name in self.styles:
                self._print_single_style(style_name, self.styles[style_name])
            else:
                print(f"Style '{style_name}' not found")
        else:
            # Print all BWA styles first
            bwa_styles = {name: info for name, info in self.styles.items() if info.get("is_bwa_style")}
            if bwa_styles:
                print("\n=== BWA STYLES ===")
                for name, info in bwa_styles.items():
                    self._print_single_style(name, info)
            
            # Print other paragraph styles
            other_paragraph_styles = {name: info for name, info in self.styles.items() 
                                    if info.get("type") == 1 and not info.get("is_bwa_style")}
            if other_paragraph_styles:
                print("\n=== OTHER PARAGRAPH STYLES ===")
                for name, info in other_paragraph_styles.items():
                    self._print_single_style(name, info)
    
    def _print_single_style(self, name, info):
        """Print information about a single style"""
        print(f"\nStyle: {name}")
        print(f"  Type: {info.get('type')}")
        if info.get("base_style"):
            print(f"  Base Style: {info['base_style']}")
        if info.get("next_style"):
            print(f"  Next Style: {info['next_style']}")
        
        # Paragraph formatting
        if any([info.get("alignment"), info.get("left_indent"), info.get("right_indent"), 
                info.get("first_line_indent"), info.get("space_before"), info.get("space_after"), 
                info.get("line_spacing")]):
            print("  Paragraph Formatting:")
            if info.get("alignment"):
                print(f"    Alignment: {info['alignment']}")
            if info.get("left_indent"):
                print(f"    Left Indent: {info['left_indent']}pt")
            if info.get("right_indent"):
                print(f"    Right Indent: {info['right_indent']}pt")
            if info.get("first_line_indent"):
                print(f"    First Line Indent: {info['first_line_indent']}pt")
            if info.get("space_before"):
                print(f"    Space Before: {info['space_before']}pt")
            if info.get("space_after"):
                print(f"    Space After: {info['space_after']}pt")
            if info.get("line_spacing"):
                print(f"    Line Spacing: {info['line_spacing']}")
            if info.get("line_spacing_rule"):
                print(f"    Line Spacing Rule: {info['line_spacing_rule']}")
        
        # Character formatting
        if any([info.get("font_name"), info.get("font_size"), info.get("font_bold"), 
                info.get("font_italic"), info.get("font_underline"), info.get("font_color")]):
            print("  Character Formatting:")
            if info.get("font_name"):
                print(f"    Font: {info['font_name']}")
            if info.get("font_size"):
                print(f"    Size: {info['font_size']}pt")
            if info.get("font_bold") is not None:
                print(f"    Bold: {info['font_bold']}")
            if info.get("font_italic") is not None:
                print(f"    Italic: {info['font_italic']}")
            if info.get("font_underline"):
                print(f"    Underline: {info['font_underline']}")
            if info.get("font_color"):
                print(f"    Color: {info['font_color']}")
        
        # Numbering and level list properties
        if info.get("numbering_id") is not None:
            print("  Numbering:")
            print(f"    Numbering ID: {info['numbering_id']}")
            if info.get("numbering_level") is not None:
                print(f"    Level: {info['numbering_level']}")
            if info.get("number_alignment"):
                print(f"    Number Alignment: {info['number_alignment']}")
            if info.get("follow_number_with"):
                print(f"    Follow Number With: {info['follow_number_with']}")
            if info.get("aligned_at"):
                print(f"    Aligned At: {info['aligned_at']}pt")
            if info.get("text_indent_at"):
                print(f"    Text Indent At: {info['text_indent_at']}pt")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python simple_style_analyzer.py <template_file.docx>")
        print("Example: python simple_style_analyzer.py '../templates/test_template_cleaned.docx'")
        sys.exit(1)
    
    template_file = sys.argv[1]
    
    if not os.path.exists(template_file):
        print(f"Error: File '{template_file}' not found.")
        sys.exit(1)
    
    try:
        analyzer = SimpleStyleAnalyzer()
        analysis = analyzer.analyze_template(template_file)
        
        if "error" not in analysis:
            # Print style details
            analyzer.print_style_details()
            
            # Save to JSON
            base_name = os.path.splitext(os.path.basename(template_file))[0]
            json_file = f"{base_name}_style_analysis.json"
            analyzer.save_analysis_to_json(analysis, json_file)
            
            # Print summary
            summary = analysis["summary"]
            print(f"\n=== SUMMARY ===")
            print(f"Total styles: {summary['total_styles']}")
            print(f"Paragraph styles: {summary['paragraph_styles']}")
            print(f"Character styles: {summary['character_styles']}")
            print(f"Table styles: {summary['table_styles']}")
            print(f"BWA styles: {summary['bwa_styles']}")
            print(f"Styles with numbering: {summary['styles_with_numbering']}")
            print(f"Styles with spacing: {summary['styles_with_spacing']}")
            print(f"Styles with indentation: {summary['styles_with_indentation']}")
            
        else:
            print(f"Error: {analysis['error']}")
            
    except Exception as e:
        print(f"Error processing template: {e}")

if __name__ == "__main__":
    main() 
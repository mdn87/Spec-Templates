#!/usr/bin/env python3
"""
Template Style Analyzer

This script analyzes paragraph styles in a Word document template to extract
detailed formatting information including spacing, indentation, and other properties.

Usage:
    python analyze_template_styles.py <template_file.docx>
"""

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import json
import os
import sys
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, asdict

@dataclass
class StyleInfo:
    """Represents detailed style information"""
    name: str
    type: str  # 'paragraph', 'character', 'table', etc.
    base_style: Optional[str] = None
    next_style: Optional[str] = None
    # Paragraph formatting
    alignment: Optional[str] = None
    left_indent: Optional[float] = None
    right_indent: Optional[float] = None
    first_line_indent: Optional[float] = None
    space_before: Optional[float] = None
    space_after: Optional[float] = None
    line_spacing: Optional[float] = None
    line_spacing_rule: Optional[str] = None
    keep_with_next: Optional[bool] = None
    keep_lines_together: Optional[bool] = None
    page_break_before: Optional[bool] = None
    widow_control: Optional[bool] = None
    # Character formatting
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_bold: Optional[bool] = None
    font_italic: Optional[bool] = None
    font_underline: Optional[str] = None
    font_color: Optional[str] = None
    font_strike: Optional[bool] = None
    font_small_caps: Optional[bool] = None
    font_all_caps: Optional[bool] = None
    # Numbering and level list properties
    numbering_id: Optional[str] = None
    numbering_level: Optional[int] = None
    # Level list position values
    number_alignment: Optional[str] = None  # left, center, right
    aligned_at: Optional[float] = None  # position in points
    text_indent_at: Optional[float] = None  # position in points
    follow_number_with: Optional[str] = None  # tab, space, nothing
    add_tab_stop_at: Optional[float] = None  # position in points
    # Additional properties
    is_bwa_style: bool = False
    bwa_level_name: Optional[str] = None

class TemplateStyleAnalyzer:
    """Analyzes styles in a Word document template"""
    
    def __init__(self):
        self.styles: Dict[str, StyleInfo] = {}
    
    def analyze_template(self, template_path: str) -> Dict[str, Any]:
        """Analyze all styles in the template document"""
        try:
            print(f"Analyzing styles in template: {template_path}")
            doc = Document(template_path)
            
            # Analyze all styles
            for style in doc.styles:
                style_info = self.extract_style_info(style, doc)
                if style_info and style.name is not None:
                    self.styles[style.name] = style_info

                # Generate summary
                summary = self.generate_summary()
            return {
                "template_path": template_path,
                "styles": {name: asdict(info) for name, info in self.styles.items()},
                "summary": summary
            }
            
        except Exception as e:
            print(f"Error analyzing template styles: {e}")
            return {"error": str(e)}
    
    def extract_style_info(self, style, doc) -> StyleInfo:
        """Extract detailed information from a style"""
        try:
            style_info = StyleInfo(
                name=style.name,
                type=style.type
            )
            
            # Get base style and next style
            if hasattr(style, 'base_style') and style.base_style:
                style_info.base_style = style.base_style.name
            if hasattr(style, 'next_style') and style.next_style:
                style_info.next_style = style.next_style.name
            
            # Check if this is a BWA style
            style_info.is_bwa_style = "bwa" in style.name.lower()
            if style_info.is_bwa_style:
                style_info.bwa_level_name = style.name
            
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
                    style_info.alignment = alignment_map.get(pf.alignment, str(pf.alignment))
                
                # Indentation (convert to points)
                if pf.left_indent:
                    style_info.left_indent = pf.left_indent.pt
                if pf.right_indent:
                    style_info.right_indent = pf.right_indent.pt
                if pf.first_line_indent:
                    style_info.first_line_indent = pf.first_line_indent.pt
                # Note: hanging_indent is not a direct attribute in python-docx
                # It's calculated as: left_indent + first_line_indent (where first_line_indent is negative)
                
                # Spacing
                if pf.space_before:
                    style_info.space_before = pf.space_before.pt
                if pf.space_after:
                    style_info.space_after = pf.space_after.pt
                if pf.line_spacing:
                    style_info.line_spacing = pf.line_spacing
                
                # Line spacing rule
                if hasattr(pf, 'line_spacing_rule'):
                    style_info.line_spacing_rule = str(pf.line_spacing_rule)
                
                # Other paragraph properties
                if hasattr(pf, 'keep_with_next'):
                    style_info.keep_with_next = pf.keep_with_next
                if hasattr(pf, 'keep_lines_together'):
                    style_info.keep_lines_together = pf.keep_lines_together
                if hasattr(pf, 'page_break_before'):
                    style_info.page_break_before = pf.page_break_before
                if hasattr(pf, 'widow_control'):
                    style_info.widow_control = pf.widow_control
            
            # Extract character formatting
            if hasattr(style, 'font'):
                font = style.font
                
                if font.name:
                    style_info.font_name = font.name
                if font.size:
                    style_info.font_size = font.size.pt
                if font.bold is not None:
                    style_info.font_bold = font.bold
                if font.italic is not None:
                    style_info.font_italic = font.italic
                if font.underline:
                    style_info.font_underline = str(font.underline)
                if font.color.rgb:
                    style_info.font_color = str(font.color.rgb)
                if font.strike is not None:
                    style_info.font_strike = font.strike
                if font.small_caps is not None:
                    style_info.font_small_caps = font.small_caps
                if font.all_caps is not None:
                    style_info.font_all_caps = font.all_caps
            
            # Extract numbering information and level list properties
            if hasattr(style, '_element'):
                style_element = style._element
                if style_element is not None:
                    pPr = style_element.find(qn('w:pPr'))
                    if pPr is not None:
                        numPr = pPr.find(qn('w:numPr'))
                        if numPr is not None:
                            numId = numPr.find(qn('w:numId'))
                            if numId is not None:
                                style_info.numbering_id = numId.get(qn('w:val'))
                            ilvl = numPr.find(qn('w:ilvl'))
                            if ilvl is not None:
                                style_info.numbering_level = int(ilvl.get(qn('w:val')))
            
            # Extract level list position values from numbering definitions
            if style_info.numbering_id:
                self.extract_level_list_properties(style_info)
            
            return style_info
            
        except Exception as e:
            print(f"Error extracting style info for {style.name}: {e}")
            raise  # Re-raise the exception to avoid returning None and violating the return type
    
    def extract_level_list_properties(self, style_info: StyleInfo):
        """Extract level list position values from numbering definitions"""
        try:
            # This would need to access the numbering.xml to get level list properties
            # For now, we'll extract what we can from the style element
            # Note: StyleInfo does not have an _element attribute; use style_info.style if available
            style = getattr(style_info, 'style', None)
            if style is not None and hasattr(style, '_element'):
                style_element = style._element
                if style_element is not None:
                    pPr = style_element.find(qn('w:pPr'))
                    if pPr is not None:
                        numPr = pPr.find(qn('w:numPr'))
                        if numPr is not None:
                            numId = numPr.find(qn('w:numId'))
                            if numId is not None:
                                style_info.numbering_id = numId.get(qn('w:val'))
                            ilvl = numPr.find(qn('w:ilvl'))
                            if ilvl is not None:
                                try:
                                    style_info.numbering_level = int(ilvl.get(qn('w:val')))
                                except (TypeError, ValueError):
                                    style_info.numbering_level = None
                    # Could extract more properties here as needed
        except Exception as e:
            print(f"Error extracting level list properties: {e}")
    
    def generate_summary(self) -> Dict[str, Any]:
        """Generate summary statistics"""
        total_styles = len(self.styles)
        paragraph_styles = sum(1 for s in self.styles.values() if s.type == 1)  # WD_STYLE_TYPE.PARAGRAPH
        character_styles = sum(1 for s in self.styles.values() if s.type == 2)  # WD_STYLE_TYPE.CHARACTER
        table_styles = sum(1 for s in self.styles.values() if s.type == 3)  # WD_STYLE_TYPE.TABLE
        bwa_styles = sum(1 for s in self.styles.values() if s.is_bwa_style)
        
        # Count styles with specific properties
        styles_with_numbering = sum(1 for s in self.styles.values() if s.numbering_id is not None)
        styles_with_spacing = sum(1 for s in self.styles.values() if s.space_before is not None or s.space_after is not None)
        styles_with_indentation = sum(1 for s in self.styles.values() if s.left_indent is not None or s.right_indent is not None)
        
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
    
    def save_analysis_to_json(self, analysis: Dict[str, Any], output_path: str):
        """Save style analysis to JSON file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(analysis, f, indent=2, ensure_ascii=False)
        print(f"Style analysis saved to: {output_path}")

    def print_style_details(self, style_name: str = ""):
        """Print detailed information about styles"""
        if style_name:
            if style_name in self.styles:
                self._print_single_style(style_name, self.styles[style_name])
            else:
                print(f"Style '{style_name}' not found")
        else:
            # Print all BWA styles first
            bwa_styles = {name: info for name, info in self.styles.items() if info.is_bwa_style}
            if bwa_styles:
                print("\n=== BWA STYLES ===")
                for name, info in bwa_styles.items():
                    self._print_single_style(name, info)
            
            # Print other paragraph styles
            other_paragraph_styles = {name: info for name, info in self.styles.items() 
                                    if info.type == 1 and not info.is_bwa_style}
            if other_paragraph_styles:
                print("\n=== OTHER PARAGRAPH STYLES ===")
                for name, info in other_paragraph_styles.items():
                    self._print_single_style(name, info)
    
    def _print_single_style(self, name: str, info: StyleInfo):
        """Print information about a single style"""
        print(f"\nStyle: {name}")
        print(f"  Type: {info.type}")
        if info.base_style:
            print(f"  Base Style: {info.base_style}")
        if info.next_style:
            print(f"  Next Style: {info.next_style}")
        
        # Paragraph formatting
        if any([
            info.alignment,
            info.left_indent,
            info.right_indent,
            info.first_line_indent,
            info.space_before,
            info.space_after,
            info.line_spacing
        ]):
            print("  Paragraph Formatting:")
            if info.alignment:
                print(f"    Alignment: {info.alignment}")
            if info.left_indent is not None:
                print(f"    Left Indent: {info.left_indent}pt")
            if info.right_indent is not None:
                print(f"    Right Indent: {info.right_indent}pt")
            if info.first_line_indent is not None:
                print(f"    First Line Indent: {info.first_line_indent}pt")
            # Calculate hanging indent if applicable
            if (
                info.left_indent is not None
                and info.first_line_indent is not None
                and info.first_line_indent < 0
            ):
                hanging = info.left_indent + info.first_line_indent
                print(f"    Hanging Indent: {hanging}pt")
            if info.space_before is not None:
                print(f"    Space Before: {info.space_before}pt")
            if info.space_after is not None:
                print(f"    Space After: {info.space_after}pt")
            if info.line_spacing is not None:
                print(f"    Line Spacing: {info.line_spacing}")
            if info.line_spacing_rule:
                print(f"    Line Spacing Rule: {info.line_spacing_rule}")
                # Calculate hanging indent if applicable
                if info.left_indent is not None and info.first_line_indent is not None and info.first_line_indent < 0:
                    hanging = info.left_indent + info.first_line_indent
                    print(f"    Hanging Indent: {hanging}pt")
            if info.space_before is not None:
                print(f"    Space Before: {info.space_before}pt")
            if info.space_after is not None:
                print(f"    Space After: {info.space_after}pt")
            if info.line_spacing is not None:
                print(f"    Line Spacing: {info.line_spacing}")
            if info.line_spacing_rule:
                print(f"    Line Spacing Rule: {info.line_spacing_rule}")
        
        # Character formatting
        if any([info.font_name, info.font_size, info.font_bold, info.font_italic, 
                info.font_underline, info.font_color]):
            print("  Character Formatting:")
            if info.font_name:
                print(f"    Font: {info.font_name}")
            if info.font_size:
                print(f"    Size: {info.font_size}pt")
            if info.font_bold is not None:
                print(f"    Bold: {info.font_bold}")
            if info.font_italic is not None:
                print(f"    Italic: {info.font_italic}")
            if info.font_underline:
                print(f"    Underline: {info.font_underline}")
            if info.font_color:
                print(f"    Color: {info.font_color}")
        
        # Numbering
        if info.numbering_id is not None:
            print("  Numbering:")
            print(f"    Numbering ID: {info.numbering_id}")
            if info.numbering_level is not None:
                print(f"    Level: {info.numbering_level}")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python analyze_template_styles.py <template_file.docx>")
        print("Example: python analyze_template_styles.py '../templates/test_template_cleaned.docx'")
        sys.exit(1)
    
    template_file = sys.argv[1]
    
    if not os.path.exists(template_file):
        print(f"Error: File '{template_file}' not found.")
        sys.exit(1)
    
    try:
        analyzer = TemplateStyleAnalyzer()
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
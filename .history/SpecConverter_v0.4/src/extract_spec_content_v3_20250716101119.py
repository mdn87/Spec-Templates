#!/usr/bin/env python3
"""
Specification Content Extractor - Version 3

This script extracts multi-level list content from Word documents (.docx) and converts it to JSON format.
It combines the best of both approaches:
- JSON output structure (working well)
- Header/footer/margin extraction from rip scripts
- Comments extraction
- BWA list level detection and mapping
- Template-based validation

Features:
- Extracts section headers, titles, parts, subsections, items, and lists
- Handles both numbered and unnumbered structures
- Validates numbering sequences and reports errors
- Extracts header, footer, margin, and comment information
- Maps content to BWA list levels from template
- Outputs comprehensive JSON with level information
- Generates detailed error reports

Usage:
    python extract_spec_content_v3.py <docx_file> [output_dir] [template_file]

Example:
    python extract_spec_content_v3.py "SECTION 26 05 00.docx"
"""

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from datetime import datetime

# Import the header/footer extractor module
from header_footer_extractor import HeaderFooterExtractor

# Import the template list detector module
from template_list_detector import TemplateListDetector

@dataclass
class ExtractionError:
    """Represents an error found during content extraction"""
    line_number: int
    error_type: str
    message: str
    context: str
    expected: Optional[str] = None
    found: Optional[str] = None

@dataclass
class ContentBlock:
    """Represents a content block with level information and styling"""
    text: str
    level_type: str
    number: Optional[str] = None
    content: str = ""
    level_number: Optional[int] = None
    bwa_level_name: Optional[str] = None
    numbering_id: Optional[str] = None
    numbering_level: Optional[int] = None
    style_name: Optional[str] = None
    # Styling information
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_bold: Optional[bool] = None
    font_italic: Optional[bool] = None
    font_underline: Optional[str] = None
    font_color: Optional[str] = None
    paragraph_alignment: Optional[str] = None
    paragraph_indent_left: Optional[float] = None
    paragraph_indent_right: Optional[float] = None
    paragraph_indent_first_line: Optional[float] = None
    paragraph_indent_hanging: Optional[float] = None
    paragraph_spacing_before: Optional[float] = None
    paragraph_spacing_after: Optional[float] = None
    paragraph_line_spacing: Optional[float] = None

class SpecContentExtractorV3:
    """Extracts specification content with comprehensive metadata"""
    
    def __init__(self, template_path: Optional[str] = None):
        self.errors: List[ExtractionError] = []
        self.line_count = 0
        self.section_header_found = False
        self.section_title_found = False
        
        # Document structure
        self.section_number = ""
        self.section_title = ""
        self.end_of_section = ""
        
        # Template analysis
        self.template_path = template_path
        self.bwa_list_levels = {}
        self.template_numbering = {}
        
        # Content blocks
        self.content_blocks: List[ContentBlock] = []
        
        # List numbering tracking
        self.list_counters = {}  # {(numId, ilvl): current_number}
        self.list_fixes = []  # Track numbering fixes for reporting
        
        # Regex patterns
        self.section_pattern = re.compile(r'^SECTION\s+(.+)$', re.IGNORECASE)
        self.end_section_pattern = re.compile(r'^END\s+OF\s+SECTION\s*(.+)?$', re.IGNORECASE)
        self.part_pattern = re.compile(r'^(\d+\.0)\s+(.+)$')
        self.subsection_pattern = re.compile(r'^(\d+\.\d{2})\s+(.+)$')
        self.subsection_alt_pattern = re.compile(r'^(\d+\.\d)\s+(.+)$')
        self.item_pattern = re.compile(r'^([A-Z])\.\s+(.+)$')
        self.list_pattern = re.compile(r'^(\d+)\.\s+(.+)$')
        self.sub_list_pattern = re.compile(r'^([a-z])\.\s+(.+)$')
        
        # Load template if provided
        if template_path:
            self.load_template_analysis(template_path)
    
    def load_template_analysis(self, template_path: str):
        """Load and analyze template structure using the template list detector module"""
        try:
            print(f"Loading template analysis from: {template_path}")
            
            # Use the template list detector module
            detector = TemplateListDetector()
            analysis = detector.analyze_template(template_path)
            
            # Store the analysis results
            self.template_numbering = analysis.numbering_definitions
            self.bwa_list_levels = analysis.bwa_list_levels
            self.template_analysis = analysis
            
            print(f"Template loaded: {len(self.bwa_list_levels)} BWA list levels found")
            
        except Exception as e:
            print(f"Warning: Could not load template analysis: {e}")
            self.template_analysis = None
    
    def get_paragraph_level(self, paragraph) -> Optional[int]:
        """Get the list level of a paragraph"""
        try:
            pPr = paragraph._p.pPr
            if pPr is not None and pPr.numPr is not None:
                if pPr.numPr.ilvl is not None:
                    return pPr.numPr.ilvl.val
        except:
            pass
        return None
    
    def get_paragraph_numbering_id(self, paragraph) -> Optional[str]:
        """Get the numbering ID of a paragraph"""
        try:
            pPr = paragraph._p.pPr
            if pPr is not None and pPr.numPr is not None:
                if pPr.numPr.numId is not None:
                    return str(pPr.numPr.numId.val)
        except:
            pass
        return None
    
    def extract_paragraph_styling(self, paragraph) -> Dict[str, Any]:
        """Extract styling information from a paragraph"""
        styling = {}
        
        try:
            # Paragraph properties
            pPr = paragraph._p.pPr
            if pPr is not None:
                # Alignment
                jc_elem = pPr.find(qn('w:jc'))
                if jc_elem is not None:
                    styling['paragraph_alignment'] = jc_elem.get(qn('w:val'))
                
                # Indentation
                indent_elem = pPr.find(qn('w:ind'))
                if indent_elem is not None:
                    left = indent_elem.get(qn('w:left'))
                    if left is not None:
                        styling['paragraph_indent_left'] = float(left) / 20.0  # Convert twips to points
                    
                    right = indent_elem.get(qn('w:right'))
                    if right is not None:
                        styling['paragraph_indent_right'] = float(right) / 20.0
                    
                    first_line = indent_elem.get(qn('w:firstLine'))
                    if first_line is not None:
                        styling['paragraph_indent_first_line'] = float(first_line) / 20.0
                    
                    hanging = indent_elem.get(qn('w:hanging'))
                    if hanging is not None:
                        styling['paragraph_indent_hanging'] = float(hanging) / 20.0
                
                # Spacing
                spacing_elem = pPr.find(qn('w:spacing'))
                if spacing_elem is not None:
                    before = spacing_elem.get(qn('w:before'))
                    if before is not None:
                        styling['paragraph_spacing_before'] = float(before) / 20.0
                    
                    after = spacing_elem.get(qn('w:after'))
                    if after is not None:
                        styling['paragraph_spacing_after'] = float(after) / 20.0
                    
                    line = spacing_elem.get(qn('w:line'))
                    if line is not None:
                        styling['paragraph_line_spacing'] = float(line) / 240.0  # Convert to line spacing ratio
            
            # Run properties (font information)
            # Get the most common font properties from all runs
            font_properties = self.extract_run_styling(paragraph.runs)
            styling.update(font_properties)
            
        except Exception as e:
            print(f"Warning: Could not extract paragraph styling: {e}")
        
        return styling
    
    def extract_run_styling(self, runs) -> Dict[str, Any]:
        """Extract styling information from paragraph runs"""
        styling = {}
        
        if not runs:
            return styling
        
        try:
            # Collect font properties from all runs
            font_names = []
            font_sizes = []
            font_bolds = []
            font_italics = []
            font_underlines = []
            font_colors = []
            
            for run in runs:
                rPr = run._r.rPr
                if rPr is not None:
                    # Font family
                    r_fonts = rPr.find(qn('w:rFonts'))
                    if r_fonts is not None:
                        ascii_font = r_fonts.get(qn('w:ascii'))
                        if ascii_font:
                            font_names.append(ascii_font)
                    
                    # Font size
                    sz = rPr.find(qn('w:sz'))
                    if sz is not None:
                        size_val = sz.get(qn('w:val'))
                        if size_val:
                            font_sizes.append(float(size_val) / 2.0)  # Convert half-points to points
                    
                    # Bold
                    b = rPr.find(qn('w:b'))
                    if b is not None:
                        bold_val = b.get(qn('w:val'))
                        if bold_val is not None:
                            font_bolds.append(bold_val == 'true' or bold_val == '1')
                        else:
                            font_bolds.append(True)  # Default to True if present but no value
                    
                    # Italic
                    i = rPr.find(qn('w:i'))
                    if i is not None:
                        italic_val = i.get(qn('w:val'))
                        if italic_val is not None:
                            font_italics.append(italic_val == 'true' or italic_val == '1')
                        else:
                            font_italics.append(True)
                    
                    # Underline
                    u = rPr.find(qn('w:u'))
                    if u is not None:
                        underline_val = u.get(qn('w:val'))
                        if underline_val:
                            font_underlines.append(underline_val)
                    
                    # Color
                    color = rPr.find(qn('w:color'))
                    if color is not None:
                        color_val = color.get(qn('w:val'))
                        if color_val:
                            font_colors.append(color_val)
            
            # Use the most common font properties
            if font_names:
                styling['font_name'] = max(set(font_names), key=font_names.count)
            if font_sizes:
                styling['font_size'] = max(set(font_sizes), key=font_sizes.count)
            if font_bolds:
                styling['font_bold'] = max(set(font_bolds), key=font_bolds.count)
            if font_italics:
                styling['font_italic'] = max(set(font_italics), key=font_italics.count)
            if font_underlines:
                styling['font_underline'] = max(set(font_underlines), key=font_underlines.count)
            if font_colors:
                styling['font_color'] = max(set(font_colors), key=font_colors.count)
            
        except Exception as e:
            print(f"Warning: Could not extract run styling: {e}")
        
        return styling
    
    def extract_header_footer_margins(self, docx_path: str) -> Dict[str, Any]:
        """Extract header, footer, and margin information using the header/footer extractor module"""
        try:
            extractor = HeaderFooterExtractor()
            return extractor.extract_header_footer_margins(docx_path)
        except Exception as e:
            print(f"Error extracting header/footer/margins: {e}")
            return {"header": {}, "footer": {}, "margins": {}}
    
    def extract_comments(self, docx_path: str) -> List[Dict[str, Any]]:
        """Extract comments from document using the header/footer extractor module"""
        try:
            extractor = HeaderFooterExtractor()
            return extractor.extract_comments(docx_path)
        except Exception as e:
            print(f"Error extracting comments: {e}")
            return []
    
    def classify_paragraph_level(self, text: str) -> Tuple[str, Optional[str], str]:
        """
        Classify a paragraph into its hierarchical level
        Returns: (level_type, number, content)
        """
        text = text.strip()
        if not text:
            return "empty", None, ""
        
        # Check for section header (must be the very first line)
        if text.upper().startswith("SECTION") and not self.section_header_found:
            match = self.section_pattern.match(text)
            if match:
                self.section_header_found = True
                return "section", match.group(1), ""
        elif text.upper().startswith("SECTION") and self.section_header_found:
            self.add_error("Structure Error", "Multiple section headers found", text)
            return "content", None, text
        
        # Check for section title (must be the second line after section header)
        if (self.section_header_found and 
            not self.section_title_found and
            len(text.strip()) > 0):
            self.section_title_found = True
            return "title", None, text
        
        # Check for end of section
        if text.upper().startswith("END OF SECTION"):
            match = self.end_section_pattern.match(text)
            if match:
                self.end_of_section = match.group(1).strip() if match.group(1) else ""
                return "end_of_section", None, self.end_of_section
        
        # Check for part level with numbering (1.0, 2.0, etc.)
        match = self.part_pattern.match(text)
        if match:
            return "part", match.group(1), match.group(2)
        
        # Check for part titles with various formats
        part_names = ["DESCRIPTION", "PRODUCTS", "EXECUTION", "GENERAL"]
        for part_name in part_names:
            match = re.match(rf'(?:PART\s*)?(\d+)\.0?\s*[-]?\s*{part_name}$', text.upper())
            if match:
                part_number = f"{match.group(1)}.0"
                return "part_title", part_number, part_name
            elif text.strip().upper() == part_name:
                part_number = f"{len([b for b in self.content_blocks if b.level_type == 'part']) + 1}.0"
                return "part_title", part_number, part_name
        
        # Check for subsection level with numbering (1.01, 1.02, etc.)
        match = self.subsection_pattern.match(text)
        if match:
            return "subsection", match.group(1), match.group(2)
        
        # Check for subsection level with alternative numbering (1.1, 1.2, etc.)
        match = self.subsection_alt_pattern.match(text)
        if match:
            return "subsection", match.group(1), match.group(2)
        
        # Check for item level (A., B., C., etc.) - MOVED BEFORE subsection titles
        match = self.item_pattern.match(text)
        if match:
            return "item", match.group(1), match.group(2)
        
        # Check for list level (1., 2., etc.)
        match = self.list_pattern.match(text)
        if match:
            return "list", match.group(1), match.group(2)
        
        # Check for sub-list level (a., b., etc.)
        match = self.sub_list_pattern.match(text)
        if match:
            return "sub_list", match.group(1), match.group(2)
        
        # Check for subsection titles without numbering - MOVED AFTER item patterns
        subsection_titles = [
            "SCOPE", "EXISTING CONDITIONS", "CODES AND REGULATIONS", "DEFINITIONS",
            "DRAWINGS AND SPECIFICATIONS", "SITE VISIT", "DEVIATIONS",
            "STANDARDS FOR MATERIALS AND WORKMANSHIP", "SHOP DRAWINGS AND SUBMITTAL",
            "RECORD (AS-BUILT) DRAWINGS AND MAINTENANCE MANUALS",
            "COORDINATION", "PROTECTION OF MATERIALS", "TESTS, DEMONSTRATION AND INSTRUCTIONS",
            "GUARANTEE"
        ]
        
        # Check for exact match first (more specific)
        if text.upper() in [title.upper() for title in subsection_titles]:
            return "subsection_title", None, text
        
        # Check for partial matches last (less specific)
        text_upper = text.upper().strip()
        for title in subsection_titles:
            if title.upper() in text_upper or text_upper in title.upper():
                return "subsection_title", None, title
        
        # If no pattern matches, it's regular content
        return "content", None, text
    
    def correct_level_type_based_on_numbering(self, level_type: str, numbering_id: Optional[str], 
                                            numbering_level: Optional[int], text: str) -> str:
        """Correct level type based on paragraph numbering information"""
        # If it was classified as content but has numbering, it's likely a list item
        if level_type == "content" and numbering_id is not None:
            # Check if this looks like a list item (short text, no obvious structure)
            text_clean = text.strip()
            if len(text_clean) < 100 and not text_clean.upper().startswith(("SECTION", "PART", "GENERAL", "SCOPE")):
                # Based on numbering level, determine the type
                if numbering_level == 0:
                    return "list"  # Top-level list items
                elif numbering_level == 1:
                    return "sub_list"  # Sub-list items
                else:
                    return "list"  # Default to list for any numbered content
        
        return level_type

    def extract_list_number(self, numbering_id: Optional[str], numbering_level: Optional[int], 
                          detected_number: Optional[str], text: str) -> Tuple[Optional[str], bool]:
        """
        Extract the correct list number from Word's numbering system
        Returns: (correct_number, was_fixed)
        """
        if numbering_id is None or numbering_level is None:
            return detected_number, False
        
        # Create key for tracking this specific list
        key = (numbering_id, numbering_level)
        
        # Initialize counter if this is a new list or level
        if key not in self.list_counters:
            self.list_counters[key] = 1
        else:
            self.list_counters[key] += 1
        
        correct_number = str(self.list_counters[key])
        
        # Check if we need to fix the detected number
        was_fixed = False
        if detected_number is not None and detected_number != correct_number:
            was_fixed = True
            self.list_fixes.append({
                "line_number": self.line_count,
                "text": text[:50] + "..." if len(text) > 50 else text,
                "detected_number": detected_number,
                "correct_number": correct_number,
                "numbering_id": numbering_id,
                "numbering_level": numbering_level
            })
        
        return correct_number, was_fixed

    def map_to_bwa_level(self, paragraph, level_type: str) -> Tuple[Optional[int], Optional[str]]:
        """Map paragraph to BWA list level based on template analysis"""
        try:
            # Standard mapping for level_number
            level_mapping = {
                "part": 0,
                "part_title": 0,
                "subsection": 1,
                "subsection_title": 1,
                "item": 2,
                "list": 3,
                "sub_list": 4
            }
            level_number = level_mapping.get(level_type)

            # Map to BWA style name for label
            level_type_to_bwa_mapping = {
                "part": "BWA-PART",
                "part_title": "BWA-PART",
                "subsection": "BWA-SUBSECTION",
                "subsection_title": "BWA-SUBSECTION",
                "item": "BWA-Item",
                "list": "BWA-List",
                "sub_list": "BWA-SubList"
            }
            bwa_level_name = None
            if level_type in level_type_to_bwa_mapping:
                bwa_style_name = level_type_to_bwa_mapping[level_type]
                if bwa_style_name in self.bwa_list_levels:
                    bwa_level_name = bwa_style_name
                else:
                    bwa_level_name = bwa_style_name  # Even if not in template, use the label

            return level_number, bwa_level_name
        except Exception as e:
            print(f"Error mapping to BWA level: {e}")
            return None, None
    
    def add_error(self, error_type: str, message: str, context: str = "", 
                  expected: Optional[str] = None, found: Optional[str] = None):
        """Add an error to the error list"""
        error = ExtractionError(
            line_number=self.line_count,
            error_type=error_type,
            message=message,
            context=context,
            expected=expected,
            found=found
        )
        self.errors.append(error)
    
    def extract_content(self, docx_path: str) -> Dict[str, Any]:
        """Extract content from a Word document"""
        try:
            doc = Document(docx_path)
            
            # Extract header, footer, margin, and comment information
            header_footer_data = self.extract_header_footer_margins(docx_path)
            comments = self.extract_comments(docx_path)
            
            # Extract all paragraphs
            paragraphs = []
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if text:  # Only include non-empty paragraphs
                    paragraphs.append((text, paragraph))
                    self.line_count += 1
            
            # Extract section header and title
            section_number, section_title = self.extract_section_header_and_title([p[0] for p in paragraphs])
            
            # Process each paragraph
            for text, paragraph in paragraphs:
                level_type, number, content = self.classify_paragraph_level(text)
                
                # Skip empty content
                if level_type == "empty":
                    continue
                
                # Get numbering information
                numbering_id = self.get_paragraph_numbering_id(paragraph)
                numbering_level = self.get_paragraph_level(paragraph)
                style_name = paragraph.style.name if paragraph.style else None
                
                # Extract styling information
                styling = self.extract_paragraph_styling(paragraph)
                
                # Post-process classification based on numbering information
                corrected_level_type = self.correct_level_type_based_on_numbering(
                    level_type, numbering_id, numbering_level, text
                )
                
                # Extract correct list number and check for fixes
                correct_number, was_fixed = self.extract_list_number(
                    numbering_id, numbering_level, number, text
                )
                
                # Map to BWA level using corrected level type
                level_number, bwa_level_name = self.map_to_bwa_level(paragraph, corrected_level_type)
                
                # Create content block with styling information
                block = ContentBlock(
                    text=text,
                    level_type=corrected_level_type,
                    number=correct_number,  # Use the correct number from Word's numbering system
                    content=content,
                    level_number=level_number,
                    bwa_level_name=bwa_level_name,
                    numbering_id=numbering_id,
                    numbering_level=numbering_level,
                    style_name=style_name,
                    # Styling information
                    font_name=styling.get('font_name'),
                    font_size=styling.get('font_size'),
                    font_bold=styling.get('font_bold'),
                    font_italic=styling.get('font_italic'),
                    font_underline=styling.get('font_underline'),
                    font_color=styling.get('font_color'),
                    paragraph_alignment=styling.get('paragraph_alignment'),
                    paragraph_indent_left=styling.get('paragraph_indent_left'),
                    paragraph_indent_right=styling.get('paragraph_indent_right'),
                    paragraph_indent_first_line=styling.get('paragraph_indent_first_line'),
                    paragraph_indent_hanging=styling.get('paragraph_indent_hanging'),
                    paragraph_spacing_before=styling.get('paragraph_spacing_before'),
                    paragraph_spacing_after=styling.get('paragraph_spacing_after'),
                    paragraph_line_spacing=styling.get('paragraph_line_spacing')
                )
                
                self.content_blocks.append(block)
            
            # Build the final data structure
            extracted_data = {
                "header": header_footer_data["header"],
                "footer": header_footer_data["footer"],
                "margins": header_footer_data["margins"],
                "comments": comments,
                "section_number": section_number,
                "section_title": section_title,
                "end_of_section": self.end_of_section,
                "content_blocks": [
                    {
                        "text": block.text,
                        "level_type": block.level_type,
                        "number": block.number,
                        "content": block.content,
                        "level_number": block.level_number,
                        "bwa_level_name": block.bwa_level_name,
                        "numbering_id": block.numbering_id,
                        "numbering_level": block.numbering_level,
                        "style_name": block.style_name,
                        # Styling information
                        "font_name": block.font_name,
                        "font_size": block.font_size,
                        "font_bold": block.font_bold,
                        "font_italic": block.font_italic,
                        "font_underline": block.font_underline,
                        "font_color": block.font_color,
                        "paragraph_alignment": block.paragraph_alignment,
                        "paragraph_indent_left": block.paragraph_indent_left,
                        "paragraph_indent_right": block.paragraph_indent_right,
                        "paragraph_indent_first_line": block.paragraph_indent_first_line,
                        "paragraph_indent_hanging": block.paragraph_indent_hanging,
                        "paragraph_spacing_before": block.paragraph_spacing_before,
                        "paragraph_spacing_after": block.paragraph_spacing_after,
                        "paragraph_line_spacing": block.paragraph_line_spacing
                    }
                    for block in self.content_blocks
                ],
                "template_analysis": self.get_template_analysis_section()
            }
            
            return extracted_data
            
        except Exception as e:
            self.add_error("Extraction Error", f"Failed to extract content: {str(e)}", "")
            return {}
    
    def get_template_analysis_section(self) -> Dict[str, Any]:
        """Get the template analysis section for JSON output"""
        if not hasattr(self, 'template_analysis') or self.template_analysis is None:
            return {
                "template_path": self.template_path,
                "bwa_list_levels": {},
                "template_numbering": {},
                "level_mappings": {},
                "summary": {
                    "total_abstract_numbering": 0,
                    "total_num_mappings": 0,
                    "total_bwa_levels": 0,
                    "level_mappings_count": 0,
                    "level_types": {},
                    "analysis_timestamp": datetime.now().isoformat(),
                    "error": "No template analysis available"
                }
            }
        
        # Convert ListLevelInfo objects to dictionaries for JSON serialization
        bwa_list_levels_dict = {}
        for key, level_info in self.template_analysis.bwa_list_levels.items():
            bwa_list_levels_dict[key] = {
                "level_number": level_info.level_number,
                "numbering_id": level_info.numbering_id,
                "abstract_num_id": level_info.abstract_num_id,
                "level_text": level_info.level_text,
                "number_format": level_info.number_format,
                "start_value": level_info.start_value,
                "suffix": level_info.suffix,
                "justification": level_info.justification,
                "style_name": level_info.style_name,
                "bwa_label": level_info.bwa_label,
                "is_bwa_level": level_info.is_bwa_level
            }
        
        return {
            "template_path": self.template_analysis.template_path,
            "bwa_list_levels": bwa_list_levels_dict,
            "template_numbering": self.template_analysis.numbering_definitions,
            "level_mappings": self.template_analysis.level_mappings,
            "summary": self.template_analysis.summary
        }
    
    def extract_section_header_and_title(self, paragraphs: List[str]) -> Tuple[str, str]:
        """Extract section header and title from the first few paragraphs"""
        section_number = ""
        section_title = ""
        
        if len(paragraphs) >= 2:
            # First paragraph should be section header
            section_text = paragraphs[0].strip()
            if section_text.upper().startswith("SECTION"):
                section_match = re.search(r'^SECTION\s+(.+)$', section_text, re.IGNORECASE)
                if section_match:
                    section_content = section_match.group(1).strip()
                    
                    # Try to extract section number from various formats
                    number_match = re.search(r'(\d+)\s+(\d+)\s+(\d+)', section_content)
                    if number_match:
                        section_number = f"{number_match.group(1)}{number_match.group(2)}{number_match.group(3)}"
                    else:
                        number_match = re.search(r'(\d+)-(\d+)-(\d+)', section_content)
                        if number_match:
                            section_number = f"{number_match.group(1)}{number_match.group(2)}{number_match.group(3)}"
                        else:
                            number_match = re.search(r'(\d{6})', section_content)
                            if number_match:
                                section_number = number_match.group(1)
                            else:
                                section_number = section_content.replace(" ", "").replace("-", "")
            
            # Second paragraph should be section title
            title_text = paragraphs[1].strip()
            if title_text and not title_text.upper().startswith("SECTION"):
                section_title = title_text
        
        return section_number, section_title
    
    def generate_error_report(self) -> str:
        """Generate a comprehensive error report"""
        report = f"ERROR REPORT - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "=" * 60 + "\n\n"
        
        # Report list numbering fixes
        if self.list_fixes:
            report += f"LIST NUMBERING FIXES ({len(self.list_fixes)} found):\n"
            report += "-" * 40 + "\n"
            
            for i, fix in enumerate(self.list_fixes, 1):
                report += f"{i}. Line {fix['line_number']}: {fix['text']}\n"
                report += f"   Detected: {fix['detected_number']}, Corrected: {fix['correct_number']}\n"
                report += f"   Numbering ID: {fix['numbering_id']}, Level: {fix['numbering_level']}\n"
                report += "\n"
        
        # Report other errors
        if self.errors:
            # Group errors by type
            error_types = {}
            for error in self.errors:
                if error.error_type not in error_types:
                    error_types[error.error_type] = []
                error_types[error.error_type].append(error)
            
            for error_type, errors in error_types.items():
                report += f"{error_type} ERRORS ({len(errors)} found):\n"
                report += "-" * 40 + "\n"
                
                for i, error in enumerate(errors, 1):
                    report += f"{i}. Line {error.line_number}: {error.message}\n"
                    if error.context:
                        report += f"   Context: {error.context}\n"
                    if error.expected and error.found:
                        report += f"   Expected: {error.expected}, Found: {error.found}\n"
                    report += "\n"
        else:
            report += "No errors found during extraction.\n\n"
        
        # Add summary statistics
        report += "SUMMARY:\n"
        report += "-" * 20 + "\n"
        total_errors = len(self.errors)
        total_fixes = len(self.list_fixes)
        report += f"Total errors: {total_errors}\n"
        report += f"List numbering fixes: {total_fixes}\n"
        
        if self.errors:
            error_types = {}
            for error in self.errors:
                if error.error_type not in error_types:
                    error_types[error.error_type] = []
                error_types[error.error_type].append(error)
            
            for error_type, errors in error_types.items():
                report += f"{error_type}: {len(errors)} errors\n"
        
        return report
    
    def save_to_json(self, data: Dict[str, Any], output_path: str):
        """Save extracted data to JSON file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    def save_error_report(self, report: str, output_path: str):
        """Save error report to text file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report)
    
    def save_modular_json_files(self, data: Dict[str, Any], base_name: str, output_dir: str):
        """Save separate JSON files for each modular component"""
        try:
            # 1. Header/Footer JSON
            header_footer_data = {
                "header": data.get("header", {}),
                "footer": data.get("footer", {}),
                "margins": data.get("margins", {}),
                "extraction_timestamp": datetime.now().isoformat(),
                "source_file": base_name
            }
            header_footer_path = os.path.join(output_dir, f"{base_name}_header_footer.json")
            with open(header_footer_path, 'w', encoding='utf-8') as f:
                json.dump(header_footer_data, f, indent=2, ensure_ascii=False)
            print(f"Header/footer data saved to: {header_footer_path}")
            
            # 2. Comments JSON
            comments_data = {
                "comments": data.get("comments", []),
                "extraction_timestamp": datetime.now().isoformat(),
                "source_file": base_name
            }
            comments_path = os.path.join(output_dir, f"{base_name}_comments.json")
            with open(comments_path, 'w', encoding='utf-8') as f:
                json.dump(comments_data, f, indent=2, ensure_ascii=False)
            print(f"Comments data saved to: {comments_path}")
            
            # 3. Template Analysis JSON
            template_data = data.get("template_analysis", {})
            if template_data:
                template_path = os.path.join(output_dir, f"{base_name}_template_analysis.json")
                with open(template_path, 'w', encoding='utf-8') as f:
                    json.dump(template_data, f, indent=2, ensure_ascii=False)
                print(f"Template analysis saved to: {template_path}")
            
            # 4. Content Blocks JSON (with list levels and numbering)
            content_data = {
                "section_number": data.get("section_number", ""),
                "section_title": data.get("section_title", ""),
                "end_of_section": data.get("end_of_section", ""),
                "content_blocks": data.get("content_blocks", []),
                "extraction_timestamp": datetime.now().isoformat(),
                "source_file": base_name
            }
            content_path = os.path.join(output_dir, f"{base_name}_content_blocks.json")
            with open(content_path, 'w', encoding='utf-8') as f:
                json.dump(content_data, f, indent=2, ensure_ascii=False)
            print(f"Content blocks saved to: {content_path}")
            
        except Exception as e:
            print(f"Warning: Could not save modular JSON files: {e}")

def main():
    """Main function to run the extraction"""
    if len(sys.argv) < 2:
        print("Usage: python extract_spec_content_v3.py <docx_file> [output_dir] [template_file]")
        print("Example: python extract_spec_content_v3.py 'SECTION 26 05 00.docx'")
        print("Example: python extract_spec_content_v3.py 'SECTION 26 05 00.docx' . 'templates/test_template_cleaned.docx'")
        print("Note: All output files will be saved to <output_dir>/output/")
        print("Note: Template file must be explicitly specified - no auto-detection")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "."
    template_path = sys.argv[3] if len(sys.argv) > 3 else None
    
    if not os.path.exists(docx_path):
        print(f"Error: File '{docx_path}' not found.")
        sys.exit(1)
    
    # Require explicit template specification
    if not template_path:
        print("Error: Template file must be specified as the third argument.")
        print("Example: python extract_spec_content_v3.py 'document.docx' . 'templates/test_template_cleaned.docx'")
        print("Available templates:")
        print("  - templates/test_template_cleaned.docx (recommended)")
        print("  - templates/test_template.docx")
        print("  - templates/test_template_orig.docx")
        sys.exit(1)
    
    if not os.path.exists(template_path):
        print(f"Error: Template file '{template_path}' not found.")
        print("Please ensure the template file exists and the path is correct.")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    # If we're in src directory, go up one level to project root
    if os.path.basename(os.getcwd()) == "src":
        output_dir = os.path.join(os.path.dirname(os.getcwd()), "output")
    else:
        output_dir = os.path.join(output_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    # Initialize extractor with template if provided
    extractor = SpecContentExtractorV3(template_path)
    
    # Extract content
    print(f"Extracting content from '{docx_path}'...")
    data = extractor.extract_content(docx_path)
    
    if not data:
        print("Error: Failed to extract content.")
        sys.exit(1)
    
    # Generate output filenames
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    main_json_path = os.path.join(output_dir, f"{base_name}_v3.json")
    error_path = os.path.join(output_dir, f"{base_name}_v3_errors.txt")
    
    # Save main comprehensive JSON (contains everything)
    extractor.save_to_json(data, main_json_path)
    print(f"Main content saved to: {main_json_path}")
    
    # Save separate modular JSON files
    extractor.save_modular_json_files(data, base_name, output_dir)
    
    # Generate and save error report
    error_report = extractor.generate_error_report()
    extractor.save_error_report(error_report, error_path)
    print(f"Error report saved to: {error_path}")
    
    # Print summary
    print(f"\nExtraction Summary:")
    print(f"- Content blocks found: {len(data.get('content_blocks', []))}")
    print(f"- Header paragraphs: {len(data.get('header', {}).get('paragraphs', []))}")
    print(f"- Footer paragraphs: {len(data.get('footer', {}).get('paragraphs', []))}")
    print(f"- Comments found: {len(data.get('comments', []))}")
    print(f"- BWA list levels: {len(data.get('template_info', {}).get('bwa_list_levels', {}))}")
    print(f"- Errors found: {len(extractor.errors)}")
    
    if extractor.errors:
        print(f"\nWARNING: {len(extractor.errors)} errors were found during extraction.")
        print(f"Please review the error report: {error_path}")
    else:
        print("\nExtraction completed successfully with no errors.")

if __name__ == "__main__":
    main() 
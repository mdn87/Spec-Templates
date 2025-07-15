# SpecConverter v0.4

A comprehensive Python toolkit for extracting and converting specification content from Word documents (.docx) to JSON format, with modular architecture for header/footer extraction, comments processing, and template analysis.

## Project Structure

```
SpecConverter_v0.4/
├── src/                    # Source code files
│   ├── extract_spec_content_v3.py      # Main extraction script (latest version)
│   ├── header_footer_extractor.py      # Modular header/footer extraction
│   ├── template_list_detector.py       # Template list level detection
│   ├── clean_template.py               # Template cleaning utility
│   ├── rip-header-footer.py           # Legacy header/footer extraction
│   ├── rip-comments-to-json.py        # Comments extraction utility
│   ├── debug_doc_content.py           # Debug utility for document content
│   ├── format_xml.py                  # XML formatting utility
│   └── test_generator.py              # Test document generator
├── templates/              # Template files
│   ├── test_template.docx             # Original template
│   ├── test_template_cleaned.docx     # Cleaned template (recommended)
│   └── test_template_orig.docx        # Backup of original template
├── output/                 # Generated output files
│   ├── *.json              # JSON output files
│   ├── *_errors.txt        # Error reports
│   ├── *_processed.docx    # Processed documents
│   └── template_analysis.json # Template analysis results
├── docs/                   # Documentation
│   ├── README_extract_spec_content.md  # Detailed usage guide
│   ├── Dev goals for Spec Templates script.docx # Development goals
│   └── new-multilevel-list-steps.txt  # Implementation steps
├── examples/               # Example documents
│   ├── SECTION 26 05 00.docx          # Example specification section
│   ├── SECTION 26 05 29.docx          # Example specification section
│   └── SECTION 00 00 00.docx          # Example specification section
└── README.md              # This file
```

## Quick Start

### Prerequisites
- Python 3.7+
- Required packages: `python-docx`, `lxml`

### Installation
```bash
pip install python-docx lxml
```

### Basic Usage

1. **Extract content from a specification document:**
   ```bash
   cd src
   python extract_spec_content_v3.py "../examples/SECTION 26 05 00.docx"
   ```

2. **Use with a template for validation:**
   ```bash
   python extract_spec_content_v3.py "../examples/SECTION 26 05 00.docx" . "../templates/test_template_cleaned.docx"
   ```

### Output Structure

The main script generates several output files:

- **Main JSON file**: `{document_name}_v3.json` - Complete extracted data
- **Modular JSON files**:
  - `{document_name}_header_footer.json` - Header/footer data
  - `{document_name}_comments.json` - Comments data
  - `{document_name}_template_analysis.json` - Template analysis
  - `{document_name}_content_blocks.json` - Content blocks with list levels
- **Error report**: `{document_name}_v3_errors.txt` - Validation errors
- **Processing report**: `{document_name}_processing_report.txt` - Summary

## Key Features

### Modular Architecture
- **Header/Footer Extraction**: Separate module for extracting document headers and footers
- **Comments Processing**: Dedicated module for extracting and processing comments
- **Template Analysis**: Comprehensive template structure analysis with numbering patterns
- **Content Block Processing**: Hierarchical content extraction with list level preservation

### Advanced Features
- **Multi-level List Support**: Handles complex hierarchical structures (parts, subsections, items, lists, sub-lists)
- **Numbering Validation**: Validates numbering sequences and reports errors
- **Template Validation**: Compares document structure against template expectations
- **Error Reporting**: Comprehensive error detection and reporting
- **JSON Output**: Structured JSON output for easy processing and reconstruction

### Template Management
- **Template Cleaning**: Utility to clean templates by removing unwanted numbering definitions
- **Template Analysis**: Detailed analysis of template structure and numbering patterns
- **Auto-detection**: Automatically detects cleaned templates for validation

## Script Descriptions

### Main Scripts
- **`extract_spec_content_v3.py`**: Latest version with modular architecture and comprehensive JSON output
- **`extract_spec_content_final.py`**: Previous version with template validation
- **`extract_spec_content_final_v2.py`**: Version with document processing and output generation

### Utility Scripts
- **`header_footer_extractor.py`**: Modular header/footer extraction
- **`template_list_detector.py`**: Template list level detection and analysis
- **`clean_template.py`**: Template cleaning utility
- **`rip-header-footer.py`**: Legacy header/footer extraction
- **`rip-comments-to-json.py`**: Comments extraction utility
- **`debug_doc_content.py`**: Debug utility for document content analysis
- **`format_xml.py`**: XML formatting utility
- **`test_generator.py`**: Test document generator

## JSON Output Structure

The main JSON output includes:

```json
{
  "header": {
    "bwa_number": "2025-XXXX",
    "client_number": "ZZZ# 00000",
    "project_name": "PROJECT NAME",
    "company_name": "CLIENT NAME",
    "section_number": "260500",
    "section_title": "Common Work Results for Electrical"
  },
  "footer": {
    "paragraphs": [],
    "tables": [],
    "text_boxes": []
  },
  "margins": {
    "top_margin": 1.0,
    "bottom_margin": 1.0,
    "left_margin": 1.0833333333333333,
    "right_margin": 1.0833333333333333,
    "header_distance": 1.0,
    "footer_distance": 1.0
  },
  "comments": [],
  "template_analysis": {
    "template_path": "path/to/template.docx",
    "analysis_timestamp": "2025-01-15T12:00:00",
    "paragraphs": [],
    "numbering_definitions": {},
    "level_patterns": {},
    "summary": {}
  },
  "content_blocks": {
    "section_number": "260500",
    "section_title": "Common Work Results for Electrical",
    "parts": [
      {
        "part_number": "1.0",
        "title": "GENERAL",
        "subsections": [
          {
            "subsection_number": "1.01",
            "title": "SCOPE",
            "items": [
              {
                "item_number": "A",
                "text": "Item content",
                "lists": [
                  {
                    "list_number": "1",
                    "text": "List content",
                    "sub_lists": [
                      {
                        "sub_list_number": "a",
                        "text": "Sub-list content"
                      }
                    ]
                  }
                ]
              }
            ]
          }
        ]
      }
    ]
  }
}
```

## Error Handling

The system provides comprehensive error reporting:

- **Structure Errors**: Issues with document hierarchy
- **Numbering Errors**: Broken or inconsistent numbering sequences
- **Content Errors**: Unexpected content types
- **Template Validation Errors**: Mismatches with expected template structure
- **Processing Errors**: General extraction failures

## Development

### Adding New Features
1. Create new modules in the `src/` directory
2. Update the main extraction script to integrate new modules
3. Update this README with new features and usage instructions

### Testing
- Use the example documents in the `examples/` directory
- Compare output with expected results in the `output/` directory
- Review error reports for validation issues

## Version History

- **v0.4**: Modular architecture with separate JSON files for each component
- **v0.3**: Enhanced template analysis and validation
- **v0.2**: Document processing and output generation
- **v0.1**: Basic content extraction and JSON output

## License

This project is for internal use at BWA Engineering. 
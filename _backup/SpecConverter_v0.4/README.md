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

## Template Setup and BWA Style Configuration

Proper template setup is crucial for successful content extraction and regeneration. This section provides detailed instructions for creating and configuring templates with BWA styles.

### Creating a Template Document

1. **Start with a clean Word document**

   - Create a new Word document
   - Set up your desired page margins, headers, and footers
   - Configure any document-level settings (page size, orientation, etc.)
2. **Set document template properties**

   - Go to **File → Options → Advanced**
   - Under "General", check **"Prompt to save Normal template"**
   - Go to **File → Save As**
   - Choose **"Word Template (*.dotx)"** as the file type
   - Save as `test_template_cleaned.docx` in the `templates/` folder

### **CRITICAL: Style Inheritance Requirement**

> **All styles in your template must have**:
>
> - **Style based on: (no style)**, or
> - **Style based on: (a BWA style)**
>
> **Do NOT base any BWA style on Normal, Heading, or any built-in Word style.**
>
> This is required for correct extraction and regeneration. If you use 'Style based on: Normal' or any non-BWA style, formatting and numbering may break.

### Setting Up BWA Styles

The system recognizes specific BWA style names for different content levels. Create these styles in your template:

#### Required BWA Style Names

- `BWA-SectionNumber` - For section numbers (e.g., "SECTION 26 05 00")
- `BWA-SectionTitle` - For section titles
- `BWA-PART` - For part titles (e.g., "1.0 GENERAL")
- `BWA-SUBSECTION` - For subsection titles (e.g., "1.01 SCOPE")
- `BWA-Item` - For item level content (e.g., "A. Description")
- `BWA-List` - For list items (e.g., "1. List content")
- `BWA-SubList` - For sub-list items (e.g., "a. Sub-list content")

#### Creating BWA Styles

1. **Open the Styles pane**

   - Go to **Home → Styles** (or press Ctrl+Alt+Shift+S)
   - Click the **"New Style"** button
2. **Create each BWA style**

   - **Style name**: Enter the exact BWA style name (e.g., "BWA-PART")
   - **Style type**: Choose "Paragraph"
   - **Style based on**: Select **"(no style)"** - This is crucial for proper inheritance
   - **Formatting**: Set your desired font, size, spacing, etc.
   - **Add to Quick Style list**: Check this option
   - Click **OK**
3. **Configure style properties**

   - Right-click each BWA style in the Styles pane
   - Select **"Modify"**
   - Set appropriate formatting:
     - **Font**: Arial (recommended)
     - **Size**: 10pt (recommended)
     - **Alignment**: Left (for most styles)
     - **Spacing**: Set before/after spacing as needed
     - **Indentation**: Set left/right indentation for hierarchy

### Setting Up Multilevel Lists

1. **Create a multilevel list**

   - Go to **Home → Multilevel List**
   - Click **"Define New Multilevel List"**
2. **Configure each level**

   - **Level 1**: Set to decimal format (1, 2, 3...)
   - **Level 2**: Set to decimal format (1.01, 1.02...)
   - **Level 3**: Set to upper letter format (A, B, C...)
   - **Level 4**: Set to decimal format (1, 2, 3...)
   - **Level 5**: Set to lower letter format (a, b, c...)
3. **Link levels to BWA styles**

   - For each level, click **"More"**
   - In **"Link level to style"**, select the corresponding BWA style
   - Set appropriate indentation and alignment

### Configuring Template for New Documents *** THIS MAY NOT BE ACCURATE ***

1. **Set template as default for new documents**

   - Go to **File → Options → Advanced**
   - Under "General", click **"File Locations"**
   - Set **"User templates"** to point to your template folder
   - Go to **File → Save As**
   - Choose **"Word Template (*.docx)" ** .dotx does not work**
   - Save as `Normal.dotx` in the User templates folder
2. **Enable "New documents based on this template"**

   - Open your template document
   - Go to **File → Options → Advanced**
   - Under "General", check **"Prompt to save Normal template"**
   - Save the template

### Template Validation

After setting up your template:

1. **Test the template**

   - Create a new document based on your template
   - Apply BWA styles to sample content
   - Verify that multilevel lists work correctly
   - Check that numbering and formatting are preserved
2. **Run template analysis**

   ```bash
   cd src
   python template_list_detector.py "../templates/your_template.docx"
   ```

   This will analyze your template and report any issues.

### Common Template Issues and Solutions

#### Issue: Styles not applying correctly in regenerated documents

**Solution**: Ensure all BWA styles are set to "Style based on (no style)" and the template is set to "New documents based on this template".

#### Issue: Courier font appearing unexpectedly

**Solution**: Check that your template's Normal style uses Arial font, and all BWA styles inherit from "(no style)" rather than Normal.

#### Issue: Numbering not working in regenerated documents

**Solution**: Verify that multilevel lists are properly linked to BWA styles in the template.

#### Issue: Header/footer formatting not preserved

**Solution**: Ensure header and footer styles in the template use the same fonts and formatting you want in the final documents.

### Template Best Practices

1. **Use consistent naming**: Always use the exact BWA style names listed above
2. **Avoid style inheritance**: Set all BWA styles to "Style based on (no style)"
3. **Test thoroughly**: Create sample documents to verify template functionality
4. **Document your setup**: Keep notes on your template configuration for future reference
5. **Version control**: Save multiple versions of your template as you refine it

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

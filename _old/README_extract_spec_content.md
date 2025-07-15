# Specification Content Extractor

A Python script to extract multi-level list content from Word documents (.docx) and convert it to JSON format. This tool is specifically designed for specification documents with hierarchical structures including parts, subsections, items, and lists.

## Features

- **Multi-level Content Extraction**: Handles parts, subsections, items, and lists
- **Flexible Structure Recognition**: Works with both numbered and unnumbered document structures
- **Error Detection & Reporting**: Identifies and reports structural issues and numbering inconsistencies
- **JSON Output**: Generates structured JSON files compatible with specification systems
- **Comprehensive Error Reports**: Detailed error analysis with line numbers and context

## Requirements

- Python 3.6+
- python-docx library

Install dependencies:

```bash
pip install python-docx
```

## Usage

### Basic Usage

```bash
python extract_spec_content_final.py <docx_file> [output_dir]
```

### Examples

```bash
# Extract content from a specification document
python extract_spec_content_final.py "SECTION 26 05 00.docx"

# Extract content and save to specific directory
python extract_spec_content_final.py "SECTION 26 05 00.docx" "output_folder"
```

## Output Files

The script generates two output files:

1. **`{filename}_content.json`** - The extracted content in JSON format
2. **`{filename}_errors.txt`** - Detailed error report

### JSON Structure

The extracted JSON follows this structure:

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
              "text": "Item content...",
              "lists": [
                {
                  "list_number": "1",
                  "text": "List item content..."
                }
              ]
            }
          ]
        }
      ]
    }
  ]
}
```

## Document Structure Recognition

The script recognizes the following content types:

### Section Level

- **Section Header**: `SECTION 26 05 00`
- **Section Title**: `Common Work Results for Electrical`

### Part Level

- **Numbered Parts**: `1.0 GENERAL`, `2.0 PRODUCTS`
- **Unnumbered Parts**: `GENERAL`, `PRODUCTS`, `EXECUTION`

### Subsection Level

- **Numbered Subsections**: `1.01 SCOPE`, `1.02 EXISTING CONDITIONS`
- **Unnumbered Subsections**: `SCOPE`, `EXISTING CONDITIONS`, `CODES AND REGULATIONS`

### Item Level

- **Items**: `A. Item content`, `B. Item content`

### List Level

- **Lists**: `1. List item`, `2. List item`
- **Sub-lists**: `a. Sub-list item`, `b. Sub-list item`

## Error Types

The script detects and reports several types of errors:

### Structure Errors

- Items found without preceding subsections
- Subsections found without preceding parts
- List items found without parent items

### Numbering Sequence Errors

- Unexpected part numbers (e.g., found 2.0 when expecting 1.0)
- Unexpected subsection numbers (e.g., found 1.03 when expecting 1.02)
- Unexpected item numbers (e.g., found C when expecting B)

### Content Warnings

- Unstructured content that doesn't match expected patterns
- Content that appears to be continuation text

## Error Report Format

Error reports include:

- Line number where the error occurred
- Error type and description
- Context (the actual content that caused the error)
- Expected vs. found values (for numbering errors)

Example error report:

```
ERROR REPORT - 2025-07-10 13:10:07
============================================================

Structure Error (23 errors):
-------------------------
Line 5: Item found without preceding subsection
  Context: Division 26 includes all Specifications in the 26 00 00 Series...

Numbering Sequence Error (2 errors):
---------------------------------
Line 33: Unexpected part number
  Context: PRODUCTS
  Expected: 1.0, Found: 2.0
```

## Handling Different Document Formats

The script is designed to handle various document structures:

### Well-Structured Documents

Documents with proper numbering (1.0, 1.01, A., 1., etc.) will extract cleanly with minimal errors.

### Partially Structured Documents

Documents with some numbering but missing elements will be processed with auto-generated numbers and error reporting.

### Unstructured Documents

Documents without clear structure will still be processed, but will generate many content warnings and structure errors.

## Customization

### Adding New Subsection Titles

To recognize additional subsection titles, modify the `subsection_titles` list in the `parse_paragraph_content` method:

```python
subsection_titles = [
    "SCOPE", "EXISTING CONDITIONS", "CODES AND REGULATIONS", "DEFINITIONS",
    "DRAWINGS AND SPECIFICATIONS", "SITE VISIT", "DEVIATIONS",
    "STANDARDS FOR MATERIALS AND WORKMANSHIP", "SHOP DRAWINGS AND SUBMITTAL",
    "RECORD (AS-BUILT) DRAWINGS AND MAINTENANCE MANUALS",
    "COORDINATION", "PROTECTION OF MATERIALS", "TESTS, DEMONSTRATION AND INSTRUCTIONS",
    "GUARANTEE",
    "YOUR_NEW_SUBSECTION_TITLE"  # Add new titles here
]
```

### Modifying Regex Patterns

Adjust the regex patterns in the `__init__` method to match different numbering formats:

```python
self.part_pattern = re.compile(r'^(\d+\.0)\s+(.+)$')
self.subsection_pattern = re.compile(r'^(\d+\.\d{2})\s+(.+)$')
self.item_pattern = re.compile(r'^([A-Z])\.\s+(.+)$')
```

## Troubleshooting

### Common Issues

1. **"File not found" error**

   - Ensure the .docx file exists in the specified path
   - Check file permissions
2. **"Could not open document" error**

   - Verify the file is a valid .docx file
   - Check if the file is corrupted
3. **Many structure errors**

   - The document may not follow the expected specification format
   - Review the error report to understand the actual document structure
4. **Missing content**

   - Check if the document uses different formatting than expected
   - Review the debug output to see the actual document structure

### Debug Mode

To see the raw document structure, use the debug script:

```bash
python debug_doc_content.py "your_document.docx"
```

This will show each paragraph with its style and any numbering information.

## Performance

- Processing time depends on document size and complexity
- Typical processing time: 1-5 seconds for standard specification documents
- Memory usage is minimal and scales with document size

## Limitations

- Only works with .docx files (not .doc files)
- Relies on text content rather than Word formatting styles
- May not handle complex nested structures perfectly
- Requires documents to follow general specification document patterns

## Contributing

To improve the script:

1. Test with different document formats
2. Add new pattern recognition rules
3. Improve error reporting
4. Add support for additional content types

## License

This script is provided as-is for educational and professional use.

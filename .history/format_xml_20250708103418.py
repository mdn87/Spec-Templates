import xml.dom.minidom


def format_xml_file(input_file, output_file):
    """Format XML file with proper indentation"""
    try:
        # Read the XML content
        with open(input_file, 'r', encoding='utf-8') as f:
            xml_content = f.read()
        
        # Parse and format the XML
        dom = xml.dom.minidom.parseString(xml_content)
        formatted_xml = dom.toprettyxml(indent="  ")
        
        # Write the formatted XML
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(formatted_xml)
        
        print(f"XML formatted successfully! Output saved to: {output_file}")
        
    except Exception as e:
        print(f"Error formatting XML: {e}")

if __name__ == "__main__":
    input_file = "Spec Templates/Ready for Final Release/26 05 19 Low-voltage Electrical Power Conductors and Cables/word/styles.xml"
    output_file = "Spec Templates/Ready for Final Release/26 05 19 Low-voltage Electrical Power Conductors and Cables/word/styles_formatted.xml"
    format_xml_file(input_file, output_file) 
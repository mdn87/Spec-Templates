import xml.dom.minidom
import os


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
    # Get the input filename from user
    input_filename = input("Enter the name of the XML file to format (e.g., styles.xml): ")
    
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Construct full paths
    input_file = os.path.join(script_dir, input_filename)
    
    # Generate output filename by adding _formatted before .xml
    name_without_ext = os.path.splitext(input_filename)[0]
    output_filename = f"{name_without_ext}_formatted.xml"
    output_file = os.path.join(script_dir, output_filename)
    
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    
    format_xml_file(input_file, output_file) 
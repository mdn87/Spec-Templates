import xml.dom.minidom
import os
import sys


def pretty_format_xml(file_path):
    """Format XML file with proper indentation and save back to same file"""
    try:
        # Read the XML content
        with open(file_path, 'r', encoding='utf-8') as f:
            xml_content = f.read()
        
        # Parse and format the XML
        dom = xml.dom.minidom.parseString(xml_content)
        formatted_xml = dom.toprettyxml(indent="  ")
        
        # Write the formatted XML back to the same file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(formatted_xml)
        
        print(f"XML formatted successfully! File updated: {file_path}")
        return True
        
    except Exception as e:
        print(f"Error formatting XML: {e}")
        return False


def main():
    if len(sys.argv) > 1:
        # If filename provided as command line argument
        file_path = sys.argv[1]
    else:
        # Get the filename from user input
        file_path = input("Enter the path to the XML file to format: ")
    
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # If only filename provided, assume it's in the same directory as script
    if not os.path.isabs(file_path):
        file_path = os.path.join(script_dir, file_path)
    
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        return
    
    # Check if it's an XML file
    if not file_path.lower().endswith('.xml'):
        print(f"Warning: File doesn't have .xml extension: {file_path}")
        response = input("Continue anyway? (y/n): ")
        if response.lower() != 'y':
            return
    
    print(f"Formatting XML file: {file_path}")
    pretty_format_xml(file_path)


if __name__ == "__main__":
    main() 
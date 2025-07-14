from docx import Document
import json
import os
from datetime import datetime

def extract_comments(docx_path):
    doc = Document(docx_path)
    comments = []
    
    # Check if the document has comments
    if hasattr(doc.part, '_comments_part') and doc.part._comments_part is not None:
        for c in doc.part._comments_part.comments:
            # Assemble the full comment text (may be multiple paragraphs)
            full_text = "\n".join(p.text for p in c.paragraphs).strip()
            
            # Extract comment metadata
            comment_data = {
                "text": full_text,
                "ref": None
            }
            
            # Try to get comment ID
            try:
                comment_data["id"] = str(c._element.get('id', ''))
            except:
                comment_data["id"] = None
            
            # Try to get author
            try:
                comment_data["author"] = str(c._element.get('author', ''))
            except:
                comment_data["author"] = "Unknown"
            
            # Try to get date
            try:
                date_attr = c._element.get('date', '')
                if date_attr:
                    # Convert Word's date format to readable format
                    try:
                        # Word dates are often in a specific format
                        comment_data["date"] = str(date_attr)
                    except:
                        comment_data["date"] = str(date_attr)
                else:
                    comment_data["date"] = None
            except:
                comment_data["date"] = None
            
            # Try to get initials
            try:
                comment_data["initials"] = str(c._element.get('initials', ''))
            except:
                comment_data["initials"] = None
            
            comments.append(comment_data)
    else:
        print("No comments found in the document")
    
    return comments

def save_comments_to_json(comments, output_path):
    """Save comments to JSON file"""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump({"comments": comments}, f, indent=2)
    print(f"Comments saved to JSON: {output_path}")

def save_comments_to_txt(comments, output_path):
    """Save comments to TXT file"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("COMMENTS FROM DOCUMENT\n")
        f.write("=" * 50 + "\n\n")
        for i, comment in enumerate(comments, 1):
            f.write(f"Comment {i}:\n")
            f.write(f"Text: {comment['text']}\n")
            if comment.get('author'):
                f.write(f"Author: {comment['author']}\n")
            if comment.get('date'):
                f.write(f"Date: {comment['date']}\n")
            if comment.get('initials'):
                f.write(f"Initials: {comment['initials']}\n")
            if comment.get('id'):
                f.write(f"ID: {comment['id']}\n")
            f.write("-" * 30 + "\n\n")
    print(f"Comments saved to TXT: {output_path}")

if __name__ == "__main__":
    try:
        docx_file = "SECTION 26 05 00.docx"
        comments = extract_comments(docx_file)
        
        if comments:
            # Save to JSON
            json_file = "SECTION 26 05 00_comments.json"
            save_comments_to_json(comments, json_file)
            
            # Save to TXT
            txt_file = "SECTION 26 05 00_comments.txt"
            save_comments_to_txt(comments, txt_file)
            
            print(f"\nExtracted {len(comments)} comments from {docx_file}")
            print("Files created:")
            print(f"  - {json_file}")
            print(f"  - {txt_file}")
            
            # Show a preview of the extracted data
            print("\nComment Preview:")
            for i, comment in enumerate(comments[:2], 1):  # Show first 2 comments
                print(f"  Comment {i}:")
                print(f"    Author: {comment.get('author', 'Unknown')}")
                print(f"    Date: {comment.get('date', 'Unknown')}")
                print(f"    Text: {comment['text'][:50]}...")
                print()
        else:
            print("No comments found to save")
            
    except FileNotFoundError:
        print("File 'SECTION 26 05 00.docx' not found")
    except Exception as e:
        print(f"Error processing document: {e}")

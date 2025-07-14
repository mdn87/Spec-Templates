from docx import Document
import json
from datetime import datetime

def extract_comments(docx_path):
    doc = Document(docx_path)
    comments = []
    
    # Check if the document has comments
    if hasattr(doc.part, '_comments_part') and doc.part._comments_part is not None:
        for c in doc.part._comments_part.comments:
            # Assemble the full comment text (may be multiple paragraphs)
            full_text = "\n".join(p.text for p in c.paragraphs).strip()
            
            # Safely get comment attributes
            comment_data = {
                "text": full_text,
                "ref": None
            }
            
            # Add available attributes
            if hasattr(c, 'author'):
                comment_data["author"] = c.author
            if hasattr(c, 'date'):
                try:
                    comment_data["date"] = c.date.isoformat() if hasattr(c.date, 'isoformat') else str(c.date)
                except:
                    comment_data["date"] = str(c.date)
            if hasattr(c, 'id'):
                comment_data["id"] = c.id
                
            comments.append(comment_data)
    else:
        print("No comments found in the document")
    
    return comments

if __name__ == "__main__":
    try:
        comments = extract_comments("SECTION 26 05 00.docx")
        print(json.dumps({"comments": comments}, indent=2))
    except FileNotFoundError:
        print("File 'SECTION 26 05 00.docx' not found")
    except Exception as e:
        print(f"Error processing document: {e}")

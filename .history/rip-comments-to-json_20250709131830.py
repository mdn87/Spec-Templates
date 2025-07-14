from docx import Document
import json

def extract_comments(docx_path):
    doc = Document(docx_path)
    comments = []
    
    # Check if the document has comments
    if hasattr(doc.part, '_comments_part') and doc.part._comments_part is not None:
        for c in doc.part._comments_part.comments:
            # Assemble the full comment text (may be multiple paragraphs)
            full_text = "\n".join(p.text for p in c.paragraphs).strip()
            
            comments.append({
                "text": full_text,
                "ref": None
            })
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

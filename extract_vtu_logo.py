from docx import Document
import os

doc = Document(r"C:\Users\user\Desktop\dsatm project\REPORT FORMATS\VTU ELE.docx")

# Extract images
for rel in doc.part.rels.values():
    if "image" in rel.target_ref:
        img_data = rel.target_part.blob
        img_name = f"vtu_logo_{rel.target_ref.split('/')[-1]}"
        with open(f"C:\\Users\\user\\Desktop\\dsatm project\\web_sports_app\\static\\{img_name}", 'wb') as f:
            f.write(img_data)
        print(f"Extracted: {img_name}")
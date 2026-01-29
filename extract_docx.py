from docx import Document

doc = Document(r"C:\Users\user\Desktop\dsatm project\REPORT FORMATS\HOD bonifide.docx")
text = ""
for paragraph in doc.paragraphs:
    text += paragraph.text + "\n"

print(text)
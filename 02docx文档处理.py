from docx import Document

def replace_text_in_docx(doc_path, old_text, new_text, output_path):
    doc = Document(doc_path)
    for para in doc.paragraphs:
        if old_text in para.text:
            para.text = para.text.replace(old_text, new_text)
    doc.save(output_path)

replace_text_in_docx("学习规划.docx", "[]", "新文本", "output.docx")

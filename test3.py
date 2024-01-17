import docx
from docx import Document

# 워드 파일 열기
doc = Document('C:\\Users\\hy\\OneDrive\\eng\\Short Essay_hy.docx')
styles = doc.styles
style = styles[WD_STYLE.BODY_TEXT]
print(style)

# 단락 읽기
for para in doc.paragraphs:
    print(para.text)

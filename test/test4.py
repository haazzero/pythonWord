import docx
from docx import Document

# 워드 파일 열기
doc = Document('C:\\Users\\hy\\OneDrive\\eng\\Short Essay_hy.docx')

# 형광펜 단어를 담을 공간
highlighted_text = []

# 단락 읽기와 font 객체
for para in doc.paragraphs:
    for run in para.runs :
        if run.font.highlight_color:
            highlighted_text.append(run.text)

print(f'모르는 단어 수 : {len(highlighted_text)}')
print(highlighted_text)
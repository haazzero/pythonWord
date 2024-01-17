import docx
from docx import Document

# 워드 파일 열기
doc = Document('C:\\Users\\hy\\OneDrive\\eng\\Short Essay_hy.docx')
word_count = 0

# 각각의 단락에서 단어 개수 세기
for para in doc.paragraphs:
    words = para.text.split()
    word_count += len(words)

print(f' 총 단어 수: {word_count}')

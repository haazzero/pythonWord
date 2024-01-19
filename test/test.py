# 관련 패키지 import
import win32com.client

# Word 어플리케이션 열기
word = win32com.client.gencache.EnsureDispatch("Word.application")

# 자동으로 실행되는 것을 확인하고 싶을 때 코드 추가
word.Visible = True

# 새로운 워드 문서 열기
# doc = word.Documents.Add()

# 절대 경로를 이용하여 기존 워드 파일 열기
doc = word.Documents.Open(r'C:\Users\hy\OneDrive\eng\Short Essay_hy.docx')
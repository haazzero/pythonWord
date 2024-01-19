# Excel 라이브러리
import openpyxl
from openpyxl import Workbook

# Workbook 객체 생성
wb = Workbook()
# 활성화된 워크시트 선택 후 ws 변수에 할당
ws = wb.active
# 시트 제목을 Highlighted Word로 설정
ws.title = "Highlighted Words" 
# 시트의 A1셀에 '단어'라는 데이터를 입력
ws['A1'] = '단어'
# 워크북 엑셀 파일 저장 (제목)
wb.save('C:\\Users\\hy\\OneDrive\\eng\\test.xlsx')
# 워크북 닫기
wb.close()


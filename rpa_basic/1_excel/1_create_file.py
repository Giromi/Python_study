from openpyxl import Workbook

# 워크 북 생성
wb = Workbook()

#현재 활성화된 sheet 가져옴
ws = wb.active

#Sheet의 이름을 변경
ws.title = "Giromi Sheet"
wb.save("sample.xlsx")
wb.close()


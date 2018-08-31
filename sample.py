from openpyxl import Workbook

# Reference: https://openpyxl.readthedocs.io/en/stable/tutorial.html#create-a-workbook

# 워크북은 항상 최소한 하나의 워크시트를 갖고 있음
wb = Workbook()

# 활성화된 워크시트 가져 오기
# openpyxl.workbook.Workbook.active()
ws = wb.active

# 새로운 워크시트 생성
ws1 = wb.create_sheet("My sheet1")      # 가장 마지막 위치에 추가(기본값)
ws2 = wb.create_sheet("My sheet2", 0)   # 가장 앞에 추가

# 워크시트 제목 변경
ws.title = "New Title"

# 워크시트 제목 탭 색상 변경
ws.sheet_properties.tabColor = "1072BA"     # RRGGBB

# 워크시트 제목을 설정 했으면 설정한 제목을 사용해서 워크시트를 가져올 수 있음
ws3 = wb["New Title"]

# 워크북에 있는 워크시트 이름 모두 가져오기 (openpyxl.workbook.Workbook.sheetnames())
print(wb.sheetnames)

for sheet in wb:
    print(sheet.title)

# 엑셀 파일 저장
wb.save('sample.xlsx')

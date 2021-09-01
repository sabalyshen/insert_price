#讀取檔案
from openpyxl import load_workbook
wb = load_workbook(filename = '110年零用金流向0610.xlsx')
ws = wb['0804']

from openpyxl import load_workbook
wb2 = load_workbook(filename = '8-1-一般-ETAG.xlsx')
ws2 = wb2['101年度大隊黏貼憑證用紙']

# 依照項次加入品名
list_name = []
def wordfinder_name(searchString):
    for i in range(1, ws.max_row):
        for j in range(1, ws.max_column):
            if searchString == ws.cell(i,j).value:
                list_name.append(ws.cell(i,j + 2).value)
                ws2.cell(len(list_name) + 20, 1).value = ws.cell(i,j + 2).value

# 依照項次加入單價
list_price = []
def wordfinder(searchString):
    for i in range(1, ws.max_row):
        for j in range(1, ws.max_column):
            if searchString == ws.cell(i,j).value:
                list_price.append(ws.cell(i,j + 5).value)
                ws2.cell(len(list_price) + 20, 13).value = ws.cell(i,j + 5).value

# 依照項次加入數量
list_quantity = []
def wordfinder_quantity(searchString):
    for i in range(1, ws.max_row):
        for j in range(1, ws.max_column):
            if searchString == ws.cell(i,j).value:
                list_quantity.append(ws.cell(i,j + 4).value)
                ws2.cell(len(list_quantity) + 20, 10).value = ws.cell(i,j + 4).value

# 依照項次加入單位
list_type = []
def wordfinder_type(searchString):
    for i in range(1, ws.max_row):
        for j in range(1, ws.max_column):
            if searchString == ws.cell(i,j).value:
                list_type.append(ws.cell(i,j + 3).value)
                ws2.cell(len(list_type) + 20, 10).value = ws.cell(i,j + 3).value


d = input('請輸入項次:')
wordfinder(d)
wordfinder_quantity(d)
wordfinder_name(d)
wordfinder_type(d)
wb2.save(filename = '成品.xlsx')  

ws2.cell(29, 1).value = '2. ■本案經詢價擬以' + 'ws2.cell(27, 17).value' + '元交由   交通部高速公路局  辦理，並經驗收合格後付款。' 
print(list_price) 
print(len(list_price))
wb2.save(filename = '成品.xlsx')  
 
'''
清單金額填到請示單

mon = a.value // 10000
senn = (a.value- mon*10000) // 1000
hyaku = (a.value- mon*10000- senn*1000) // 100
ju = (a.value- mon*10000- senn*1000- hyaku*100) // 10
enn = a.value % 10

if mon > 0: #萬
    ws2['K6'].value = mon
elif mon == 0:
    ws2['K6'].value = None

if senn > 0: #千
    ws2['L6'].value = senn
elif senn == 0 and mon > 0:
    ws2['L6'].value = 0
elif senn == 0:
    ws2['L6'].value = None

if hyaku > 0: #百
    ws2['M6'].value = hyaku
elif hyaku == 0 and mon > 0:
    ws2['M6'].value = 0
elif hyaku == 0 and senn > 0:
    ws2['M6'].value = 0
elif hyaku == 0:
    ws2['M6'].value = None

if ju > 0: #十
    ws2['N6'].value = ju
elif ju == 0 and mon > 0:
    ws2['N6'].value = 0
elif ju == 0 and senn > 0:
    ws2['N6'].value = 0
elif ju == 0 and hyaku > 0:
    ws2['N6'].value = 0
elif ju == 0:
    ws2['N6'].value = None
ws2['O6'].value = enn
   
 ''' 






    

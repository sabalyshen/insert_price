#讀取檔案
from openpyxl import load_workbook
wb = load_workbook(filename = '110年零用金流向0610.xlsx')
ws = wb['0804']

from openpyxl import load_workbook
wb2 = load_workbook(filename = '8-1-一般-ETAG.xlsx')
ws2 = wb2['101年度大隊黏貼憑證用紙']


#TEST
list = []

# 依照項次加入清單
def wordfinder(searchString):
    for i in range(1, ws.max_row):
        for j in range(1, ws.max_column):
            if searchString == ws.cell(i,j).value:
                list.append(ws.cell(i,j + 3).value)
                ws2.cell(len(list) + 20, 13).value = ws.cell(i,j + 3).value

'''def insert_price(list_name):
    for d in range(len(list_name)):
        for l in list_name:
           
            ws2.cell(d + 21, 13).value = l'''
           
            
wordfinder("one")
#insert_price(list)
wb2.save(filename = '8-1-一般-ETAG.xlsx')   
print(list) 
print(len(list))

 
'''
ws.cell(i,j + 3).value
wordfinder("one")
insert_price(list)
wb2.save(filename = '8-1-一般-ETAG.xlsx')

wb2.save(filename = '8-1-一般-ETAG.xlsx')

for p in range(21, ws2.max_column):
    for l in list:
        l = int(l)
        p = l
        print(l)
wb2.save(filename = '8-1-一般-ETAG.xlsx')

#抓檔案
colname = {}
current = 0
for a in ws.iter_cols(1, ws.max_column): #1是開始值
    colname[a[0].value] = current
    current += 1

colname1 = {}
current1 = 0
for b in ws2.iter_cols(13):
    colname1[b[0].value] = current1
    current1 += 1
print(colname1)


for row_cells in ws.iter_rows(1, ws.max_row):
    if row_cells[colname['no.']].value == 1: 


   

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

wb2.save(filename = '8-1-一般-ETAG.xlsx') 
    
 ''' 






    

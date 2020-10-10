from openpyxl import load_workbook
from openpyxl import Workbook

# 获取工作表
rwb = load_workbook('C:/Users/Administrator/Desktop/new.xlsx')
rws = rwb.active

d = {}
d['编号']=[]
d['直播价']=[]

for i in range(1, rws.max_column+1):
	if rws.cell(1, i).value == '编号':
		for j in range(2, rws.max_row+1):
			d['编号'].append(rws.cell(j, i).value)
	if rws.cell(1, i).value == '直播价':
		for j in range(2, rws.max_row+1):
			d['直播价'].append(rws.cell(j, i).value)


wwb = Workbook()
wws = wwb.active

wws.cell(1,1).value = '编号'
wws.cell(1,2).value = '直播价'

for i in range(2, rws.max_row+1):
    wws.cell(i, 1).value = d['编号'][i-2]
    wws.cell(i, 2).value = d['直播价'][i-2]

wwb.save('C:/Users/Administrator/Desktop/wwb.xlsx')


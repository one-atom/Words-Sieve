import sys
reload(sys)
import string
sys.setdefaultencoding('utf-8')

import xlrd
import xlwt

fname = 'receive.xls'
data1 = xlrd.open_workbook(fname)
allsheet = data1.sheet_by_name("RECEIVE")

fname = 'all_words.xls'
data2 = xlrd.open_workbook(fname)
toeflsheet = data2.sheet_by_name("TOEFL")
gresheet = data2.sheet_by_name("GRE")
trows = toeflsheet.nrows
grows = gresheet.nrows

workbook = xlwt.Workbook(encoding = 'ascii')
worksheet1 = workbook.add_sheet('TOEFL', cell_overwrite_ok = True)
worksheet2 = workbook.add_sheet('GRE', cell_overwrite_ok = True)


wait_to_deletetf = range(0,allsheet.nrows)

for i in range(0,allsheet.nrows):
	wait_to_deletetf[i] = allsheet.cell(i,0).value


for i in range(0,trows):
	if toeflsheet.cell(i,0).value in wait_to_deletetf:
		worksheet1.write(i, 0, label = '')
		worksheet1.write(i, 1, label = '')
	else :
		worksheet1.write(i, 0, label = toeflsheet.cell(i,0).value)
		worksheet1.write(i, 1, label = toeflsheet.cell(i,1).value)

		
wait_to_deletetf = range(0,allsheet.nrows)
for i in range(0,allsheet.nrows):
	wait_to_deletetf[i] = allsheet.cell(i,2).value


for i in range(0,grows):
	if gresheet.cell(i,0).value in wait_to_deletetf:
		worksheet2.write(i, 0, label = '')
		worksheet2.write(i, 1, label = '')
	else :
		worksheet2.write(i, 0, label = gresheet.cell(i,0).value)
		worksheet2.write(i, 1, label = gresheet.cell(i,1).value)

workbook.save('all_words.xls')


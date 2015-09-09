import sys
reload(sys)
import string
sys.setdefaultencoding('utf-8')


import xlrd
import xlwt

#global variations
row_data2 = 'ljALKAJSFD'
row_data3 = 'asdfasdlj'
row_data4 = 'skjfkds'
row_data6 = 'asfasdf'
row_data7 = 'alkjsdf'

def change(row_data):	
	global row_data2
	global row_data3
	global row_data4
	global row_data5
	global row_data6
	global row_data7
	row_data2 = 'ljALKAJSFD'
	row_data3 = 'asdfasdlj'
	row_data4 = 'skjfkds'
	row_data5 = 'asdfsdf'
	row_data6 = 'asfasdf'
	row_data6 = 'asfasdf'	
	if row_data[len(row_data) -1 ] == 'y':
		row_data2 = row_data[0:len(row_data)-1] + 'ie'
		row_data3 = row_data[0:len(row_data)-1] + 'ies'
	if row_data[len(row_data) -1 ] == 'f':
		row_data4 = row_data[0:len(row_data)-1] + 'ves'
	if row_data[(len(row_data ) -2) : (len(row_data) -1)] == 'on' or row_data[(len(row_data ) -2) : (len(row_data) -1)] == 'un':
		row_data5 = row_data[0:len(row_data)-2] + 'a'
	if row_data[(len(row_data ) -2) : (len(row_data) -1)] == 'us':
		row_data6 =  row_data[0:len(row_data)-2] + 'i'
	if row_data[len(row_data) -1 ] == 'e':
		row_data4 = row_data[0:len(row_data)-1] + 'ing'

def judge(line,row_data,row_data2,row_data3,row_data4,row_data5,row_data6,row_data7):
		
		#global punctuation

		


		
		for items in [row_data,row_data2,row_data3,row_data4,row_data5,row_data6,row_data7]:
			if items in line:
				return 'yes' 	

		

	
fname = 'all_words.xlsx'
data = xlrd.open_workbook(fname)

workbook = xlwt.Workbook(encoding = 'ascii')
worksheet = workbook.add_sheet('RECEIVE', cell_overwrite_ok = True)

toeflsheet = data.sheet_by_name("TOEFL")
gresheet = data.sheet_by_name("GRE")


trows = toeflsheet.nrows
grows = gresheet.nrows


txt = open('passage.txt', 'r')

def ranging(n):

	fname1 = 'receive.xls'
	data1 = xlrd.open_workbook(fname1)
	wordsheet1 = data1.sheet_by_name('RECEIVE')
	col = wordsheet1.nrows

	global workbook
	global worksheet

	count = 0
	
	for k in range(0,col):
#		print k,n
		row_data1 = wordsheet1.cell(k,n).value
		if row_data1 != '':

			

			worksheet.write(count, n, label = row_data1)

			worksheet.write(k, n, label = '')
			count += 1
			
#			wordsheet1.put_cell(count, n, 1, wordsheet1.cell(k,n).value, 0)
#			wordsheet1.put_cell(k, n, 0 , '', 0)
	workbook.save('receive.xls')	


# search in toefl vocabulary


line = txt.readline().lower()

while line:

	for a in line:
		if a in string.punctuation:
			line = line.replace(a,' ')		
	m = line.split()

	for i in range(0,trows):		
		row_data = toeflsheet.cell(i,0).value
		explanation = toeflsheet.cell(i,1).value			
		change(row_data)
		if judge(m,row_data,row_data2,row_data3,row_data4,row_data5,row_data6,row_data7) == 'yes':
			worksheet.write(i, 0, label = row_data)
			worksheet.write(i, 1, label = explanation)
		
	line = txt.readline().lower()
		

txt.close()



txt = open('passage.txt', 'r')

line = txt.readline().lower()
while line:
	for a in line:
		if a in string.punctuation:
			line = line.replace(a,' ')		
	m = line.split()

	for i in range(0,grows):
		row_data = gresheet.cell(i,0).value
		explanation = gresheet.cell(i,1).value
		change(row_data)
		if judge(m,row_data,row_data2,row_data3,row_data4,row_data5,row_data6,row_data7) == 'yes':
			worksheet.write(i, 2, label = row_data)	
			worksheet.write(i, 3, label = explanation)	
	line = txt.readline().lower()

txt.close()


workbook.save('receive.xls')

ranging(0)
ranging(1)
ranging(2)
ranging(3)







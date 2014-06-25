from Database import *
from Workbook import *

def error(msg):
	print 'Error' + msg
	exit()

def choose_query(n):
	if n == 1:
		return 'ENG-Software'
	elif n == 2:
		return 'ENG-Controls'
	else:
		error('invalid choice!')

def analysis(db, ws, wb):
	for row in db.data():
		# check status for row assignment
		rownum = get_rownum(row) 
		# check urgency
		if row[4] == 'Med/Low':
			if row[8] == 'Future Commit':
				ws['E%s'%(rownum)] = ws['E%s'%(rownum)].value + 1
			elif row[8] == 'Past Commit':
				ws['E%s'%(rownum+1)] = ws['E%s'%(rownum+1)].value + 1
			elif row[8] == 'No Commit - Late':
				ws['E%s'%(rownum+2)] = ws['E%s'%(rownum+2)].value + 1
			elif row[8] == 'No Commit - OK':
				ws['E%s'%(rownum+3)] = ws['E%s'%(rownum+3)].value + 1	
		elif row[4] == 'High':
			if row[8] == "Future Commit":
				ws['L%s'%(rownum)] = ws['L%s'%(rownum)].value + 1
			elif row[8] == 'Past Commit':
				ws['L%s'%(rownum+1)] = ws['L%s'%(rownum+1)].value + 1
			elif row[8] == 'No Commit - Late':
				ws['L%s'%(rownum+2)] = ws['L%s'%(rownum+2)].value + 1
			elif row[8] == 'No Commit - OK':
				ws['L%s'%(rownum+3)] = ws['L%s'%(rownum+3)].value + 1
		elif row[4] == 'Linedown':
			ws['P13'] = ws['P13'].value + 1
		elif row[4] == 'Safety':
			ws['P17'] = ws['P17'].value + 1

def get_rownum(row):
	if row[5] == 'Confirmed':
		return 7
	elif row[5] == 'Deferred':
		return 12
	elif row[5] == 'In Review':
		return 17

def main():
	db = Database()
	wb = Workbook()

	n = raw_input('1 Software\n2 Controls...')
	name = choose_query(int(n))
	db.generate_query(name)

	path = raw_input('Path...')
	db.establish_connection(path)
	db.stoplight()

	wbName = raw_input('Workbook name (\'somthing.xlsx\')...')
	wsName = raw_input('Previous worksheet name (\'Sheet1\')...')
	wsNewName = raw_input('New worksheet name (\'Sheet2\')...')
	[ws, wbook] = wb.load(wbName, wsName, wsNewName)
	analysis(db, ws, wbook)
	wb.close(wbook)

if __name__ == '__main__':
	main()

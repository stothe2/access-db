import pyodbc
import os.path
import datetime

class Database:
	'''repository for PR data'''

	def __init__(self):
		self.db = []
		self.queryName = ''
		self.pathName = ''
		self.goalHigh = ''
		self.goalMedLow = ''
		self.goalLinedown = ''
		self.goalSafety = ''
		self.hol = ''

	# return database path
	def path(self):
		return self.pathName

	# return current query choice
	def query(self):
		return self.queryName

	# return data array
	def data(self):
		return self.db

	def error(self, msg):
		print 'Error: ' + msg
		exit()

	def generate_query(self, name):
		self.queryName = '\'' + name + '\''
		self.queryStr = ("""SELECT dbo_tblChange.ECR_NO AS PR_Number, dbo_tblChange.Originator, [General - Latest Reviewer].Reviewer_Name, dbo_SolutionOwner.SolutionOwnerName, dbo_tblUrgency.Urgency, dbo_tblStatus.Status, dbo_tblChange.Submit_Date, dbo_tblChange.Effectivity_Date AS Planned_Complete_Date, dbo_tblChange.ECR_NO AS Stoplight, dbo_tblReason.Reason, dbo_CauseCode.CauseCode, dbo_tblSource.Source_Name
FROM (((((((dbo_tblChange INNER JOIN dbo_tblStatus ON dbo_tblChange.Status_ID = dbo_tblStatus.Status_ID) INNER JOIN dbo_tblUrgency ON dbo_tblChange.Urgency_ID = dbo_tblUrgency.Urgency_ID) INNER JOIN dbo_tblSource ON dbo_tblChange.Source_ID = dbo_tblSource.Source_ID) INNER JOIN [General - Latest Reviewer] ON dbo_tblChange.ECR_NO = [General - Latest Reviewer].PR) INNER JOIN dbo_tblReason ON dbo_tblChange.Reason_ID = dbo_tblReason.Reason_ID) INNER JOIN [Disposition-Cycle_Time_Goals] ON dbo_tblUrgency.Urgency = [Disposition-Cycle_Time_Goals].Urgency) LEFT JOIN dbo_SolutionOwner ON dbo_tblChange.SolutionOwnerID = dbo_SolutionOwner.SolutionOwnerID) LEFT JOIN dbo_CauseCode ON dbo_tblChange.CauseCodeID = dbo_CauseCode.ID
WHERE (((dbo_tblStatus.Status) In ('In Review','Confirmed','Deferred')))
GROUP BY dbo_tblChange.ECR_NO, dbo_tblChange.Originator, [General - Latest Reviewer].Reviewer_Name, dbo_SolutionOwner.SolutionOwnerName, dbo_tblUrgency.Urgency, dbo_tblStatus.Status, dbo_tblChange.Submit_Date, dbo_tblChange.Effectivity_Date, dbo_tblReason.Reason, dbo_CauseCode.CauseCode, dbo_tblSource.Source_Name, [General - Latest Reviewer].RECEIVED_DATE, [General - Latest Reviewer].Review_Org, dbo_tblChange.Customer, [General - Latest Reviewer].Review_Group, [General - Latest Reviewer].Disposition_Date, [Disposition-Cycle_Time_Goals].Goal
HAVING ((([General - Latest Reviewer].Review_Org)=%s))
ORDER BY [General - Latest Reviewer].Review_Org, dbo_tblUrgency.Urgency, dbo_tblChange.Effectivity_Date DESC;""" % self.queryName)

	def establish_connection(self, name):
		if not os.path.isfile(name):
			self.error('invalid path name!')
		self.pathName = name
		cnxn = pyodbc.connect(
			r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=%s' % self.pathName)
		cursor = cnxn.cursor()

		self.db = cursor.execute(self.queryStr).fetchall()
		self.goalHigh = cursor.execute("SELECT Goal FROM [Disposition-Cycle_Time_Goals] WHERE Urgency In ('High')").fetchone()[0]
		self.goalMedLow = cursor.execute("SELECT Goal FROM [Disposition-Cycle_Time_Goals] WHERE Urgency In ('Med/Low')").fetchone()[0]
		self.goalLinedown = cursor.execute("SELECT Goal FROM [Disposition-Cycle_Time_Goals] WHERE Urgency In ('Linedown')").fetchone()[0]
		self.goalSafety = cursor.execute("SELECT Goal FROM [Disposition-Cycle_Time_Goals] WHERE Urgency In ('Safety')").fetchone()[0]
		self.hol = cursor.execute('SELECT Holiday FROM Holidays').fetchall()

	def stoplight(self):
		count = 0
		dateArray = self.generate_date_array()
		for row in self.db:
			# Planned_Complete_Date = NULL?
			if not row[7]:
				# Submit_Date < myDate?
				if row[6] < dateArray[count]:
					row[8] = "No Commit - Late"
				else:
					 row[8] = "No Commit - OK"
			elif row[7] < datetime.datetime.today():
				row[8] = "Past Commit"
			else:
				row[8] = "Future Commit"
			count = count + 1

	def generate_date_array(self):
		dateArray = []
		goal = self.generate_goal_array()
		for item in goal:
			dateArray.append(self.deltaworkdays(item))
		return dateArray

	def generate_goal_array(self):
		goal = []
		for row in self.db:
			if row.Urgency == "Med/Low":
				goal.append(self.goalMedLow)
			elif row.Urgency == "High":
				goal.append(self.goalHigh)
			elif row.Urgency == "Linedown":
				goal.append(self.goalLinedown)
			elif row.Urgency == "Safety":
				goal.append(self.goalSafety)
		return goal

	def deltaworkdays(self, numDays):
		addNum = 0
		dayCount = 0

		if numDays >= 0:
			addNum = 1	# days in future
		else:
			addNum = -1	# days in past

		myDate = datetime.datetime.today()
		myDate = myDate.replace(hour=0, minute=0, second=0, microsecond=0)

		while (dayCount != abs(numDays)):

			myDate = myDate + datetime.timedelta(-addNum)
			#myDate = myDate.replace(day = myDate.day + 1)

			# execute if myDate not weekend
			if myDate.weekday() != 5 and myDate.weekday() != 6:
				# ??
				x = self.find(myDate)
				# execute if no match found
				if x is 0:
					dayCount = dayCount + 1
		return myDate

	# helper function for deltaworkdays
	def find(self, myDate):
		check = 0
		for row in self.hol:
			if row.Holiday == myDate: # check syntax!!!!
				check = 1
		return check

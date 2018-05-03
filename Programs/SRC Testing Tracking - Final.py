import pandas as pd
import os
import numpy as np

import sys
sys.path.append("//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Python Functions")
from Merge_Like_Stata import MergeLikeStata as mstata

import getpass
import glob

from IPython.display import display, HTML

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time


# Import SRC's Tracker
#SRCGoogleDocsTracker = pd.read_csv('https://docs.google.com/spreadsheets/d/14fHwhs9SdHleHr-CoBC4j3RfGYpscU7y4og-Dvy2BvY/export?gid=1711374707&format=csv',index_col=0)
SRCGoogleDocsTracker = pd.read_csv('https://docs.google.com/spreadsheets/d/14fHwhs9SdHleHr-CoBC4j3RfGYpscU7y4og-Dvy2BvY/export?gid=1711374707&format=csv')
SRCGoogleDocsTracker.head()

# Locate directory whene SRC Rosters are
user = getpass.getuser()
rootdir = 'C:/Users/'+user+'/inSync Share/CLI Study (AIR SRC shared folder)/Year 2 2017-18/Lists of Selected Students'

# Create A List Of Student Rosters
student_roster_list=[]
for subdir, dirs, files in os.walk(rootdir):
	for file in files:
		if file.find('To Print')==-1 and file.find('xlsx')!=-1 and file.find('~')==-1 and file.find('Bad Conflict')==-1 and file.find('(Optouts Removed)')==-1 :
			student_roster_list.append(os.path.join(subdir, file))
print(len(student_roster_list))


# Create A Data Frame Containign All Student Rosters
student_rosters = pd.DataFrame()
for roster in student_roster_list:
	for grade in pd.ExcelFile(roster).sheet_names:
		dataframe = pd.read_excel(roster, sheet_name=grade, header=None)
		dataframe['Grade']=grade
		student_rosters = pd.concat([student_rosters, dataframe], ignore_index=True)

# Drop Non-Id Rows
student_rosters[4] = student_rosters[4].map(lambda x: str(x).strip())
student_rosters['keep'] = student_rosters.apply(lambda x: 1 if len(x[4])==6 else 0, axis=1)
student_rosters = student_rosters[student_rosters['keep'].isin([1])]
print(len(student_rosters))


# Table Of Students By District
student_rosters['column']="Count"
student_rosters['column2']=1
pd.crosstab(student_rosters[0], student_rosters.column, margins=True)
pd.crosstab(student_rosters['Grade'], student_rosters.column, margins=True)
table = student_rosters.pivot_table(values='column2', index=[10],columns=[0], aggfunc=np.sum)

# Check Duplicates
dupcheckvars=[4]
student_rosters.groupby(dupcheckvars).ngroups==len(student_rosters)


# rename columns
student_rosters.rename(columns={0:'District'}, inplace=True)
student_rosters.rename(columns={1:'School'}, inplace=True)


# Collapse To Student By School By Grade Level Counts
student_rosters['AdminCode10'] = student_rosters.apply(lambda x: 1 if x[10] ==10 else 0, axis = 1)

nottested = student_rosters.loc[student_rosters['AdminCode10'] == 0]
tested = student_rosters.loc[student_rosters['AdminCode10'] == 1]

nottested = nottested.groupby([ 'District', 'School','Grade']).size().to_frame(name = 'AdminCodeNot10').reset_index()
tested = tested.groupby([ 'District', 'School','Grade']).size().to_frame(name = 'AdminCode10').reset_index()

summary = mstata.stata_merge(tested,nottested,['District','School','Grade'],'1:1')

# Recode grades for merge with SRC rosters
def graderedef(x):
	if x['Grade'] == '1st':
		return 1
	if x['Grade'] == '2nd':
		return 2
	if x['Grade'] == '3rd':
		return 3

summary['Grade'] = summary.apply(lambda x: graderedef(x), axis = 1)
summary = summary.drop(columns=['_merge'])
SRCGoogleDocsTracker = SRCGoogleDocsTracker.dropna(axis=0, subset=['District','School','Grade'],how='all')


Compare = mstata.stata_merge(summary,SRCGoogleDocsTracker,['District','School','Grade'],'1:1')

Compare['Difference'] = Compare['# Assessed (Participated in Part 1 and Part 2)'] - Compare['AdminCode10']

Compare['Roster SRC Tracker Comparison'] =  Compare.apply(lambda x: 'School/Grade Matches Within 5 GRADE' if x['Difference']>-5 and x['Difference']<5 else 0, axis = 1)
Compare['Roster SRC Tracker Comparison'] =  Compare.apply(lambda x: 'School/Grade Off By 5 Or More GRADE' if x['Difference']<-5 or x['Difference']>5 else x['Roster SRC Tracker Comparison'], axis = 1)
Compare['Roster SRC Tracker Comparison'] =  Compare.apply(lambda x: 'School/Grade Not Recoded In Rosters' if pd.isnull(x['AdminCode10']) and pd.notnull(x['# Assessed (Participated in Part 1 and Part 2)']) else x['Roster SRC Tracker Comparison'], axis = 1)
Compare['Roster SRC Tracker Comparison'] =  Compare.apply(lambda x: 'School/Grade Not Recoded On GoogleDoc' if pd.notnull(x['AdminCode10']) and pd.isnull(x['# Assessed (Participated in Part 1 and Part 2)']) else x['Roster SRC Tracker Comparison'], axis = 1)
Compare['Roster SRC Tracker Comparison'] =  Compare.apply(lambda x: 'School/Grade Not Yet Tested' if pd.isnull(x['AdminCode10']) and pd.isnull(x['# Assessed (Participated in Part 1 and Part 2)']) else x['Roster SRC Tracker Comparison'], axis = 1)



# Create Summary Tables
Table1 = Compare.groupby([ 'District','Grade']).sum()['AdminCode10'].to_frame(name = 'Tested Students').reset_index().pivot_table(index=['Grade'], columns=['District'], margins=True, aggfunc=np.sum)
Table1=Table1.astype(int)

Table2 = Compare.groupby([ 'District','Grade']).sum()['AdminCodeNot10'].to_frame(name = 'Untested Students').reset_index().pivot_table(index=['Grade'], columns=['District'], margins=True, aggfunc=np.sum)
Table2=Table2.astype(int)

Table3 = Compare.groupby([ 'District','Roster SRC Tracker Comparison']).size().to_frame(name = 'Tracking Log').reset_index().pivot_table(index=['Roster SRC Tracker Comparison'], columns=['District'], margins=True, aggfunc=np.sum)
Table3=Table3.astype(int)

text = """\
<html>
  <head></head>
  <body>
    <p>Hi!<br><br>
       Below are tables describing the data collection process <br><br>
       Cheers!, <br><br>
       John 
    </p>
  </body>
</html>
"""

htmltables = text
htmltables += Table1.to_html()
htmltables += "<br><br>"
htmltables += Table2.to_html()
htmltables += "<br><br>"
htmltables += Table3.to_html()
htmltables += "<br><br>"

# Email
list_of_files = glob.glob('//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Year 2/Data Collection - Students/GRADE Tracking/Summaries/*')
latest_file = max(list_of_files, key=os.path.getctime)
latest_filename=latest_file.replace("//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Year 2/Data Collection - Students/GRADE Tracking/Summaries\\","")

def SendNotificaiton ():

	msg = MIMEMultipart("alternative", None, [MIMEText(htmltables,'html')])
	msg['From'] = "cliscaleup@gmail.com"
	#msg['To'] = "jmeakin@air.org"
	msg['To'] = "ntucker-bradway@air.org,dmsmith@air.org, jmeakin@air.org, kdrummond@air.org, chery@schoolreadinessconsulting.com, davis@schoolreadinessconsulting.com"
	msg['Subject'] = "SRC Testing Tracking"

	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login("cliscaleup@gmail.com", "CLIstudy4ever!")
	text = msg.as_string()
	server.sendmail("cliscaleup@gmail.com", msg["To"].split(","), text)
	server.quit()

SendNotificaiton()
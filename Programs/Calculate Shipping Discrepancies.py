# Imports For Data Management
import pandas as pd
import os
import numpy as np
import getpass
import glob

# Imports For Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time

# Imports For Merge
import sys
sys.path.append("//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Python Functions")
from Merge_Like_Stata import MergeLikeStata as mstata

# Import Only For Jupyter Notebook Visualization
from IPython.display import display, HTML

def grab_all_xlsx(path):
    list_of_files = glob.glob(path+'/*')
    df = pd.DataFrame()
    for file in list_of_files:
        if file.find('xlsx')!=-1 and file.find('~')==-1 :
            data = pd.read_excel(file)
            df = pd.concat([df, data], ignore_index=True)
    return  df


# Find Paths
user = getpass.getuser()
SRC_DIR = 'C:/Users/'+user+'/inSync Share/CLI Study (AIR SRC shared folder)/Year 2 2017-18/Shipping Logs'
AIR_DIR = '//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Year 2/Data Collection - Students/GRADE Tracking'
AIR_DIR2 = '//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Year 2/Data Collection - Students/GRADE Tracking/Summaries/'

# Apply Function To Create DataFrames
SRC_GRADE_trackers = grab_all_xlsx(SRC_DIR)
AIR_GRADE_trackers = grab_all_xlsx(AIR_DIR)

# Drop Blank Rows
SRC_GRADE_trackers = SRC_GRADE_trackers.dropna(axis=0, subset=['SchoolStudyID_1718','Grade_1718'],how='all')
AIR_GRADE_trackers = AIR_GRADE_trackers.dropna(axis=0, subset=['SchoolStudyID_1718','Grade_1718'],how='all')

# Merge
Compare_AIR_SRC = mstata.stata_merge(SRC_GRADE_trackers,AIR_GRADE_trackers,['SchoolStudyID_1718','Grade_1718'],'1:1')


# Calculate differneces
Compare_AIR_SRC['Difference'] = Compare_AIR_SRC['NumGradeSentToAIR'] - Compare_AIR_SRC['NumGRADERcvd']
#Compare_AIR_SRC = Compare_AIR_SRC.dropna(axis=0, subset=['NumGradeSentToAIR','NumGRADERcvd'],how='all')


# School Grade
Compare_AIR_SRC["School_Grade"] = Compare_AIR_SRC["School_1718_y"].map(str) +" - "+ Compare_AIR_SRC["Grade_1718"].map(str)
# Non-Recieved/Tracked

Compare_AIR_SRC["District"] = Compare_AIR_SRC["District_x"]
Compare_AIR_SRC = Compare_AIR_SRC[['District','School_Grade','NumGradeSentToAIR','NumGRADERcvd','Difference']].copy()

# Flag Columns
Compare_AIR_SRC["Matches Within 5 GRADE"] =  Compare_AIR_SRC.apply(lambda x: x["School_Grade"] if x['Difference']>-5 and x['Difference']<5 else 0, axis = 1)
Compare_AIR_SRC["Off By More Than 5 GRADE"] =  Compare_AIR_SRC.apply(lambda x: x["School_Grade"] if x['Difference']<-5 or x['Difference']>5 else 0, axis = 1)
Compare_AIR_SRC["Not Recieved Or Recoreded"] =  Compare_AIR_SRC.apply(lambda x: x["School_Grade"] if pd.notnull(x['NumGradeSentToAIR']) and pd.isnull(x['NumGRADERcvd']) else 0, axis = 1)
Compare_AIR_SRC["Not Tracked By SRC"] =  Compare_AIR_SRC.apply(lambda x: x["School_Grade"] if pd.isnull(x['NumGradeSentToAIR']) and pd.notnull(x['NumGRADERcvd']) else 0, axis = 1)
Compare_AIR_SRC["Not Tested Yet"] =  Compare_AIR_SRC.apply(lambda x: x["School_Grade"] if pd.isnull(x['NumGradeSentToAIR']) and pd.isnull(x['NumGRADERcvd']) else 0, axis = 1)


#Create dataframes out of eachrelevant colum for export to excel
Matches = Compare_AIR_SRC[['District','Matches Within 5 GRADE']].copy()
Matches = Matches[~(Matches == 0).any(axis=1)]
Issue1 = Compare_AIR_SRC[['District','Off By More Than 5 GRADE','Difference']].copy()
Issue1 = Issue1[~(Issue1 == 0).any(axis=1)]
Issue2 = Compare_AIR_SRC[['District','Not Recieved Or Recoreded','NumGradeSentToAIR']].copy()
Issue2 = Issue2[~(Issue2 == 0).any(axis=1)]
Issue3 = Compare_AIR_SRC[['District','Not Tracked By SRC']].copy()
Issue3 = Issue3[~(Issue3 == 0).any(axis=1)]
Issue4 = Compare_AIR_SRC[['District','Not Tested Yet']].copy()
Issue4 = Issue4[~(Issue4 == 0).any(axis=1)]

# Check
check = pd.concat([Matches, Issue1], ignore_index=True)
check = pd.concat([check, Issue2], ignore_index=True)   
check = pd.concat([check, Issue3], ignore_index=True)       
check = pd.concat([check, Issue4], ignore_index=True)
print(len(check))

todaysdate = time.strftime("%m-%d-%Y")
# Write to file
writer = pd.ExcelWriter(AIR_DIR2+'GRADE Tracking Summary'+todaysdate+'.xlsx', engine='xlsxwriter')
workbook  = writer.book

Matches.to_excel(writer, sheet_name='Matches', index=False)
Issue1.to_excel(writer, sheet_name='Discrepancy', index=False)
Issue2.to_excel(writer, sheet_name='NotRcvd', index=False)
Issue3.to_excel(writer, sheet_name='SRC_Error', index=False)
Issue4.to_excel(writer, sheet_name='NotTested', index=False)

# https://stackoverflow.com/questions/43425944/pandas-fromat-column-multiple-sheets?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
# http://xlsxwriter.readthedocs.io/format.html
writer.save()


# Email
list_of_files = glob.glob('//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Year 2/Data Collection - Students/GRADE Tracking/Summaries/*')
latest_file = max(list_of_files, key=os.path.getctime)
latest_filename=latest_file.replace("//tx1cifs/tx1data/Austin Share/CLI ScaleUp Grant/Year 2/Data Collection - Students/GRADE Tracking/Summaries\\","")

def SendNotificaiton ():
    
    fromaddr = "cliscaleup@gmail.com"
    msg = MIMEMultipart()
     
    msg['From'] = fromaddr
    #msg['To'] = "jmeakin@air.org"
    msg['To'] = "ntucker-bradway@air.org,dmsmith@air.org, jmeakin@air.org, kpate@air.org"
    msg['Subject'] = "AIR-SRC Shipping Tracking"
     
    body = "Hi,\n\nPlease see attachment for summary of SRC/AIR shipment tracking. \n \nCheers \n \nJohn"
     
    msg.attach(MIMEText(body, 'plain'))

    filename = latest_filename
    attachment = open(latest_file, "rb")
     
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
     
    msg.attach(part)
     
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, "CLIstudy4ever!")
    text = msg.as_string()
    server.sendmail(fromaddr, msg["To"].split(","), text)
    server.quit()

SendNotificaiton()





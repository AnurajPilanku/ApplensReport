'''
Author   :  AnurajPilanku
Use Case :  SMO Applens CAC Automation

'''

import numpy as np
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from os.path import basename
import openpyxl
import pandas as pd
import sys
from datetime import date
import datetime

maildatapath = sys.argv[1]
associatedetailsPath = sys.argv[2]
MailRecipientExcel = sys.argv[3]
nadatapath = r"\\acprd01\E\3M_CAC\ApplensSMO\Incomplete_Details.xlsx"
primarycolumn = 'Associate ID'

recipient = pd.read_excel(MailRecipientExcel, sheet_name='recipients', engine="openpyxl")
groups = pd.read_excel(MailRecipientExcel, sheet_name='groups', engine="openpyxl")
# converting column in dataframe to list and droping Blank\Nan\None values
SMO = ",".join(groups[groups.columns[0]].dropna().tolist())
MSOADMLead = ",".join(groups[groups.columns[1]].dropna().tolist())
to_xl = ",".join(recipient[recipient.columns[0]].dropna().tolist())  # +","+SMO
cc_xl = ",".join(recipient[recipient.columns[1]].dropna().tolist())  # +","+MSOADMLead
bcc_xl = ",".join(recipient[recipient.columns[2]].dropna().tolist())

# Associate Details
Associatecollection = pd.read_excel(associatedetailsPath, engine="openpyxl")
# remove extra spaces in column headers
Associatecollection.columns = Associatecollection.columns.str.strip()
AssociateDetails = Associatecollection[['Associate ID', 'Location', 'ProcessArea']]

# Project Summary Data

projectData = pd.read_excel(maildatapath, sheet_name='Project Summary', engine="openpyxl")
projectData.columns = projectData.columns.str.strip()
ReqData = projectData[['ProjectName', 'Project Effort Compliance% (All)', 'Project Associate Compliance% (All)']]
# Three_MProjectData=ReqData['ProjectName'].str.startswith('3M',na=False)#.contains('3M',case=False,na=False)
Three_MProjectDataraw = ReqData[ReqData['ProjectName'].str.startswith('3M', na=False) == True]
# rounding,changing datatype,replacing substring
decimals = 0
# y['Project Effort Compliance% (All)']=y['Project Effort Compliance% (All)'].apply(lambda x:round(x,decimals)).astype(int).astype(str)+"%"#.replace('.0','%',regex=True)

Three_MProjectDataraw['Project Effort Compliance% (All)'] = Three_MProjectDataraw[
                                                                'Project Effort Compliance% (All)'].apply(
    lambda x: round(x, decimals)).astype(int).astype(str) + "%"  # .replace('.0','%',regex=True)
Three_MProjectDataraw['Project Associate Compliance% (All)'] = Three_MProjectDataraw[
                                                                   'Project Associate Compliance% (All)'].apply(
    lambda x: round(x, decimals)).astype(int).astype(
    str) + "%"  # .replace('.0','%',regex=True).replace('%0%','100%',regex=True)

# Assosiate Summary Data
AssociateData = pd.read_excel(maildatapath, sheet_name='Associate Summary', engine="openpyxl")
AssociateData.columns = AssociateData.columns.str.strip()
# filtering with partial string in values in a column
Three_MAssociateData = AssociateData[
    AssociateData['Projectname'].str.startswith('3M', na=False) == True]  # .contains('3M',case=False,na=False)
AssociateReqData = Three_MAssociateData[
    ['Projectname', 'EmployeeID', 'EmployeeName', 'Associate Allocation(In FTE)', 'Available Hours', 'Actual Effort',
     '[Effort TS Compliance%]']]
# renaming column
AssociateReqData.rename(columns={'EmployeeID': 'Associate ID'}, inplace=True)
# vlookup
join = pd.merge(AssociateReqData, AssociateDetails, on=primarycolumn, how="left")
# sorting
ascending = join.sort_values(by=['[Effort TS Compliance%]'], ascending=True)
# filtering non #NA
filtervalue = None
na_removed = ascending[(ascending['ProcessArea'] != filtervalue) & (ascending['Location'] != filtervalue)]

# preparing blank Associate Details
comdata = join.sort_values(by=['[Effort TS Compliance%]'], ascending=True)
comdata.dropna(subset=['ProcessArea', 'Location'], how='all', inplace=True)
# getting data\row in a  dataframe which is not in the second dataframe
noncommondata = na_removed.merge(comdata, how='outer', indicator=True).loc[lambda x: x['_merge'] == 'left_only']
# AssociateWithInadequateInformationraw= noncommondata[['Associate ID', 'EmployeeName']]
AssociateWithInadequateInformationraw = noncommondata[
    ['Projectname', 'Associate ID', 'EmployeeName', 'Associate Allocation(In FTE)', 'Available Hours', 'Actual Effort',
     '[Effort TS Compliance%]']]
#AssociateWithInadequateInformationraw.to_excel(nadatapath, index=False)

# onsite
onsiteraw = na_removed[na_removed['Location'] == 'Onsite']
# offshore
offshoreraw = na_removed[na_removed['Location'] == 'Offshore']

# onsite,offshore-Filter Red to Amber(yellow)
# remove duplicate\identical rows
onsite = onsiteraw[onsiteraw['[Effort TS Compliance%]'] < 80].drop_duplicates(subset=None, keep="first", inplace=False)
offshore = offshoreraw[offshoreraw['[Effort TS Compliance%]'] < 80].drop_duplicates(subset=None, keep="first",
                                                                                    inplace=False)
AssociateWithInadequateInformation = AssociateWithInadequateInformationraw[
    AssociateWithInadequateInformationraw['[Effort TS Compliance%]'] < 80].drop_duplicates(subset=None, keep="first",
                                                                                           inplace=False)

# collecting mail ids of onshore and offshore
onsiterecivermailaddress = ",".join(
    list(set(pd.merge(onsite, Associatecollection, on=primarycolumn, how="left")['Mail ID'].dropna().tolist())))
offshorerecivermailaddress = ",".join(
    list(set(pd.merge(offshore, Associatecollection, on=primarycolumn, how="left")['Mail ID'].dropna().tolist())))
AssociateWithInadequateInformationmailaddress = ",".join(
    list(set(pd.merge(AssociateWithInadequateInformation, Associatecollection, on=primarycolumn, how="left")[
                 'Mail ID'].dropna().tolist())))

# groupby ONSITE
GroupyByProcessAreaOnsite = pd.DataFrame({"count": onsite['ProcessArea'].value_counts()})
# converting index to columns
GroupyByProcessAreaOnsite["processarea"] = GroupyByProcessAreaOnsite.index
# droping index
GroupyByProcessAreaOnsite.reset_index(drop=True, inplace=True)
# changing column index
GroupyByProcessAreaOnsiteorg = GroupyByProcessAreaOnsite[["processarea", "count"]]
GroupyByProcessAreaOnsiteorg = GroupyByProcessAreaOnsiteorg.append(
    {'processarea': 'Grand Total', 'count': sum(GroupyByProcessAreaOnsiteorg['count'].tolist())}, ignore_index=True)

##groupby OFFSHORE
GroupyByProcessAreaOffshore = pd.DataFrame({"count": offshore['ProcessArea'].value_counts()})
GroupyByProcessAreaOffshore["processarea"] = GroupyByProcessAreaOffshore.index
GroupyByProcessAreaOffshore.reset_index(drop=True, inplace=True)
GroupyByProcessAreaOffshoreorg = GroupyByProcessAreaOffshore[["processarea", "count"]]
GroupyByProcessAreaOffshoreorg = GroupyByProcessAreaOffshoreorg.append(
    {'processarea': 'Grand Total', 'count': sum(GroupyByProcessAreaOffshoreorg['count'].tolist())}, ignore_index=True)

# GroupyByProcessAreaOnsiteorg.to_excel(outputpath+"//"+"pivoton.xlsx",index=False)
# GroupyByProcessAreaOffshoreorg.to_excel(outputpath+"//"+"pivotoff.xlsx",index=False)
# onsite.to_excel(outputpath+"//"+"on.xlsx",index=False)
# offshore.to_excel(outputpath+"//"+"of.xlsx",index=False)

# html table
# Headers
colheaderprojectsummary = 'ProjectName,Project Effort Compliance% (All),Project Associate Compliance% (All)'.split(",")
AssocHeader = 'Projectname,EmployeeID,EmployeeName,Associate Allocation(In FTE),Available Hours,Actual Effort,[Effort TS Compliance%],Location,Process Area'.split(
    ",")
PivotHeaders = 'Process area,Count of EmployeeName'.split(",")
InadequeteHeader='Projectname,EmployeeID,EmployeeName,Associate Allocation(In FTE),Available Hours,Actual Effort,[Effort TS Compliance%]'.split(
    ",")

# Header preparation for Project Summary
onestr = str()
prjc = '<td bgcolor="#99B2FF">{celval}</td>' + "\n"
for header in colheaderprojectsummary:
    onestr += prjc.format(celval=header)
colheaderprojectheadhtml = '<tr>{codevb}</tr>'.format(codevb=onestr)

# Header preparation for Associate Summary
onestr1 = str()
prjc1 = '<td bgcolor="#D9D9BF">{celval}</td>' + "\n"
for header in AssocHeader:
    onestr1 += prjc1.format(celval=header)
colheaderAssociateheadhtml = '<tr>{codevb}</tr>'.format(codevb=onestr1)

# Header preparation for Pivot
onestr2 = str()
prjc2 = '<td bgcolor="#87FEF8">{celval}</td>' + "\n"
for header in PivotHeaders:
    onestr2 += prjc2.format(celval=header)
colheaderpivotheadhtml = '<tr>{codevb}</tr>'.format(codevb=onestr2)

#Header preparation for Inadequete Associates
onestr4 = str()
prjc1 = '<td bgcolor="#D9D9BF">{celval}</td>' + "\n"
for header in InadequeteHeader:
    onestr4 += prjc1.format(celval=header)
colheaderInadequeteheadhtml= '<tr>{codevb}</tr>'.format(codevb=onestr4)


# project Summary
initiation = colheaderprojectheadhtml
td = '<td bgcolor="{ColorCode}" style="text-align:center;" >{tdval}</td>' + "\n"
for row in range(0, Three_MProjectDataraw.shape[0]):
    initiation += '<tr style="text-align:center;" >' + "\n"
    for col in range(0, Three_MProjectDataraw.shape[1]):
        if col == 2:
            if int(str(list(Three_MProjectDataraw.iloc[:, col])[row]).replace("%", "")) <= 50:
                initiation += td.format(tdval=str(list(Three_MProjectDataraw.iloc[:, col])[row]), ColorCode="##FF0000")
            elif 51 <= int(str(list(Three_MProjectDataraw.iloc[:, col])[row]).replace("%", "")) <= 80:
                initiation += td.format(tdval=str(list(Three_MProjectDataraw.iloc[:, col])[row]), ColorCode="#FFFF00")
            elif 81 <= int(str(list(Three_MProjectDataraw.iloc[:, col])[row]).replace("%", "")) <= 100:
                initiation += td.format(tdval=str(list(Three_MProjectDataraw.iloc[:, col])[row]), ColorCode="#8DFD53")

        else:
            initiation += td.format(tdval=str(list(Three_MProjectDataraw.iloc[:, col])[row]), ColorCode="#FFFFFF")
    initiation += "</tr>" + '\n'
# print(initiation)

# Assosiate Summary
# OFFSHORE HTML
AssociateInitiationOffshore = colheaderAssociateheadhtml  # str()
td = '<td bgcolor="{ColorCode}" style="text-align:center;" >{tdval}</td>' + "\n"
for row in range(0, offshore.shape[0]):
    AssociateInitiationOffshore += '<tr style="text-align:center;" >' + "\n"
    for col in range(0, offshore.shape[1]):
        if col == 6:
            print()
            if list(offshore.iloc[:, col])[row] <= 50:
                AssociateInitiationOffshore += td.format(tdval=str(list(offshore.iloc[:, col])[row]),
                                                         ColorCode="##FF0000")
            elif 50 <= list(offshore.iloc[:, col])[
                row] <= 80:  # list(offshore.iloc[:,col])[row]>50 & list(offshore.iloc[:,col])[row]<80:
                AssociateInitiationOffshore += td.format(tdval=str(list(offshore.iloc[:, col])[row]),
                                                         ColorCode="#FFFF00")
        else:
            AssociateInitiationOffshore += td.format(tdval=str(list(offshore.iloc[:, col])[row]), ColorCode="#FFFFFF")
    AssociateInitiationOffshore += "</tr>" + '\n'
# print(AssociateInitiationOffshore)

# Pivot

PivotInitiationOffshore = colheaderpivotheadhtml  # str()
td = '<td bgcolor="{ColorCode}" style="text-align:center;" >{tdval}</td>' + "\n"
for row in range(0, GroupyByProcessAreaOffshoreorg.shape[0]):
    PivotInitiationOffshore += '<tr style="text-align:center;" >' + "\n"
    for col in range(0, GroupyByProcessAreaOffshoreorg.shape[1]):
        if row == GroupyByProcessAreaOffshoreorg.shape[0] - 1:
            PivotInitiationOffshore += td.format(tdval=str(list(GroupyByProcessAreaOffshoreorg.iloc[:, col])[row]),
                                                 ColorCode="#CF9FFF")
        else:
            PivotInitiationOffshore += td.format(tdval=str(list(GroupyByProcessAreaOffshoreorg.iloc[:, col])[row]),
                                                 ColorCode="#FFFFFF")
    PivotInitiationOffshore += "</tr>" + '\n'
# print(PivotInitiationOffshore)

# ONSITE HTML
AssociateInitiationOnsite = colheaderAssociateheadhtml  # str()
td = '<td bgcolor="{ColorCode}" style="text-align:center;" >{tdval}</td>' + "\n"
for row in range(0, onsite.shape[0]):
    AssociateInitiationOnsite += '<tr style="text-align:center;" >' + "\n"
    for col in range(0, onsite.shape[1]):
        if col == 6:
            print()
            if list(onsite.iloc[:, col])[row] <= 50:
                AssociateInitiationOnsite += td.format(tdval=str(list(onsite.iloc[:, col])[row]), ColorCode="##FF0000")
            elif 50 <= list(onsite.iloc[:, col])[
                row] <= 80:  # list(onsite.iloc[:,col])[row]>50 & list(onsite.iloc[:,col])[row]<80:
                AssociateInitiationOnsite += td.format(tdval=str(list(onsite.iloc[:, col])[row]), ColorCode="#FFFF00")
        else:
            AssociateInitiationOnsite += td.format(tdval=str(list(onsite.iloc[:, col])[row]), ColorCode="#FFFFFF")
    AssociateInitiationOnsite += "</tr>" + '\n'
# print(AssociateInitiationOnsite)

# Pivot

PivotInitiationOnsite = colheaderpivotheadhtml  # str()
td = '<td bgcolor="{ColorCode}" style="text-align:center;" >{tdval}</td>' + "\n"
for row in range(0, GroupyByProcessAreaOnsiteorg.shape[0]):
    PivotInitiationOnsite += '<tr style="text-align:center;" >' + "\n"
    for col in range(0, GroupyByProcessAreaOnsiteorg.shape[1]):
        if row == GroupyByProcessAreaOnsiteorg.shape[0] - 1:
            PivotInitiationOnsite += td.format(tdval=str(list(GroupyByProcessAreaOnsiteorg.iloc[:, col])[row]),
                                               ColorCode="#CF9FFF")
        else:
            PivotInitiationOnsite += td.format(tdval=str(list(GroupyByProcessAreaOnsiteorg.iloc[:, col])[row]),
                                               ColorCode="#FFFFFF")
    PivotInitiationOnsite += "</tr>" + '\n'

# print(PivotInitiationOffshore)


#ASSOCIATE DETAILS INADEQUETE
InadequeteInitiation = colheaderInadequeteheadhtml  # str()
td = '<td bgcolor="{ColorCode}" style="text-align:center;" >{tdval}</td>' + "\n"
for row in range(0, AssociateWithInadequateInformation.shape[0]):
    InadequeteInitiation += '<tr style="text-align:center;" >' + "\n"
    for col in range(0, AssociateWithInadequateInformation.shape[1]):
        if col == 6:
            print()
            if list(AssociateWithInadequateInformation.iloc[:, col])[row] <= 50:
                InadequeteInitiation += td.format(tdval=str(list(AssociateWithInadequateInformation.iloc[:, col])[row]),
                                                         ColorCode="##FF0000")
            elif 50 <= list(AssociateWithInadequateInformation.iloc[:, col])[
                row] <= 80:  # list(AssociateWithInadequateInformation.iloc[:,col])[row]>50 & list(AssociateWithInadequateInformation.iloc[:,col])[row]<80:
                InadequeteInitiation += td.format(tdval=str(list(AssociateWithInadequateInformation.iloc[:, col])[row]),
                                                         ColorCode="#FFFF00")
        else:
            InadequeteInitiation += td.format(tdval=str(list(AssociateWithInadequateInformation.iloc[:, col])[row]), ColorCode="#FFFFFF")
    InadequeteInitiation += "</tr>" + '\n'

currentMonth = datetime.datetime.now().strftime("%B")


def applensmail(onoff, associateData, pivotdata ,recieveraddress):
    greeting = "Hi All"
    bodysentence = "We need immediate attention to capture the efforts in Applens."
    From = 'USSACPrd@mmm.com'
    reciever = to_xl#"P.Anuraj@cognizant.com" #+ "," +
    carboncopy = cc_xl#"ac5qdzz@mmm.com"#
    blindcarboncopy = bcc_xl#"ac5qdzz@mmm.com"#
    subject = "Applens compliance for {CurrentMonth} {onoroff}".format(CurrentMonth=currentMonth,
                                                                       onoroff=onoff)  # *****1
    attachments = ""
    mailfontstyle = "Cambria"
    html_file = '''<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Title</title>
    </head>
    <body style="font-family:{bodyfont}">
     <br/><img src='cid:image1'<br/>
      <br>
      <br>
      <br /><font face='{bodyfont}'>'''.format(bodyfont=mailfontstyle) + greeting + ''' </a></font><br/>
      <br /><font face='{bodyfont}'>'''.format(bodyfont=mailfontstyle) + bodysentence + ''' </a></font><br/>
      <br>
      <br>
      <br /><font face='{bodyfont}'>Project Summary </a></font><br/>
      <br><br/>
    <div style="overflow-x:auto;">
        <style>
            body{
            text-align:center;
            }
            table{
            border-collapse:collapse;}
            th,td{
            border: 1px solid black}
            th,td{
            padding:1px}

        </style>
        </style>
            <table>''' + initiation + '''</table>
    </div>
    <br /><font face='{bodyfont}'>Overview</a></font><br/>
    <br><br/>
    <div>
    <table  border="2pxsingleblack">''' + pivotdata + '''</table>
    </div>
    <br /><font face='{bodyfont}'>Detailed View</a></font><br/>
    <br><br/>
    <div>
    <table  border="2pxsingleblack">''' + associateData + '''</table>
    </div>

    <br /><font face='{bodyfont}'>Regards </a></font><br/>
    <br /><font face='{bodyfont}'>3M Automation Center Team </a></font><br/>
    <br>
    <br>
    <br/><img src='cid:image3'<br/>
    </body>
    </html>'''.format(bodyfont=mailfontstyle)

    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = From
    msgRoot['Cc'] = carboncopy
    msgRoot['To'] = reciever  +","+recieveraddress
    msgRoot['Bcc'] = blindcarboncopy
    msgRoot.preamble = '====================================================='
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)
    msgText = MIMEText('Please find ')
    msgAlternative.attach(msgText)
    msgText = MIMEText(html_file, 'html')
    msgAlternative.attach(msgText)
    msgAlternative.attach(msgText)
    fp = open(r"\\acprd01\3M_CAC\EDI_Ageing\head.png", 'rb')
    # fp2 = open(sys.argv[7], 'rb')#"//acdev01/3M_CAC/IPM_FSM/Mail_elements/new.png"
    fp3 = open(r"\\acprd01\3M_CAC\EDI_Ageing\footer.png", 'rb')
    msgImage = MIMEImage(fp.read())
    # msgImage1 = MIMEImage(fp2.read())
    msgImage2 = MIMEImage(fp3.read())
    fp.close()
    fp3.close()
    msgImage.add_header('Content-ID', '<image1>')
    msgImage2.add_header('Content-ID', '<image3>')
    msgRoot.attach(msgImage)
    msgRoot.attach(msgImage2)
    filepaths = [attachments]
    # for f in filepaths:
    # with open(f, "rb") as file:
    # part = MIMEApplication(file.read(), Name=basename(f))
    # part["Content-Disposition"] = 'attachment;filename="%s"' % basename(f)
    # msgRoot.attach(part)
    smtp = smtplib.SMTP()
    smtp.connect("mailserv.mmm.com")
    # smtp.sendmail(From,To, msgRoot.as_string())
    smtp.send_message(msgRoot)
    smtp.quit()
    print("Email is sent successfully")


applensmail("Offshore", AssociateInitiationOffshore, PivotInitiationOffshore ,offshorerecivermailaddress)
applensmail("Onsite", AssociateInitiationOnsite, PivotInitiationOnsite,offshorerecivermailaddress)
applensmail("Associates with Inadequete Information"+" "+str(datetime.datetime.today())[:10],InadequeteInitiation,"Associate with Inadequate Location and Process Area Information",offshorerecivermailaddress)

print("success")





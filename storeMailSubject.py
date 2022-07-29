#ADM Associate Compliance Report_19_July_2022_Interim
#ADM Associate Compliance Report - Jun-2022 Monthly Final

'''
Auther       : AnurajPilanku
Use case     : SMO Applens
Code Utility : Save mail subject in activity Attribute
'''
import datetime
import sys
weeknum=datetime.datetime.now().day
month=datetime.datetime.now().strftime("%B")
year=datetime.datetime.now().year
if weeknum not in [5,'5']:
    InterimORFinal='Interim'
else:
    InterimORFinal = 'Monthly_Final'
FirstData='ADM Associate Compliance Report_{weeknum}_{month}_{year}_{InterimORFinal}'.format(weeknum=weeknum,month=month,year=year,InterimORFinal=InterimORFinal)

applenssubject=dict()
applenssubject['applenssubject']=FirstData
output = {'output':applenssubject, 'additional_attributes': applenssubject}
sys.stdout.write(str(output) + '\n')


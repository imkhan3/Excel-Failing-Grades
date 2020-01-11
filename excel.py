import pandas as pd
import numpy as np
# denote spreadsheets
excel_file_1 = "T1 Failure Report.xlsx"
excel_file_2 = 'PR3 Failure Report.xlsx'

df1 = pd.read_excel(excel_file_1, sheet_name='T1 - FAILING by ALPHA')
df2 = pd.read_excel(excel_file_2, sheet_name='PR3 FAILING by ALPHA REPORT')

#Rename Columns
df1.rename(columns={"ID":"ID", "Course":"COURSE", "Desc":"COURSE DESC", "Teacher":"TEACHER", "Section":"SEC",
        "AHS ENTRY DATE":"ENTRY DATE", "STUDENT":"STUDENT", "GR":"GR", "T1":"T1"}, inplace=True)

df2.rename(columns={"ID":"ID", "COURSE":"COURSE", "COURSE DESC":"COURSE DESC", "TEACHER":"TEACHER", "SEC":"SEC",
        "AHS ENTRY DATE":"ENTRY DATE", "STUDENT":"STUDENT", "GR":"GR", "PR3":"PR3"}, inplace=True)


#combines spreadsheets and sort by student and course
f3 = pd.concat([df1, df2], sort=False)
f3 = f3.sort_values(["STUDENT", "COURSE"], ascending=[1,1])


#shape[0] is rows    shape[1] is columns
print(f"**** Row = {f3.shape[0]} Col = {f3.shape[1]} ****")

total_rows = int(f3.shape[0])
row_iterator = f3.iterrows()

#unpack tuple
_, last = next(row_iterator)
f4 = pd.DataFrame()

#loop iterates through all the rows to combine PR3 and T1 grades if they match ID and COURSE
skip = 0
for i, row in row_iterator:
     #if this row's ID and COURSE match the next, combine PR3 grade and append to new DataFrame
     if (row['ID'] == last['ID']) & (row['COURSE'] == last['COURSE']):
          last['PR3']=row['PR3']
          f4=f4.append([{'ENTRY DATE': last['ENTRY DATE'],'ID': last['ID'],'STUDENT': last['STUDENT'],'GR': last['GR'],'COURSE': last['COURSE'],\
                       'COURSE DESC': last['COURSE DESC'],'TEACHER': last['TEACHER'],'T1': last['T1'],'PR3': last['PR3']}]) #,ignore_index=True)
          skip = 1

     #skip if skip = 1; do not append
     elif skip == 1:
         skip = 0
         last = row

     #if unique, append to new DataFrame
     elif skip == 0:
          f4=f4.append([{'ENTRY DATE': last['ENTRY DATE'],'ID': last['ID'],'STUDENT': last['STUDENT'],'GR': last['GR'],'COURSE': last['COURSE'],\
                       'COURSE DESC': last['COURSE DESC'],'TEACHER': last['TEACHER'],'T1': last['T1'],'PR3': last['PR3']}]) #,ignore_index=True)


     last = row
     print("Progress ",i, "out of", total_rows)

#create new Excel spreadsheet with combined PR3 and T1 grades
with pd.ExcelWriter('PR3 And T1 Failure Report.xlsx') as writer:
      f4.to_excel(writer, sheet_name = "test1")
print("Complete")

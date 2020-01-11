# Excel-Failing-Grades
In one spreadsheet, student's failing progress report (PR3) grades. In another, their failing semester term (T1) grades. 
The program creates a new spreadsheet that places the PR and T1 grades next to each other to help administration highlight 
students failing their classes.

The program uses numpy and panda to read two existing excel files in the formats I have posted. 
PR3 Failure Report - contains students failing their Progress report
T1 Failure Report - contains students failing the entire semester

It then iterates through all the rows (originally 2,500). If the student is failing a course in the PR3 and T1, 
it is put together in a new file. 

The purpose of the program was to help administration easily see which students failed both PR3 and T1. Instead 
of searching both excel spreadsheet, the program places both grades side by side for quicker reading.

# Attendance
A small code that is used to automate the attendance Excel sheet specific for the Teachers Assistance role.

This code was created for "Environmental System Principles" -- An architectural class

This code utilizes the "OpenPyxl" library for Python.
Information in the .CSV file is pulled and then used to fill out the "ESP 2023 ATTENDANCE.xlsx" file.

The CSV file contains a question that was asked in the class via "polleverywhere.com" where the students can respond to the question pertaining to the class.
The automation keeps track of the analytical data for the students who answered the question correctly/incorrectly, which in turn keeps track of the students attendance. If the question was left un-answered by the student, that student is then marked absent via the color coding in the Excel sheet (purple).
Otherwise, if the student did answer the question, they are given a score of 1 (correctly answered) or 0 (in-correctly answered) which allows for an easy visual representation, as well as an easy way to calculate responses per student for analytical purposes.

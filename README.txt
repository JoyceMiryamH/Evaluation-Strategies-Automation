This program uses two files to output a strategies file. Those two files must conform to a certain format for the program to work properly:

* The relevant data of both files must be on the first (or only) sheet of that Excel file.
* The data source file must have a Name column, with the name of the facility where the sample was taken. Facility names must be consistent.
* The data source file must have a Date column, with the date of each sample written in text following this format: YYYY-MM-DD
* All the indicators must be on the right of the Date column.

* The indicator file must have "Indicators" (without quotes) in cell A1.
* The indicator file must have Name, Target, Threshold and Worst columns, with their column names respectively being in cells A2, B2, C2, D2.
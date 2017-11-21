# Evaluation-Strategies-Automation

This program uses two files to output a strategies file. Those two files must conform to a certain format for the program to work properly:

* The relevant data of both files must be on the first (or only) sheet of that Excel file.
* The data source file must have a Facility Name column, with the name of the facility where the sample was taken. Facility names must be consistent.
* The data source file must have a EEM Water Qual Mon Date column, with the date of each sample written in text following this format: YYYY-MM-DD
* All the indicators must be on the right of the EEM Water Qual Mon Date column.

* The indicator file must have Name, Target, Threshold and Worst columns, with these column names being written in the second row.
* The indicator file must have "Indicators" (without quotes) written in the cell above the Name column, in the first row.
* The indicators' names in the indicator file must be consistent with the indicators' names in the source file.
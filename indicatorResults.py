# PREREQUISITES in the README.txt document

import csv
from openpyxl import load_workbook
import pandas as pd
import sys
sys.path.extend(('C:\\Python34\\lib\\site-packages\\win32', 'C:\\Python34\\lib\\site-packages\\win32\\lib', 'C:\\Python34\\lib\\site-packages\\Pythonwin'))
import datetime as dt
import calendar
from dateutil.relativedelta import relativedelta
import numpy as np
import math

class INDICATORRESULTS():
	# Method for enumerating all the time periods of selected length between the start of the start year and the end of the end year
	# NOTE: While pratical for time spans larger than that, it is very inefficient for a daily timespan.
	def get_delimitation_dates(self, startYear, endYear, timespan):
		current = dt.date(startYear, 1, 1)

		if (timespan == '10years'):
			date_increment = relativedelta(years=10)
		elif (timespan == '5years'):
			date_increment = relativedelta(years=5)
		elif (timespan == '3years'):
			date_increment = relativedelta(years=3)
		elif (timespan == 'year'):
			date_increment = relativedelta(years=1)
		elif (timespan == 'bi-annual'):
			date_increment = relativedelta(months=6)
		elif (timespan == 'quarter'):
			date_increment = relativedelta(months=3)
		elif (timespan == 'month'):
			date_increment = relativedelta(months=1)
		elif (timespan == 'day'):
			date_increment = relativedelta(days=1)
		else:
			print("Not implemented.")

		dates_array = [[current.isoformat()]]
		i = 0
		while current.year <= endYear:
			current = current + date_increment
			dates_array[i].append((current - relativedelta(days=1)).isoformat())
			dates_array.append([current.isoformat()])
			i = i + 1

		del dates_array[-1]

		return dates_array

	# Method for getting a list of all the indicators and their line in the indicator file
	def get_attributes_list(self, sourcefile, indicatorsheet):
		attributes_src = list(sourcefile)[list(sourcefile).index('Date')+1:]
		attributes_ind = []
		for i in indicatorsheet:
			attributes_ind.append([i[0].value,i[0].row])
		# table with two values for each input: the attribute & the line number in the excel indicator file (e.g. for data in B22, 22 would be extracted)
		# the following allows to extract the list's headers
		del attributes_ind[0]
		del attributes_ind[0]

		attributes = list(set(attributes_src) & set([i[0] for i in attributes_ind]))
		for i in range(len(attributes_ind),0,-1):
			if (attributes_ind[i-1][0] not in attributes):
				del attributes_ind[i-1]
		# test prints :
		#print("Attributes list (source side):", attributes_src)
		#print("\nAttributes list (indicator side):", [i[0] for i in attributes_ind])
		#print("\nAttributes list (in common between the two):", attributes)
		#print("\nAttributes list (of only the attributes that will be taken in account b/c they're in both files):", attributes_ind)
		return attributes_ind

	# Method for getting the same list, but ordered in the same way as the source file.
	def get_best_list(self, sourcefile, attributesMatchedList):
		attributes_src = list(sourcefile)[list(sourcefile).index('Date')+1:]
		newattributes = [['Name', 'no']]
		for a in attributes_src:
			found = 0
			for i in attributesMatchedList:
				if a == i[0]:
					newattributes.append([a, i[1]])
					found = 1
					break
			if not found:
				newattributes.append([a, 'no'])
		return newattributes

	# Method for helping to create specific strategy names for each kind of periodicity
	def name_that_period(self, date_full, facility, timespan):
		date = date_full.split('-')
		if (timespan == 'year'):
			periodname = date[0]
		elif (timespan == '3years'):
			periodname = date[0] + '-' + str(int(date[0])+2)
		elif (timespan == '5years'):
			periodname = date[0] + '-' + str(int(date[0])+4)
		elif (timespan == '10years'):
			periodname = date[0] + '-' + str(int(date[0])+9)
		elif (timespan == 'bi-annual'):
			if (int(date[1]) < 6):
				periodname = 'S1 ' + date[0]
			else:
				periodname = 'S2 ' + date[0]
		elif (timespan == 'quarter'):
			if (int(date[1]) < 3):
				periodname = 'Q1 ' + date[0]
			elif (int(date[1]) < 6):
				periodname = 'Q2 ' + date[0]
			elif (int(date[1]) < 9):
				periodname = 'Q3 ' + date[0]
			else:
				periodname = 'Q4 ' + date[0]
		elif (timespan == 'month'):
			periodname = calendar.month_name[int(date[1])] + ' ' + date[0]
		elif (timespan == 'day'):
			periodname = date_full
		return facility + ' ' + str(periodname)
	
	# Method used to make sure the dates are in the right format (specifically, adding a month and day in case the date given is just a year)
	def correct_dates(self, date):
		if len(str(date)) == 4:
			return str(date) + "-01-01"
		else:
			return str(date)

	# Method used to make sure data that can't be parsed into a number gets sorted out as nan, making it simpler to treat it further down the line  
	def make_it_float(self, x):
		if not (x is None):
			try:
				return float(x)
			except ValueError:
				return np.nan

	# Method used to calculate the value of each indicator (before processing it as a strategy evaluation value) for each entity for each time period
	def main_loop(self, sourcefile, facility, dates, attributes, timespan):
		dfs_row = [self.name_that_period(dates[0], facility, timespan)]
		for a in attributes[1:]:
			if a[1] != 'no':
				pf1 = sourcefile.loc[(sourcefile['Name']==facility) & (sourcefile['Date']>=dates[0]) & (sourcefile['Date']<=dates[1])]
				pf2 = pf1[[a[0]]].dropna(axis=0, how='all')
				if pf2.empty:
					dfs_row.append('empty')
				else:
					dfs_row.append(pf2.iloc[:,0].mean())
			else:
				dfs_row.append('empty')

		# uncomment below line to see list created in main_loop
		#print(dfs_row)

		return dfs_row

	# Method used to calculate the strategy evaluation value for a given value of a given indicator
	def quantitative(self, target, threshold, worst, current):
		if not (isinstance(target, (int, float))):
			target = 0
		if not (isinstance(threshold, (int, float))):
			threshold = 0
		if not (isinstance(worst, (int, float))):
			worst = 0

		if (target<worst):
			if(current<=target):
				quanSatisfaction = 100
			elif(current>=worst):
				quanSatisfaction = -100
			elif(current<threshold):
				quanSatisfaction = (abs(current-threshold)/abs(target-threshold))*(100)
			else:
				quanSatisfaction = (abs(current-threshold)/abs(threshold-worst))*(-100)
		elif(current>=target):
			quanSatisfaction = 100
		elif(current<=worst):
			quanSatisfaction = -100
		elif(current>=threshold):
			quanSatisfaction = (abs(current-threshold)/abs(target-threshold))*(100)
		else:
			quanSatisfaction = (abs(current-threshold)/abs(threshold-worst))*(-100)
		return self.qualitative(quanSatisfaction)

	def qualitative(self, quanSatisfaction):
		return (quanSatisfaction/2)+50

	def main(self, source, indicator, results, startYear, endYear, timespan):
		print('starting automation\n')
		xls_file = pd.ExcelFile(source)
		df = xls_file.parse()
		wb2 = load_workbook(filename = indicator)
		ws2 = wb2.active
		print('files read without issues')

		df['Date'] = df['Date'].apply(self.correct_dates)

		# facilities : list of the different facilities found in the Names column in the source file
		facilities = pd.unique(df['Name']).tolist()
		facilities = [x for x in facilities if str(x) != 'nan']

		# attributes: the list of attributs FOR WHICH THE NAME IS IDENTICAL IN BOTH FILES (source and indicator)
		# 			  with the following two dimensions: name of attribute and its row position in the indicator file
		attributes = self.get_attributes_list(df, ws2)
		attributes_ordered = self.get_best_list(df, attributes)

		for a in attributes_ordered[1:]:
			df[a[0]] = df[a[0]].apply(self.make_it_float)

		# dates : list of all date values found between the defined timespan set by the start and end year values (inclusively)
		dates = self.get_delimitation_dates(startYear, endYear, timespan)
		print('data obtained without issues')


		# data_for_strategies: list which contains all the data for the strategies calculation
		#						each sub-table contains the name of the strategy (ex.: 'MM1030 2009') and all the values of the considered attributes
		data_for_strategies = []
		for i in facilities:
			for j in dates:
				row = self.main_loop(df, i, j, attributes_ordered, timespan)
				if not all(r == 'empty' for r in row[1:]):
					data_for_strategies.append(row)

		strategies_desc = []
		strategies_data = []

		for data in data_for_strategies:
			values = [data[0]]
			for d in range(1, len(data)):
				if (attributes_ordered[d][1] != 'no'):
					if (data[d] == 'empty'):
						values.append('')
					else:
						target = ws2['B'+str(attributes_ordered[d][1])].value
						threshold = ws2['C'+str(attributes_ordered[d][1])].value
						worst = ws2['D'+str(attributes_ordered[d][1])].value
						values.append(int(self.quantitative(target, threshold, worst, data[d])))
						#print(attributes_ordered[d][0], ":", data[d], "@ column:", attributes_ordered[d][1], ", result = ", values[-1])

			strategies_desc.append([data[0], '"esp"', '"No description"', '""'])
			strategies_data.append(values)

		with open(results, 'w', newline='') as csvfile:
			filewriter = csv.writer(csvfile, delimiter=',',
									quotechar='|', lineterminator='\n', quoting=csv.QUOTE_MINIMAL)

			#stratfilename = '"' + results[:-4] + '"'
			#filewriter.writerow(['GRL Strategies for', stratfilename])
			#filewriter.writerow([''])
			#filewriter.writerow([''])

			#filewriter.writerow(['Strategy Name', ' Author', ' Description', ' "Included Strategies"'])
			#for i in strategies_desc:
			#	filewriter.writerow(i)

			#filewriter.writerow([''])
			#filewriter.writerow([''])

			colNames = ['Strategy Name']
			for d in range(1, len(attributes_ordered)):
				if (attributes_ordered[d][1] != 'no'):
					colNames.append(attributes_ordered[d][0])
			filewriter.writerow(colNames)

			for i in strategies_data:
				filewriter.writerow(i)

		print('\nresult file created\n')
		sys.exit()

	if __name__ == "__main__": main()

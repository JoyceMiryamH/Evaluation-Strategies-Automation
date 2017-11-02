# TO BE IMPLEMENTED
	# boucle principale
	# indépendance des variances de fichiers source / indicateur (noms de fichiers, nom de colonnes)
	
# PREREQUISITES (à noter quelque part / intégrer dans l'user interface / etc. ou demander confirmation)
	# about the DATA SOURCE
		# the headers must be in the first row
		# all attributes must be in columns AFTER the date column, all other data BEFORE the date column
		# the name of the date column MUST be 'EEM Water Qual Mon Date'
	# about the INDICATOR TEMPLATE
		# the names of the attributes in this list must match the names of the attributes in the source file

import csv
from openpyxl import load_workbook
import pandas as pd
import sys
sys.path.extend(('C:\\Python34\\lib\\site-packages\\win32', 'C:\\Python34\\lib\\site-packages\\win32\\lib', 'C:\\Python34\\lib\\site-packages\\Pythonwin'))
import xlwings as xw
import datetime as dt
from dateutil.relativedelta import relativedelta

class INDICATORRESULTS():
	# doit pouvoir générer une liste des couples de dates de l'année de départ à l'année de fin en fonction de la périodicité 
	def get_delimitation_dates(self, startYear, endYear, timespan):
		current = dt.date(startYear, 1, 1)
		
		if (timespan == 'year'):
			date_increment = relativedelta(years=1)
		elif (timespan == 'semester'):
			date_increment = relativedelta(months=6)
		elif (timespan == 'trimester'):
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
	
	def main_loop(self, source, indicator, results, dates):
		return 0
		# boucle extérieure : par facility
		# boucle intérieure : par période
		
	
	def main(self, source, indicator, results, startYear, endYear, timespan):
		print('start automation\n')
		print('reading data source\n')
		xls_file = pd.ExcelFile(source)
		df = xls_file.parse('Sheet1')
		print('reading indicator template\n')
		wb2 = load_workbook(filename = indicator)
		ws2 = wb2.active

		facility = pd.unique(df['Facility Name'])
		
		attributes_src = list(df)[list(df).index('EEM Water Qual Mon Date')+1:]
		attributes_ind = []
		for i in ws2:
			attributes_ind.append([i[0].value,i[0].coordinate,0])
		attributes_ind.remove(['Indicators','A1', 0])
		attributes_ind.remove(['Name','A2', 0])
		print("Attributes list (source side):", attributes_src)
		print("\nAttributes list (indicator side):", [i[0] for i in attributes_ind])
		
		attributes = list(set(attributes_src) & set (attributes_ind))
		print("\nAttributes list (in common between the two):", attributes)
		
		
		dates = self.get_delimitation_dates(startYear, endYear, timespan)
		
		
		# set value for 2009
		pf1 = df.loc[(df['Facility Name']==facility[0]) & (df['EEM Water Qual Mon Date']>dates[0][0]) & (df['EEM Water Qual Mon Date']<dates[0][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val1 = sht['H22'].value
		wb.close()

		# set value for 2010
		pf1 = df.loc[(df['Facility Name']==facility[0]) & (df['EEM Water Qual Mon Date']>dates[1][0]) & (df['EEM Water Qual Mon Date']<dates[1][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val2 = sht['H22'].value
		wb.close()

		# set value for 2011
		pf1 = df.loc[(df['Facility Name']==facility[0]) & (df['EEM Water Qual Mon Date']>dates[2][0]) & (df['EEM Water Qual Mon Date']<dates[2][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val3 = sht['H22'].value
		wb.close()

		# set value for 2012
		pf1 = df.loc[(df['Facility Name']==facility[0]) & (df['EEM Water Qual Mon Date']>dates[3][0]) & (df['EEM Water Qual Mon Date']<dates[3][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val4 = sht['H22'].value
		wb.close()

		# set value for 2013
		pf1 = df.loc[(df['Facility Name']==facility[0]) & (df['EEM Water Qual Mon Date']>dates[4][0]) & (df['EEM Water Qual Mon Date']<dates[4][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val5 = sht['H22'].value
		wb.close()

		# set value for 2014
		pf1 = df.loc[(df['Facility Name']==facility[0]) & (df['EEM Water Qual Mon Date']>dates[5][0]) & (df['EEM Water Qual Mon Date']<dates[5][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val6 = sht['H22'].value
		wb.close()
		

		with open(results, 'w', newline='') as csvfile:
			filewriter = csv.writer(csvfile, delimiter=',',
									quotechar='|', lineterminator='\n', quoting=csv.QUOTE_MINIMAL) 
			filewriter.writerow(['Strategy Name', 'pH'])
			filewriter.writerow(['MM1030 2009', val1])
			filewriter.writerow(['MM1030 2010', val2])
			filewriter.writerow(['MM1030 2011', val3])
			filewriter.writerow(['MM1030 2012', val4])
			filewriter.writerow(['MM1030 2013', val5])
			filewriter.writerow(['MM1030 2014', val6])

		print('\nresult file created\n')
		sys.exit()


	if __name__ == "__main__": main()
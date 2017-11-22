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
import datetime as dt
import calendar
from dateutil.relativedelta import relativedelta
import numpy as np

class INDICATORRESULTS():
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

	def get_attributes_list(self, sourcefile, indicatorsheet):
		attributes_src = list(sourcefile)[list(sourcefile).index('EEM Water Qual Mon Date')+1:]
		attributes_ind = []
		for i in indicatorsheet:
			attributes_ind.append([i[0].value,i[0].row])
        # tableau avec deux valeurs pour chaque entrée: le nom d'attribut et le numéro de ligne dans le fichier indicateur
		del attributes_ind[0]
		del attributes_ind[0]
		# ^ permet de retirer les en-têtes de la liste


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
		
	def get_best_list(self, sourcefile, attributesMatchedList):
		attributes_src = list(sourcefile)[list(sourcefile).index('EEM Water Qual Mon Date')+1:]
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
			
	
	def name_that_period(self, date_full, facility, timespan):
		date = date_full.split('-')
		if (timespan == 'year'):
			periodname = date[0]
		elif (timespan == 'semester'):
			if (int(date[1]) < 6):
				periodname = 'S1 ' + date[0]
			else:
				periodname = 'S2 ' + date[0]
		elif (timespan == 'trimester'):
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
		#print(facility)
		#print(periodname)
		return facility + ' ' + str(periodname)

	def main_loop(self, sourcefile, facility, dates, attributes, timespan):
		dfs_row = [self.name_that_period(dates[0], facility, timespan)]
		for a in attributes:
			pf1 = sourcefile.loc[(sourcefile['Facility Name']==facility) & (sourcefile['EEM Water Qual Mon Date']>=dates[0]) & (sourcefile['EEM Water Qual Mon Date']<=dates[1])]
			pf2 = pf1[[a[0]]].dropna(axis=0, how='all')
			if pf2.empty:
				dfs_row.append('empty')
			else:
				dfs_row.append(pd.to_numeric(pf2.iloc[0]).mean())

        # décommenter la ligne ci-dessous pour voir la liste que produit main_loop
		#print(dfs_row)

		return dfs_row

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
		df = xls_file.parse('Sheet1')
		wb2 = load_workbook(filename = indicator)
		ws2 = wb2.active
		print('files read without issues')

		# facilities : la liste des différentes facilités comprises dans Facility Names dans le document source
		facilities = pd.unique(df['Facility Name']).tolist()
		del facilities[-1] 				#parce que unique donne un array, et après conversion en liste il reste l'élément NaN à la fin de la liste, donc snip, on coupe ça

		# attributes : la liste des attributs DONT LE NOM EST IDENTIQUE DANS LES DEUX FICHIERS SEULEMENT, avec trois dimensions : le nom de l'attribut,
		# 			   et sa rangée dans le fichier indicateur
		attributes = self.get_attributes_list(df, ws2)
		attributes_ordered = self.get_best_list(df, attributes)
		#print([i[0] for i in attributes_ordered])

		# dates : la liste des dates de début et de fin de chaque période comprise entre les deux années inclusivement (les dates de début et dates de fin sont les deux dimensions)
		dates = self.get_delimitation_dates(startYear, endYear, timespan)
		print('data obtained without issues')


		# data_for_strategies : liste qui va comprendre toutes les données pour le calcul de stratégies
		# 						chaque sous-tableau comprend le nom de la stratégie (ex.: 'MM1030 2009') et toutes les valeurs des attributs pris en compte
		data_for_strategies = []
		for i in facilities:
			for j in dates:
				row = self.main_loop(df, i, j, attributes, timespan)
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
						values.append(self.quantitative(target, threshold, worst, data[d]))
			
			strategies_desc.append([data[0], "esp", "No description"])
			strategies_data.append(values)

		with open(results, 'w', newline='') as csvfile:
			filewriter = csv.writer(csvfile, delimiter=',',
									quotechar='|', lineterminator='\n', quoting=csv.QUOTE_MINIMAL)
			
			filewriter.writerow(['GRL Strategies for', results[:-4]])
			filewriter.writerow([''])
			filewriter.writerow([''])
			
			filewriter.writerow(['Strategy Name', ' Author', ' Description', ' "Included Strategies"'])
			for i in strategies_desc:
				filewriter.writerow(i)
			filewriter.writerow([''])
			filewriter.writerow([''])
			
			colNames = ['Strategy Name']
			colNames.extend([i[0] for i in attributes])
			filewriter.writerow(colNames)
			for i in strategies_data:
				filewriter.writerow(i)

		print('\nresult file created\n')
		sys.exit()


	if __name__ == "__main__": main()

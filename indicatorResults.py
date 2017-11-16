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
import calendar
from dateutil.relativedelta import relativedelta

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
	
	def name_that_period(self, date_full, facility, timespan):
		date = date_full.split('-')
		if (timespan == 'year'):
			periodname = date[0]
		elif (timespan == 'semester'):
			if (date[1] < 6):
				periodname = 'S1 ' + date[0]
			else:
				periodname = 'S2 ' + date[0]
		elif (timespan == 'trimester'):
			if (date[1] < 3):
				periodname = 'Q1 ' + date[0]
			elif (date[1] < 6):
				periodname = 'Q2 ' + date[0]
			elif (date[1] < 9):
				periodname = 'Q3 ' + date[0]
			else:
				periodname = 'Q4 ' + date[0]
		elif (timespan == 'month'):
			periodname = calendar.month_name[date[1]] + ' ' + date[0]
		elif (timespan == 'day'):
			periodname = date_full
		#print(facility)
		#print(periodname)
		return facility + ' ' + str(periodname)
	
	def main_loop(self, sourcefile, facility, dates, attributes, timespan):
		dfs_row = [self.name_that_period(dates[0], facility, timespan)]
		for a in attributes:
			pf1 = sourcefile.loc[(sourcefile['Facility Name']==facility) & (sourcefile['EEM Water Qual Mon Date']>=dates[0]) & (sourcefile['EEM Water Qual Mon Date']<=dates[1])]
			pf2 = pf1[[a[0]]].dropna(axis=1, how='all')
			if pf2.empty:
				dfs_row.append('empty')
			else:
				dfs_row.append(pd.to_numeric(pf2.iloc[0]).mean())
                
        # décommenter la ligne ci-dessous pour voir la liste que produit main_loop 
		#print(dfs_row)
        
		return dfs_row
	
	def main(self, source, indicator, results, startYear, endYear, timespan):
		print('start automation\n')
		xls_file = pd.ExcelFile(source)
		df = xls_file.parse()
		wb2 = load_workbook(filename = indicator)
		ws2 = wb2.active
		print('files read without issues')
		
		# facilities : la liste des différentes facilités comprises dans Facility Names dans le document source
		facilities = pd.unique(df['Facility Name']).tolist()
		del facilities[-1] 				#parce que unique donne un array, et après conversion en liste il reste l'élément NaN à la fin de la liste, donc snip, on coupe ça
		
		# attributes : la liste des attributs DONT LE NOM EST IDENTIQUE DANS LES DEUX FICHIERS SEULEMENT, avec trois dimensions : le nom de l'attribut,
		# 			   et sa rangée dans le fichier indicateur
		attributes = self.get_attributes_list(df, ws2)
		
		# dates : la liste des dates de début et de fin de chaque période comprise entre les deux années inclusivement (les dates de début et dates de fin sont les deux dimensions)
		dates = self.get_delimitation_dates(startYear, endYear, timespan)
		print('data obtained without issues')
		
		
		# data_for_strategies : liste qui va comprendre toutes les données pour le calcul de stratégies
		# 						chaque sous-tableau comprend le nom de la stratégie (ex.: 'MM1030 2009') et toutes les valeurs des attributs pris en compte
		data_for_strategies = []
		for i in facilities:
			for j in dates:
				data_for_strategies.append(self.main_loop(df, i, j, attributes, timespan))
		
		# set value for 2009
		pf1 = df.loc[(df['Facility Name']==facilities[0]) & (df['EEM Water Qual Mon Date']>dates[0][0]) & (df['EEM Water Qual Mon Date']<dates[0][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val1 = sht['H22'].value
		wb.close()

		# set value for 2010
		pf1 = df.loc[(df['Facility Name']==facilities[0]) & (df['EEM Water Qual Mon Date']>dates[1][0]) & (df['EEM Water Qual Mon Date']<dates[1][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val2 = sht['H22'].value
		wb.close()

		# set value for 2011
		pf1 = df.loc[(df['Facility Name']==facilities[0]) & (df['EEM Water Qual Mon Date']>dates[2][0]) & (df['EEM Water Qual Mon Date']<dates[2][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val3 = sht['H22'].value
		wb.close()

		# set value for 2012
		pf1 = df.loc[(df['Facility Name']==facilities[0]) & (df['EEM Water Qual Mon Date']>dates[3][0]) & (df['EEM Water Qual Mon Date']<dates[3][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val4 = sht['H22'].value
		wb.close()

		# set value for 2013
		pf1 = df.loc[(df['Facility Name']==facilities[0]) & (df['EEM Water Qual Mon Date']>dates[4][0]) & (df['EEM Water Qual Mon Date']<dates[4][1])]
		pf2 = pf1[['Facility Name', 'EEM Water Qual Mon Date', 'pH']]
		meanVal = pf2.loc[:,'pH'].mean()
		ws2['F22'] = meanVal
		wb2.save(indicator)

		wb = xw.Book(indicator)
		sht = wb.sheets['Indicators']
		val5 = sht['H22'].value
		wb.close()

		# set value for 2014
		pf1 = df.loc[(df['Facility Name']==facilities[0]) & (df['EEM Water Qual Mon Date']>dates[5][0]) & (df['EEM Water Qual Mon Date']<dates[5][1])]
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
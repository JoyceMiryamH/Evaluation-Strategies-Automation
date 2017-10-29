# TO BE IMPLEMENTED
    # sélection de la période
    # sélection de la liste des attributs
    # boucle principale
    # indépendance des variances de fichiers source / indicateur (noms de fichiers, nom de colonnes)

import csv
from openpyxl import load_workbook
import pandas as pd
import sys
sys.path.extend(('C:\\Python34\\lib\\site-packages\\win32', 'C:\\Python34\\lib\\site-packages\\win32\\lib', 'C:\\Python34\\lib\\site-packages\\Pythonwin'))
import xlwings as xw

class INDICATORRESULTS():
    # doit pouvoir générer une liste des couples de dates de l'année de départ à l'année de fin en fonction de la périodicité 
    def get_delimitation_dates(self, startYear, endYear, timespan):
        return [['2009-01-01', '2009-12-31'], ['2010-01-01', '2010-12-31'], ['2011-01-01', '2011-12-31'], ['2012-01-01', '2012-12-31'], ['2013-01-01', '2013-12-31'], ['2014-01-01', '2014-12-31']]
    
    def main_loop(self, source, indicator, results, dates):
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

        # gets all facility names
        facility = pd.unique(df['Facility Name'])
        
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
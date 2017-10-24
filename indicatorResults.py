import csv
from openpyxl import load_workbook
import pandas as pd
import sys
sys.path.extend(('C:\\Python34\\lib\\site-packages\\win32', 'C:\\Python34\\lib\\site-packages\\win32\\lib', 'C:\\Python34\\lib\\site-packages\\Pythonwin'))
import xlwings as xw

class INDICATORRESULTS():
    def main():
        print('start automation\n')
        print('reading data source\n')
        wb1 = load_workbook(filename = '..\data-source.xlsx')
        ws1 = wb1.active
        print('reading indicator template\n')
        wb2 = load_workbook(filename = '..\indicator-template.xlsx')
        sh2 = wb2['Indicators']
        ws2 = wb2.active

        sheet_ranges = wb2['Indicators']

        xls_file = pd.ExcelFile('..\data-source.xlsx')
        df = xls_file.parse('Sheet1')

        facility = 'MM1030'
        
        # set value for 2009
        pf1 = df.loc[(df['Facility Name']==facility) & (df['Date']>'2009-01-01') & (df['Date']<'2009-12-31')]
        pf2 = pf1[['Facility Name', 'Date', 'pH']]
        meanVal = pf2.loc[:,'pH'].mean()
        ws2['F22'] = meanVal
        wb2.save('..\indicator-template.xlsx')

        wb = xw.Book('..\indicator-template.xlsx')
        sht = wb.sheets['Indicators']
        val1 = sht['H22'].value
        wb.close()

        # set value for 2010
        pf1 = df.loc[(df['Facility Name']==facility) & (df['Date']>'2010-01-01') & (df['Date']<'2010-12-31')]
        pf2 = pf1[['Facility Name', 'Date', 'pH']]
        meanVal = pf2.loc[:,'pH'].mean()
        ws2['F22'] = meanVal
        wb2.save('..\indicator-template.xlsx')

        wb = xw.Book('..\indicator-template.xlsx')
        sht = wb.sheets['Indicators']
        val2 = sht['H22'].value
        wb.close()

        # set value for 2011
        pf1 = df.loc[(df['Facility Name']==facility) & (df['Date']>'2011-01-01') & (df['Date']<'2011-12-31')]
        pf2 = pf1[['Facility Name', 'Date', 'pH']]
        meanVal = pf2.loc[:,'pH'].mean()
        ws2['F22'] = meanVal
        wb2.save('..\indicator-template.xlsx')

        wb = xw.Book('..\indicator-template.xlsx')
        sht = wb.sheets['Indicators']
        val3 = sht['H22'].value
        wb.close()

        # set value for 2012
        pf1 = df.loc[(df['Facility Name']==facility) & (df['Date']>'2012-01-01') & (df['Date']<'2012-12-31')]
        pf2 = pf1[['Facility Name', 'Date', 'pH']]
        meanVal = pf2.loc[:,'pH'].mean()
        ws2['F22'] = meanVal
        wb2.save('..\indicator-template.xlsx')

        wb = xw.Book('..\indicator-template.xlsx')
        sht = wb.sheets['Indicators']
        val4 = sht['H22'].value
        wb.close()

        # set value for 2013
        pf1 = df.loc[(df['Facility Name']==facility) & (df['Date']>'2013-01-01') & (df['Date']<'2013-12-31')]
        pf2 = pf1[['Facility Name', 'Date', 'pH']]
        meanVal = pf2.loc[:,'pH'].mean()
        ws2['F22'] = meanVal
        wb2.save('..\indicator-template.xlsx')

        wb = xw.Book('..\indicator-template.xlsx')
        sht = wb.sheets['Indicators']
        val5 = sht['H22'].value
        wb.close()

        # set value for 2014
        pf1 = df.loc[(df['Facility Name']==facility) & (df['Date']>'2014-01-01') & (df['Date']<'2014-12-31')]
        pf2 = pf1[['Facility Name', 'Date', 'pH']]
        meanVal = pf2.loc[:,'pH'].mean()
        ws2['F22'] = meanVal
        wb2.save('..\indicator-template.xlsx')

        wb = xw.Book('..\indicator-template.xlsx')
        sht = wb.sheets['Indicators']
        val6 = sht['H22'].value
        wb.close()
        

        with open('..\strategie-results.csv', 'w', newline='') as csvfile:
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


    if __name__ == "__main__": main()
 
        

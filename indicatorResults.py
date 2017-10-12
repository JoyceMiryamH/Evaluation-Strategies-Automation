import csv
from openpyxl import load_workbook

class INDICATORRESULTS():
    def main():
        print('start automation\n')
        print('reading data source\n')
        wb1 = load_workbook(filename = 'data-source.xlsx')
        print('reading indicator template\n')
        wb2 = load_workbook(filename = 'indicator-template.xlsx')

        sheet_ranges = wb2['Indicators']
        print(sheet_ranges['D22'].value)

        with open('strategie-results.csv', 'wb') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',',
                                    quotechar='|', quoting=csv.QUOTE_MINIMAL)
    

    if __name__ == "__main__": main()
 
        

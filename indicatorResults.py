import csv
from openpyxl import load_workbook

class INDICATORRESULTS():
    def main():
        print('start automation\n')
        print('reading data source\n')
        wb1 = load_workbook(filename = '..\data-source.xlsx')
        ws1 = wb1.active
        print('reading indicator template\n')
        wb2 = load_workbook(filename = '..\indicator-template.xlsx')
        ws2 = wb2.active

        sheet_ranges = wb2['Indicators']

        colValues = {}
        for row in ws2.iter_rows(min_row=3, max_col=1, max_row=25):
            for cell in row:
                colValues[cell] = cell.value

        for cols in ws1:
            for cell in cols:
                if cell.value == "Date":
                    print(cell.value)
            

        with open('..\strategie-results.csv', 'w', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',',
                                    quotechar='|', lineterminator='\n', quoting=csv.QUOTE_MINIMAL)


        
        
        print('result file created\n')


    if __name__ == "__main__": main()
 
        

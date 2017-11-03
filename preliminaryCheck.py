import csv
from openpyxl import load_workbook
import pandas as pd
import sys
sys.path.extend(('C:\\Python34\\lib\\site-packages\\win32', 'C:\\Python34\\lib\\site-packages\\win32\\lib', 'C:\\Python34\\lib\\site-packages\\Pythonwin'))
import xlwings as xw
import datetime as dt
from dateutil.relativedelta import relativedelta

class PreliminaryCheck():
	def check_data_source(self, source):
		try:
			xls_file = pd.ExcelFile(source)
		except Exception:
			return("Not an Excel file.")
		df = xls_file.parse()
		
		if 'Facility Name' not in df.columns:
			return("No Facility Name column found. (the data must be in the first sheet of the file)")
		elif 'EEM Water Qual Mon Date' not in df.columns:
			return("No EEM Water Qual Mon Date column found.")
		else:
			return("Valid file.")
		
	
	def main(self):
		sys.exit()


	if __name__ == "__main__": main()
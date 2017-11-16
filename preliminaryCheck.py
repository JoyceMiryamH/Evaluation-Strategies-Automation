import csv
from openpyxl import load_workbook
import pandas as pd
import sys
sys.path.extend(('C:\\Python34\\lib\\site-packages\\win32', 'C:\\Python34\\lib\\site-packages\\win32\\lib', 'C:\\Python34\\lib\\site-packages\\Pythonwin'))
import xlwings as xw
import datetime as dt
from dateutil.relativedelta import relativedelta
from indicatorResults import INDICATORRESULTS as ir

class PreliminaryCheck():
	def check_data_source(self, source):
		xls_file = pd.ExcelFile(source)
		
		df = xls_file.parse()
		
		if 'Facility Name' not in df.columns:
			return("ERROR: No Facility Name column found in Source file.\n\nPlease check the documentation to know how your file ought to be formatted.")
		elif 'EEM Water Qual Mon Date' not in df.columns:
			return("ERROR: No EEM Water Qual Mon Date column found in Source file.\n\nPlease check the documentation to know how your file ought to be formatted.")
		else:
			return("Valid file.")
		
	def check_indicator(self, indicator):
		xls_file = pd.ExcelFile(indicator)
		
		df = xls_file.parse(header=1)
		
		if 'Name' not in df.columns:
			return("ERROR: No Name column found in Indicator file.\n\nPlease check the documentation to know how your file ought to be formatted.")
		elif 'Target' not in df.columns:
			return("ERROR: No Target column found in Indicator file.\n\nPlease check the documentation to know how your file ought to be formatted.")
		elif 'Threshold' not in df.columns:
			return("ERROR: No Threshold column found in Indicator file.\n\nPlease check the documentation to know how your file ought to be formatted.")
		elif 'Worst' not in df.columns:
			return("ERROR: No Worst column found in Indicator file.\n\nPlease check the documentation to know how your file ought to be formatted.")
		else:
			return("Valid file.")
			
	def get_indicators(self, source, indicator):
		xls_file = pd.ExcelFile(source)
		df = xls_file.parse()
		wb2 = load_workbook(filename = indicator)
		ws2 = wb2.active
		indicators = ir().get_attributes_list(df, ws2)
		
		return "The indicators we found were: " + ", ".join([i[0] for i in indicators]) + ".\n\nIf you expected other indicators to be found, please make sure they have the exact same name in both files.\n\nIf not, please click OK."
		
		
	
	def main(self):
		sys.exit()


	if __name__ == "__main__": main()
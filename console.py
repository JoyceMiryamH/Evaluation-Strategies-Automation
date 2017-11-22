import sys
import os.path
import re
from preliminaryCheck import PreliminaryCheck as pc
from indicatorResults import INDICATORRESULTS as ir

# méthode pour simplifier la vérification des années en argument
def isInt(s):
	try:
		int(s)
		return True
	except ValueError:
		return False


def process(src, ind, tsp, syr, eyr, dst):
	if check(src, ind, tsp, syr, eyr, dst, "silent"):
		if (dst.split('.')[-1] != "csv" or len(dst.split('.')) == 1):
			dst = dst + ".csv"
		return ir().main(src, ind, dst, syr, eyr, tsp)
	else:
		print("Note: please don't manipulate the source and indicator files while running")
		print("          this program.")

# méthode pour vérifier la validité des arguments
def check(src, ind, tsp, syr, eyr, dst, mode):
	if not (os.path.isfile(src)):
		print("ERROR: Your first argument must be a valid file path (pointing to your")
		print("          data source excel file).")
	elif not (os.path.isfile(ind)):
		print("ERROR: Your second argument must be a valid file path (pointing to your")
		print("          indicator excel file).")
	elif (src == ind):
		print("ERROR: The source and indicator files must not be the same file.")
	elif (src.split('.')[-1] != "xlsx" or ind.split('.')[-1] != "xlsx"):
		print("ERROR: The source and indicator files must have the \".xlsx\" extension. (must ")
		print("          be in lowercase)")
	elif not (tsp in ['day', 'month', 'trimester', 'semester', 'year']):
		print("ERROR: The time span must be either 'day', 'month', 'trimester', 'semester' or")
		print("          'year' (without the quotation marks).")
	elif not (isInt(syr) and isInt(eyr)):
		print("ERROR: The \"From / to:\" fields must both represent years, written as integers")
		print("          (no decimal value, no characters other than numbers).")
	elif (syr > eyr):
		print("ERROR: The second \"From / to:\" field must represent the last year of your time")
		print("          span, while the first field represents the first year. The last year")
		print("          cannot be set before the first year.")
	elif not (re.match("^[A-Za-z0-9\_\-\.]+$", dst)) or dst == ".":
		print("ERROR: The strategies file name cannot contain any characters beside letters,")
		print("          numbers, underscores, dashes and points.")
	else:
		try:
			source_state = pc().check_data_source(src)
			if source_state == "Valid file.":
				try:
					indicator_state = pc().check_indicator(ind)
					if indicator_state == "Valid file.":
						if mode != "silent":
							print(pc().get_indicators(src, ind))
							if mode == "manual":
								if input("\nProceed? [Y/N] ") in ["y", "Y", "Yes", "YES", "yes"]:
									return 1
								else:
									print("Process aborted. Please do the appropriate modifications BEFORE calling this")
									print("          program again.")
									return 0
						return 1
					else:
						print(indicator_state)
				except Exception:
					print("ERROR: Indicator file is not an Excel file.")
			else:
				print(source_state)
		except Exception:
			print("ERROR: Source file is not an Excel file.")
			
	return 0


manual = 0
if (len(sys.argv) == 8):
	if sys.argv[7] == '--m':
		manual = 1	

if (len(sys.argv) != 7 and not manual) or (sys.argv[1] == 'help'):

	print("\nERROR: Wrong use of the command line arguments. Please make sure you write")
	print("          your command the following way, with all the following arguments.")
	print("   console.py <sourcepath> <indicatorpath> <timespan> <startyear> <endyear>")
	print("                 <destination> [--m]\n")
	
	print("Here are the arguments you have to use:")
	print("   sourcepath      The path to your data source file (e.g.: \"C:\\src.xlsx\" ).")
	print("   indicatorpath   The path to your indicator file (same as above).")
	print("   timespan        The time span you want to use to calculate your strategies.")
	print("                      Must be either 'day', 'month', 'trimester', 'semester' or")
	print("                      'year' (without the quotation marks).")
	print("   startyear       The first year for which you want to calculate strategies.")
	print("   endyear         The last year for which you want to calculate strategies")
	print("                      (inclusively).")
	print("   destination     The name you want to use for the resulting strategies file.\n")
	
	print("Here is an optional argument that you may add after the mandatory ones:")
	print("   --m             To enable manual validation of found indicators.")
else:
	src = sys.argv[1]
	ind = sys.argv[2]
	tsp = sys.argv[3]
	syr = sys.argv[4]
	eyr = sys.argv[5]
	dst = sys.argv[6]
	if manual:
		if check(src, ind, tsp, syr, eyr, dst, "manual"):
			process(src, ind, tsp, int(syr), int(eyr), dst)
	else:
		if check(src, ind, tsp, syr, eyr, dst, "normal"):
				process(src, ind, tsp, int(syr), int(eyr), dst)
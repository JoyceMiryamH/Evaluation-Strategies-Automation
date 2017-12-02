import sys
import os.path
import re
from preliminaryCheck import PreliminaryCheck as pc
from indicatorResults import INDICATORRESULTS as ir

# Method to simplify checking the values of the years given as arguments
def isInt(s):
	try:
		int(s)
		return True
	except ValueError:
		return False

# Method to process the checked files (does a last check just in case the user is modifying the file while processing it)
def process(src, ind, tsp, syr, eyr, dst):
	if check(src, ind, tsp, syr, eyr, dst, "silent"):
		if (dst.split('.')[-1] != "csv" or len(dst.split('.')) == 1):
			dst = dst + ".csv"
		return ir().main(src, ind, dst, syr, eyr, tsp)
	else:
		print("Note: please don't manipulate the source and indicator files while running")
		print("          this program.")

# Method to check the validity of the values given as arguments
def check(src, ind, tsp, syr, eyr, dst, mode):
	if not (os.path.isfile(src)):
		print("\nERROR: Your first argument must be a valid file path (pointing to your")
		print("          data source excel file).")
	elif not (os.path.isfile(ind)):
		print("\nERROR: Your second argument must be a valid file path (pointing to your")
		print("          indicator excel file).")
	elif (src == ind):
		print("\nERROR: The source and indicator files must not be the same file.")
	elif (src.split('.')[-1] != "xlsx" or ind.split('.')[-1] != "xlsx"):
		print("\nERROR: The source and indicator files must have the \".xlsx\" extension. (must ")
		print("          be in lowercase)")
	elif not (tsp in ['day', 'month', 'quarter', 'bi-annual', 'year', '3years', '5years', '10years']):
		print("\nERROR: The time span must be either 'day', 'month', 'quarter', 'bi-annual',")
		print("          'year', '3years', '5years' or '10years' (without the quotation")
		print("          marks).")
	elif not (isInt(syr) and isInt(eyr)):
		print("\nERROR: The \"From / to:\" fields must both represent years, written as integers")
		print("          (no decimal value, no characters other than numbers).")
	elif (syr > eyr):
		print("\nERROR: The second \"From / to:\" field must represent the last year of your time")
		print("          span, while the first field represents the first year. The last year")
		print("          cannot be set before the first year.")
	elif ((tsp == '3years') and ((int(eyr) - int(syr)) < 2)) or ((tsp == '5years') and ((int(eyr) - int(syr)) < 4)) or ((tsp == '10years') and ((int(eyr) - int(syr)) < 9)):
		print("\nERROR: The time span cannot be larger than the total time between the")
		print("          beginning of the first year and the end of the last year.")
	elif not (re.match("^[A-Za-z0-9\_\-\.]+$", dst)) or dst == ".":
		print("\nERROR: The strategies file name cannot contain any characters beside letters,")
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
						print("\n", indicator_state, sep='')
				except Exception:
					print("\nERROR: Indicator file is not an Excel file.")
			else:
				print(source_state)
		except Exception:
			print("\nERROR: Source file is not an Excel file.")

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
	print("                      Must be either 'day', 'month', 'quarter', 'bi-annual',")
	print("                      'year', '3years', '5years' or '10years' (without the")
	print("                      quotation marks).")
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

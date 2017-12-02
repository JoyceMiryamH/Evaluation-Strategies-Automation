# to create a working windows executable: http://www.pyinstaller.org/


from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import *
from tkinter import scrolledtext as tkst
from preliminaryCheck import PreliminaryCheck as pc
from indicatorResults import INDICATORRESULTS as ir
import re

class Window:
	status = 0

	# Creation of structure and objects of the UI
	def __init__(self, master):
		self.filenames = ["",""]

		# The first line, for the source file
		self.choice1=Label(root, text="Source file: " ).grid(row=1, column=1, sticky = E)
		self.bar1=Entry(master, state='disabled', disabledbackground="white", disabledforeground="black")
		self.bar1.grid(row=1, column=2, sticky = W + E, padx = 10, columnspan = 2)
		self.bbutton= Button(root, text="Browse", command= lambda: self.browsexlsx(0,self.bar1,self.filenames[0]))
		self.bbutton.grid(row=1, column=4, sticky = E)

		# The second line, for the indicator file
		self.choice2=Label(root, text="Indicator file: " ).grid(row=2, column=1, sticky = E)
		self.bar2=Entry(master, state='disabled', disabledbackground="white", disabledforeground="black")
		self.bar2.grid(row=2, column=2, sticky = W + E, padx = 10, columnspan = 2)
		self.bbutton= Button(root, text="Browse", command= lambda: self.browsexlsx(1,self.bar2,self.filenames[1]))
		self.bbutton.grid(row=2, column=4, sticky = E)

		# The third line, for the time period selection
		self.choice3=Label(root, text="Time span: " ).grid(row=3, column=1, sticky = E)
		self.value = StringVar(root)
		self.value.set('year')
		self.bar3=OptionMenu(root, self.value, 'day', 'month', 'quarter', 'bi-annual', 'year', '3years', '5years', '10years')
		self.bar3.grid(row=3, column=2, columnspan=2, padx = 10, sticky = W+E)

		# The fifth line (displayed in 4th place), for the choice of the select years (or other time value)
		self.choice5=Label(root, text="From / to: ").grid(row=4, column=1, sticky = E)
		self.bar5dot1=Entry(master)
		self.bar5dot2=Entry(master)
		self.bar5dot1.grid(row=4, column=2, sticky = W + E, padx = 10)
		self.bar5dot2.grid(row=4, column=3, sticky = W + E, padx = 10)
		self.choice5=Label(root, text="(years)", fg="grey").grid(row=4, column=4, sticky = W)

		# The fourth line (displayed in 5th place), for the choice of name of the result file
		self.choice4=Label(root, text="Strategies \n file name: ").grid(row=5, column=1, sticky = E)
		self.bar4=Entry(master)
		self.bar4.grid(row=5, column=2, sticky = W + E, padx = 10, columnspan = 3)

		# The antepenultimate line, to display error messages if needed
		self.status=tkst.ScrolledText(root, bg = "white", relief="groove", width=30, height=10, wrap=WORD, state=DISABLED)
		self.status.grid(row=8, column=1, columnspan=4, sticky=W+E)

		# The final buttons
		self.cbutton= Button(root, text="Check", command= lambda: self.preliminaryCheck("normal"))
		self.cbutton.grid(row=10, column=3, sticky = E)
		self.obutton= Button(root, text="OK", command=self.process_strategies, state=DISABLED)
		self.obutton.grid(row=10, column=4, sticky = W + E)

		# Spacing
		root.grid_columnconfigure(0, minsize=10)
		root.grid_columnconfigure(5, minsize=10)
		root.grid_rowconfigure(0, minsize=10)
		root.grid_rowconfigure(7, minsize=10)
		root.grid_rowconfigure(8, minsize=50)
		root.grid_rowconfigure(9, minsize=10)
		root.grid_rowconfigure(11, minsize=10)

		# Return touch binding, for efficiency's sake
		root.bind('<Return>', lambda e: self.preliminaryCheck("normal"))

	# Method for editing the text in the srollable text module
	def newText(self, text, color):
		self.status.config(state=NORMAL, fg=color)
		self.status.delete(1.0, END)
		self.status.insert(END, text)
		self.status.config(state=DISABLED)

	# Method for checking that the year values are valid
	def isInt(self, s):
		try:
			int(s)
			return True
		except ValueError:
			return False


	# Method for checking the arguments and files, and send back feedback on them
	# The 'mode' argument can be set to 'silent' if we do not wish to put anything in the window excluding errors, which is useful for a last check before rolling
	def preliminaryCheck(self, mode):
		syr = int(self.bar5dot1.get())
		eyr = int(self.bar5dot2.get())
		tsp = self.value.get()
		if (self.filenames[0] == "" or self.filenames[1] == "" or self.bar4.get() == ""or self.bar5dot1.get() == ""or self.bar5dot2.get() == ""):
			self.newText("ERROR: Please fill in required fields (i.e, all of them).", "red")
			status = 0
		elif (self.filenames[0] == self.filenames[1]):
			self.newText("ERROR: The source and indicator files must not be the same file.", "red")
			status = 0
		elif (self.filenames[0].split('.')[-1] != "xlsx" or self.filenames[1].split('.')[-1] != "xlsx"):
			self.newText("ERROR: The source and indicator files must have the \".xlsx\" extension. (must be in lowercase)", "red")
			status = 0
		elif not (self.isInt(self.bar5dot1.get()) and self.isInt(self.bar5dot2.get())):
			self.newText("ERROR: The \"From / to:\" fields must both represent years, written as integers (no decimal value, no characters other than numbers).", "red")
			status = 0
		elif (self.bar5dot1.get() > self.bar5dot2.get()):
			self.newText("ERROR: The second \"From / to:\" field must represent the last year of your time span, while the first field represents the first year. The last year cannot be set before the first year.", "red")
			status = 0
		elif ((tsp == '3years') and ((eyr - syr) < 2)) or ((tsp == '5years') and ((eyr - syr) < 4)) or ((tsp == '10years') and ((eyr - syr) < 9)):
			self.newText("ERROR: The time span cannot be larger than the total time between the beginning of the first year and the end of the last year.", "red")
			status = 0
		elif not (re.match("^[A-Za-z0-9\_\-\.]+$", self.bar4.get())) or self.bar4.get() == ".":
			self.newText("ERROR: The strategies file name cannot contain any characters beside letters, numbers, underscores, dashes and points.", "red")
			status = 0
		else:
			try:
				source_state = pc().check_data_source(self.filenames[0])
				if source_state == "Valid file.":
					try:
						indicator_state = pc().check_indicator(self.filenames[1])
						if indicator_state == "Valid file.":
							if mode != "silent":
								self.newText(pc().get_indicators(self.filenames[0], self.filenames[1]) + "\n\nIf you are fine with the currently selected indicators, please click OK. Else, please do the appropriate modifications and click Check.", "black")
							status = 1
						else:
							self.newText(indicator_state, "red")
							status = 0
					except Exception:
						self.newText("ERROR: Indicator file is not an Excel file.", "red")
						status = 0
				else:
					self.newText(source_state, "red")
					status = 0
			except Exception:
				self.newText("ERROR: Source file is not an Excel file.", "red")
				status = 0

		if status:
			self.obutton.config(state="normal")
			root.bind('<Return>', lambda e: self.process_strategies())
		else:
			self.obutton.config(state=DISABLED)
			root.bind('<Return>', lambda e: self.preliminaryCheck("normal"))
		return status

	# Method implementing the functionality of the 'Browse' button
	def browsexlsx(self, filename_id, bar, original_filename):
		Tk().withdraw()
		bar.config(state='normal')
		self.filenames[filename_id] = askopenfilename()
		if (self.filenames[filename_id] == ""):
			self.filenames[filename_id] = original_filename
		else:
			bar.delete(0, END)
			bar.insert(10, self.extract_filename(self.filenames[filename_id]))
		bar.config(state='disabled')

	# Method to properly display the names of the selected files results (code is seperate due to possible portability issues cause)
	# WARNING: will probably not work on Mac/Linux operating systems, better solutions may exist
	def extract_filename(self, filepath):
		pathlist = filepath.split('/')
		return pathlist[-1]

	# Method for actually running the processor (the core functionality)
	def process_strategies(self):
		if self.preliminaryCheck("silent"):
			newfilename = self.bar4.get()
			if (self.bar4.get().split('.')[-1] != "csv" or len(self.bar4.get().split('.')) == 1):
				newfilename = newfilename + ".csv"

			return ir().main(self.filenames[0], self.filenames[1], newfilename, int(self.bar5dot1.get()), int(self.bar5dot2.get()), self.value.get())

root = Tk()
root.title("Evalution Strategies Processor")
window=Window(root)
root.columnconfigure(2, weight=1)
root.update()
root.minsize(320, root.winfo_height())
root.resizable(width=False, height=False)
root.mainloop()

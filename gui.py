from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import *
import sys

class Window:
	# crée la structure et les objets de l'UI
	def __init__(self, master):
		self.filenames = ["",""]
	
		# la première ligne, pour le fichier source
		self.choice1=Label(root, text="Source file: " ).grid(row=1, column=1, sticky = E)
		self.bar1=Entry(master, state='disabled', disabledbackground="white", disabledforeground="black")
		self.bar1.grid(row=1, column=2, sticky = W + E, padx = 10)
		self.bbutton= Button(root, text="Browse", command= lambda: self.browsexlsx(0,self.bar1))
		self.bbutton.grid(row=1, column=3, sticky = E)
		
		# la seconde ligne, pour le fichier indicateur
		self.choice2=Label(root, text="Indicator file: " ).grid(row=2, column=1, sticky = E)
		self.bar2=Entry(master, state='disabled', disabledbackground="white", disabledforeground="black")
		self.bar2.grid(row=2, column=2, sticky = W + E, padx = 10)
		self.bbutton= Button(root, text="Browse", command= lambda: self.browsexlsx(1,self.bar2))
		self.bbutton.grid(row=2, column=3, sticky = E)
		
		# la troisième ligne, pour la sélection des périodes
		self.choice3=Label(root, text="Time span: " ).grid(row=3, column=1, sticky = E)
		self.bar3=Label(root, text="Yearly (more options to come)", fg = "grey" ).grid(row=3, column=2, columnspan=2)
		
		# la quatrième ligne, pour le choix du nom du fichier résultat
		self.choice4=Label(root, text="Strategies \n file name: ").grid(row=4, column=1, sticky = E)
		self.bar4=Entry(master)
		self.bar4.grid(row=4, column=2, sticky = W + E, padx = 10, columnspan = 2)
		
		# l'antépénultième vraie ligne, pour displayer le message d'erreur
		self.status=Label(root, text="", fg = "red")
		self.status.grid(row=9, column=1, columnspan=3)
		
		# LE BOUTON FINAL WOOHOO (soyons enthousiastes, à ce point là la fenêtre est finie, cool non ?)  
		self.cbutton= Button(root, text="OK", command=self.process_strategies)
		self.cbutton.grid(row=10, column=3, sticky = W + E)
		
		# de l'espace
		root.grid_columnconfigure(0, minsize=10)
		root.grid_columnconfigure(4, minsize=10)
		root.grid_rowconfigure(0, minsize=10)
		root.grid_rowconfigure(11, minsize=10)
		
		#et enfin une petite touche d'interactivité pour nos amis les gens qui en ont marre de leur trackpack
		root.bind('<Return>', lambda e: self.process_strategies())
		

	# fonction lancée quand on clique sur Browse
	def browsexlsx(self, filename_id, bar):
		Tk().withdraw()
		bar.config(state='normal')
		self.filenames[filename_id] = askopenfilename()
		if (self.filenames[filename_id] == ""):
			self.filenames[filename_id] = bar.get()
		else:
			bar.delete(0, END)
			bar.insert(10, self.extract_filename(self.filenames[filename_id]))
		bar.config(state='disabled')

	# fonction pour afficher proprement le résultat du Browse, mise à part à cause de possibles problèmes de portabilité
	# ATTENTION: ça fait joli mais c'est probablement pas utilisable sur des OS Unix, à voir si on peut faire mieux
	def extract_filename(self, filepath):
		pathlist = filepath.split('/')
		return pathlist[-1]
		
	# fonction lancée quand on clique sur OK (c'est là qu'on veut mettre notre code mais on peut aussi juste appeler ton code à partir de là)
	# PLACEHOLDER A COMPLETER, du coup
	def process_strategies(self):
		if (self.filenames[0] == "" or self.filenames[1] == "" or self.bar4.get() == ""):
			self.status.config(text="Please fill in required fields (i.e, all of them).")
			print()
		elif (self.filenames[0] == self.filenames[1]):
			self.status.config(text="The source and indicator files may not be the same file.")
			print()
		elif (self.filenames[0].split('.')[-1] != "xlsx" or self.filenames[1].split('.')[-1] != "xlsx"):
			self.status.config(text="The source and indicator files must have the XLSX extension.")
			print()
		else:
			print("Data source file path: ",self.filenames[0])
			print("Indicator file path: ",self.filenames[1])
			if (self.bar4.get().split('.')[-1] != "csv" or len(self.bar4.get().split('.')) == 1):
				print("Strategies file name:",self.bar4.get() + ".csv")
			else:
				print("Strategies file name:",self.bar4.get())
			sys.exit()

root = Tk()
root.title("Evalution Strategies Processor")
window=Window(root)
root.columnconfigure(2, weight=1)
root.update()
root.minsize(320, root.winfo_height())
root.resizable(width=False, height=False)
root.mainloop()  
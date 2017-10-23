from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import *
import csv
import sys

class Window:
	# crée la structure et les objets de l'UI ( me demande pas comment ça marche en détail c'est la magie d'Internet :x )
	def __init__(self, master):
		self.filenames = ["",""]
	
		# la première ligne, pour le fichier source
		self.typefile1=Label(root, text="Source file: ").grid(row=1, column=0)
		self.bar1=Entry(master)
		self.bar1.grid(row=1, column=1)
		
		# la seconde ligne, pour le fichier indicateur
		self.typefile2=Label(root, text="Indicator file: ").grid(row=2, column=0)
		self.bar2=Entry(master)
		self.bar2.grid(row=2, column=1)
		
		# LES BOUTONS WOOHOO (soyons enthousiastes, les gens aiment les boutons, non ?)  
		y=7
		self.cbutton= Button(root, text="OK", command=self.process_strategies)
		y+=1
		self.cbutton.grid(row=10, column=3, sticky = W + E)
		self.bbutton= Button(root, text="Browse", command= lambda: self.browsexlsx(0,self.bar1))
		self.bbutton.grid(row=1, column=3)
		self.bbutton= Button(root, text="Browse", command= lambda: self.browsexlsx(1,self.bar2))
		self.bbutton.grid(row=2, column=3)

	# fonction lancée quand on clique sur Browse
	def browsexlsx(self, filename_id, bar):
		Tk().withdraw()
		bar.delete(0, END)
		self.filenames[filename_id] = askopenfilename()
		bar.insert(10, self.extract_filename(self.filenames[filename_id]))

	# fonction pour afficher proprement le résultat du Browse, mise à part pour éviter de salir le code avec une grosse ligne moche
	# ça fait joli mais c'est probablement pas utilisable sur des OS Unix, à voir si on peut faire mieux
	def extract_filename(self, filepath):
		pathlist = filepath.split('/')
		return pathlist[-1]
		
	# fonction lancée quand on clique sur OK (c'est là qu'on veut mettre notre code mais on peut aussi juste appeler ton code à partir de là)
	# PLACEHOLDER A COMPLETER, du coup
	def process_strategies(self):
		print("Data source file path: ",self.filenames[0])
		print("Indicator file path: ",self.filenames[1])
		sys.exit()

root = Tk()
window=Window(root)
root.mainloop()  
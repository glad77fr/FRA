from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
import tkinter
import pandas as pd
import json
import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy


class Application(Tk):

    def __init__(self, parent):
        Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()
        self.title("Transformation de données")
        self.resizable(width=False, height=False)
        self.load_data()
        self.mainloop()

    def initialize(self):
        self.control_menu = Frame(self)
        self.control_menu.pack()

        self.conf_menu = LabelFrame(self.control_menu, text="Configuation", width=430, height=300, labelanchor='n')
        self.conf_menu.grid(row=1, column=1, pady=10, padx=10)

        self.rep_source_affiche = StringVar()
        self.rep_cible_affiche = StringVar()

        self.entree_label = Label(self.conf_menu, text="Répertoire source")
        self.entree_label.grid(row=1, column=1, padx=20,sticky=W)

        self.entree = Entry(self.conf_menu, textvariable=self.rep_source_affiche, width=50)
        self.entree.grid(row=1, column=2,sticky=W, padx=0)

        self.imp_Button = Button(self.conf_menu, text="Sélection...", command=self.import_file)
        self.imp_Button.grid(row=1, column=3, padx=15, pady=10,sticky=W)

        self.cible_label = Label(self.conf_menu, text="Répertoire cible")
        self.cible_label.grid(row=2,column=1, padx=20,sticky=W)

        self.cible_entree = Entry(self.conf_menu, textvariable=self.rep_cible_affiche, width=50)
        self.cible_entree.grid(row=2, column=2, sticky=W, padx=0)

        self.cible_Button = Button(self.conf_menu, text="Sélection...", command = self.rep_cible)
        self.cible_Button.grid(row=2, column=3, padx=15, pady=10,sticky=W)

        self.ex_menu = LabelFrame(self.control_menu, text="Execution", width=430, height=200, labelanchor='n')
        self.ex_menu.grid(row=1, column=2,padx=10, pady=10)

        self.ex_Button = Button(self.ex_menu, text="Executer", command = self.ex_programme,height = 1, width = 18)
        self.ex_Button.grid(row=3, column=1, padx=10, pady=10,sticky=W)

        self.quitButton = Button(self.ex_menu, text="Quitter", command=self.destroy,height = 1, width = 18)
        self.quitButton.grid(row=4, column=1, padx=10, pady=10,sticky=W)
        
        # Partie critères
        self.select_crit = LabelFrame(self.control_menu, text="Selection critères", width=430, height=300, labelanchor='n')
        self.select_crit.grid(row=2, column=1, pady=10, padx=10)
        
        # Critère 1
        
        # Label
        self.crt1_label = LabelFrame(self.select_crit, text="Critère 1", width=430, height=300, labelanchor='n')
        self.crt1_label.grid(row=1, column=1, padx=20, sticky=W)
        
        # Bouton 1 ajout critère
        self.copyButton = Button(self.crt1_label, text = ">>>", command = self.create_select )
        self.copyButton.grid(row=1, column=3, padx=2, sticky=W)
        
        # Bouton 2 retrait critère
        self.deleteButton = Button(self.crt1_label, text = "<<<", command = self.remove_select)
        self.deleteButton.grid(row=1, column=2, padx=2, sticky=W)
        
        # Liste de choix
        self.to_choose = Listbox(self.crt1_label)
        self.to_choose.grid(row=1, column=1, sticky=W)
        self.to_choose.insert( END, "Chat")
        self.to_choose.insert( END, "Chien")
        
        # List des selection
        self.chosen = Listbox(self.crt1_label)
        self.chosen.grid(row=1, column=4, sticky=W)
        self.chosen.insert(END, "Tous")
        
    def create_select(self):
        try:
            selected = self.to_choose.get(self.to_choose.curselection())
    
            if selected:
                self.chosen.insert(END, selected + "\n")
            
            # suppression de l'item dans la liste de selection
                idx = self.to_choose.get(0, END).index(selected)
                self.to_choose.delete(idx)
                
            # Suppression de vide dans chosen si plus d'une valeur
            if self.chosen.size()>1:
                idx = self.chosen.get(0, END).index("Tous")
                self.chosen.delete(idx)
            
        except:
            pass
        
    def remove_select(self):
        try:
            selected = self.chosen.get(self.chosen.curselection())
            
            if selected:
                self.to_choose.insert(END, selected + "\n")
                
            # suppression de l'item dans la liste de selection
                idx = self.chosen.get(0, END).index(selected)
                self.chosen.delete(idx)
                
            if self.chosen.size()==0:
                self.chosen.insert(END, "Tous")
                   
        except:
            pass
                      
    def import_file(self):
        rep = filedialog.askopenfilename(initialdir="/", title="Select file",
                                         filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if rep != "":
            self.rep_source_affiche.set(rep)
            self.source = pd.read_excel(rep)
            try:
                self.save_data["source"] = rep
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)
            except:
                self.save_data.update({"Source": rep})
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)
                                    
    def rep_cible(self):

        self.rep_cible = filedialog.askdirectory()
        if self.rep_cible != "":
            self.rep_cible_affiche.set(self.rep_cible)

            try:
                self.save_data["cible"] = self.rep_cible
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)
            except:
                self.save_data.update({"Cible": self.rep_cible })
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)

    def ex_programme(self):

        if self.rep_source_affiche.get() == "":
            messagebox.showerror("Ficher source non sélectionné", "Veuillez sélectionner un fichier source")

        if self.rep_cible_affiche.get() == "":
            messagebox.showerror("Répertoire source non sélectionné", "Veuillez sélectionner un répertoire source")

        if self.rep_source_affiche.get() != "" and self.rep_cible_affiche.get() != "":
            colonnes=[]
            colonnes = self.source.columns
            colonnes = list(colonnes)
            colonnes = str(colonnes)
            colonnes = colonnes[1:-1]

            messagebox.showinfo(title="Info colonnes", message="Les colonnes du fichier exporté sont: " + colonnes)

    def load_data(self):
        try:
            with open('data.txt') as json_file:
                self.save_data = json.load(json_file)

            try:
                print(self.save_data["cible"])

                self.rep_cible_affiche.set(self.save_data["cible"])

            except:

                pass

            try:

                self.rep_source_affiche.set(self.save_data["source"])

            except:

                pass

        except:
            self.save_data={}


# In[42]:


toto = Application(None)

toto.mainloop()


# In[ ]:





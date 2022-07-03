from tkinter import *
import tkinter as tk
from tkinter import filedialog,messagebox,ttk
import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import math
import os
from matplotlib import pyplot as plt
import xlwings as xl

root=tk.Tk()
root.geometry("540x340")
root.title("Excel_File_Treatment")
root.pack_propagate(False)
root.resizable(0,0)
root.configure(bg='grey')

my_menu = Menu(root)
root.config(menu=my_menu)

def about():
    top = Toplevel()
    top.geometry("700x100")
    top.title("À propos")
    l = Label(top, text="Hello !\n Bienvenue dans l'interface de Traitement de fichier Excel. \nGUIDE: Vous devez d'abord sélectionner le fichier avec l\'aide de la commande *Chercher* ,\n ensuite effectuer le chargement du fichier avec le boutton *Charger*,\n"
                        " enfin cliquer sur *Traiter*").pack()

option_menu = Menu(my_menu, tearoff=0)
my_menu.add_command(label="À propos",command=about)
my_menu.add_command(label="Quitter", command=root.destroy)

title=tk.Label(root,text ="Excel File Treater",font=("Arial",21),bg="darkgreen",fg="cyan")
title.place(x=140,y=0,width=250,height=50)
#Frame pour ouvrir un fichier
file_frame=tk.LabelFrame(root,text="OUVRIR UN FICHIER EXCEL",bg='cyan',bd=10,fg='darkgreen')
file_frame.place(x=5,y=110,width=530,height=175)

#Bouttons
button1=tk.Button(file_frame,text="Chercher",fg='green',command=lambda:File_dialog())
button1.place(rely=0.65,relx=0.65,width=175)

button2=tk.Button(file_frame,text="Charger",fg='green',command=lambda:Load_excel_data())
button2.place(rely=0.65,relx=0.325,width=183)

button3=tk.Button(file_frame,text="Traiter",fg='green',command=lambda:Treat_excel_file())
button3.place(rely=0.65,relx=0.01,width=175)

button4=tk.Button(root,text="Quitter",bg="darkred",fg="cyan",command =root.destroy)
button4.place(rely=0.90,relx=0.70,height=35,width=185)

label_file=ttk.Label(file_frame,text="                                                     Pas de fichier choisi")
label_file.place(x=40,width=450,height=30)

#Fonctions
def File_dialog():
	filename=filedialog.askopenfilename(initialdir="/",title="Selectionner un fichier",filetype=(("xlsx files","*.xlsx"),("All Files","*.*")))
	label_file["text"]=filename
	return filename
	tk.messagebox.showinfo(title="Recherche", message="Fichier Valide")
#**********************************************************************************
def Load_excel_data():
	file_path=label_file["text"]
	try:
		excel_filename=r"{}".format(file_path)
		global df
		df=pd.read_excel(excel_filename)
	except ValueError:
		tk.messagebox.showerror("Information","Le fichier choisi est invalide")
		return None
	except FileNotFoundError:
		tk.messagebox.showerror("Information",f"Pas de fichier tel {file_path}")
		return None
	tk.messagebox.showinfo(title="Chargement", message="Fichier Charge")
	return df
#**********************************************************************************
def Treat_excel_file():
	# filename=File_dialog()
	# label_file["text"]=filename
	# file_path=label_file["text"]
	# file_path=label_file["text"]
	# try:
	# 	excel_filename=r"{}".format(file_path)
	# 	df=pd.read_excel(excel_filename)
	# except ValueError:
	# 	tk.messagebox.showerror("Information","Le fichier choisi est invalide")
	# 	return None
	# except FileNotFoundError:
	# 	tk.messagebox.showerror("Information",f"Pas de fichier tel {file_path}")
	# 	return None
	#Nettoyage du fichier
    Series=['Nom','nom','NOM','noms','NOMS','Noms']
    idx=df.index[df[df.columns[1]].isin(Series)].tolist()[0]
    title=df.iloc[idx,:]
    df1=df.dropna(axis=0, how='any', thresh=None, subset=None, inplace=False)
    df2=df1.rename(columns=title)
    df2.columns=df2.columns.fillna("Decision")
    df2.rename(columns = {title[15]:'TotalCredits'}, inplace = True)
    df2.reset_index(drop=True)
    tri0_df2=df2.sort_values(by=['TotalCredits',title[1],title[2]], ascending=[False,True,True])
    tri1_df2=tri0_df2.reset_index(drop=True)
    #**********************************************************************************
    #Statistiques
    Eff=tri1_df2.Decision.value_counts()
    Eff=pd.DataFrame(Eff)
    Eff.rename(columns = {"Decision":"Nombre_Etudiants"}, inplace = True)
    Pourcentages= (Eff['Nombre_Etudiants'] /Eff['Nombre_Etudiants'].sum()) * 100
    #**********************************************************************************
    #Arrondissement des pourcentages
    P=[]
    for i in range(len(Pourcentages)):
        if i==1:
            x=round_up(Pourcentages[i])
            P.append(x)
        else:
            x=truncate(Pourcentages[i])
            P.append(x)
    Eff['Pourcentages'] = P
    Eff.loc["Total"]=[sum(Eff.Nombre_Etudiants),sum(Eff.Pourcentages)]
    #**********************************************************************************
    #Enregistrement du fichier
    writer=pd.ExcelWriter("Resultats_Deliberation.xlsx",engine="xlsxwriter")
    tri1_df2.to_excel(writer,"Résultats_Globaux")
    Eff.to_excel(writer,"Statistiques")
    writer.save()

    colors=('gold','cyan','pink','grey')
    wp={'linewidth':1,'edgecolor':'green'}
    fig,ax=plt.subplots(figsize=(10,7))
    wedges,texts,autotexts=ax.pie(Eff.Pourcentages[:4],labels=Eff[:4].index,explode=(0.05,0.05,0.05,0.05),autopct='%1.1f%%',startangle=90,shadow=True,wedgeprops=wp,textprops=dict(color='darkblue'),colors=colors)
    ax.legend(wedges,Eff[:4].index,title='Résultats Globaux',loc='center left',bbox_to_anchor=(1,0,0.5,1))
    plt.setp(autotexts,size=8,weight='bold')
    ax.set_title('Diagramme circulaire des Résultats Globaux')
    plt.axis('equal')
    xl_pf=xl.Book("Resultats_Deliberation.xlsx")
    sheet=xl_pf.sheets("Statistiques")
    sheet.pictures.add(ax.get_figure(),name="Statistiques",update=True)
    tk.messagebox.showinfo(title="Chargement", message="Fichier Traite")
#Arrondissemnt des pourcentages
def round_up(n, decimals=1):
    multiplier = 10 ** decimals
    return math.ceil(n * multiplier) / multiplier
def truncate(n, decimals=1):
    multiplier = 10**decimals
    return int(n* multiplier)/multiplier
#**********************************************************************************

root.mainloop()

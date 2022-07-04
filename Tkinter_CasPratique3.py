import tkinter.font as font
import tkinter as tk
import pandas as pd
import os
from tkinter.filedialog import askopenfilename
def Message(msg):
    message = tk.Tk()
    message.configure(background="white")
    label_message = tk.Label(message, text=msg, foreground = "black", background = "white")
    label_message['font'] = f_label
    label_message.pack()
    bouton_message = tk.Button(message, text="OK Merci", command = message.destroy)
    bouton_message.configure(foreground="black", background = "white")
    bouton_message['font'] = f_bouton
    bouton_message.pack()
    message.mainloop()
def traitement():
    chemin_fichier = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx'))])
    #creation d 1 nouvel fichier excel
    chemin_dossier = os.path.dirname(chemin_fichier)
    resultats = pd.read_excel(chemin_fichier, engine='openpyxl')
    Ligne9=(resultats.iloc[9]).values
    
     ## Suppression des 10 premières lignes
    resultats.drop(resultats.index[0:11],0,inplace=True)


    ## Remplacement de Resultas par Total_credits et Nan par Decision

    resultats.columns=Ligne9
    resultats=resultats.rename(columns={"RESULTAT":"Total_Credits"})
    resultats=resultats.dropna()
    resultats.columns=resultats.columns.fillna("Decision")
        #trie du data frame
    RESULTATS=resultats.sort_values(by=["Total_Credits","Nom","Prénom"],ascending=[False,True,True])

    
    
    if "données_modifiées.xlsx" in os.listdir(chemin_dossier):
        Message("Il y a déjà un fichier nommé 'données_modifiées.xlsx' dans le dossier, veuillez renommer ce fichier pour que le programme puisse en générer un nouveau.")
    else :
        
        RESULTATS.to_excel(str(chemin_dossier) + "\\" + "données_modifiées.xlsx","Deliberation_annuelle",index=False)
        Message("Votre fichier excel est généré Vous pouvez le trouver dans le même dossier que le fichier excel initial." + "\n" + "\n" + "A bientôt!" + "\n")
    
   
   


    
    R=resultats.Decision.value_counts()
    #Calcul des pourcentages
    pourcentage=[]
    for i in range (len(R)):
        p=R[i]*100/sum(R.values)
        pourcentage.append(p)




    PC=[]
    for i in range(len(pourcentage)):
        PC.append(round(pourcentage[i],1))





    max=PC[0]%1
    m=0
    for i in range(len(PC)):
        if PC[i]%1>max:
            max=PC[i]%1
            a=PC[i]
            m=i
    PC[m]=a+0.1



    max=PC[0]%1
    m=0
    for i in range(len(PC)):
        if PC[i]%1>max:
            max=PC[i]%1
            a=PC[i]
            m=i   
    PC[m]=a+0.1





    #creation du data frame
    dico1={"Valeur":R.values,"Pourcentage":PC}
    df1=pd.DataFrame(dico1, index=R.index) # Création du DataFrame
    L=[sum(R.values), sum(PC)]
    df1.loc["EffectifTotal"]=L
    
    #insertion du dataframe dans l fichier excel
    
    result=pd.ExcelWriter(str(chemin_dossier) + "\\" + "données_modifiées.xlsx", engine="openpyxl",mode="a")
    df1.to_excel(result,"Statistique",index=True)
    result.save()

    #diagramme
    import matplotlib.pyplot as plt
    labels = R.index
    serie = PC
    separation = (0, 0, 0, 0) # Séparation des tranches
    fig, ax = plt.subplots()
    ax.pie(serie, explode=separation, labels=labels, autopct='%1.0f%%',
    shadow=True, startangle=90)
    ax.axis('equal') # Equal aspect ratio ensures that pie is drawn as a circle.
    #plt.show()

    
    
    
    
    #enregistremnt diagramme
    import xlwings as xw
    RE=xw.Book(str(chemin_dossier) + "\\" + "données_modifiées.xlsx")
    #E.sheets.add("Statistique")
    E=RE.sheets["Statistique"]
    ax=ax.get_figure()
    E.pictures.add(ax,name="Statistique",update=True)
    
    
    
interface = tk.Tk()
interface.title("MODIFICATION D'UN FICHIER EXCEL")
interface.configure(background="white")

#Ici on affuche un message pour inviter l'utilisateur à selectionner un fichier
f_label = font.Font(family='Times New Roman', size=20)
f_bouton = font.Font(family='Times New Roman', size=15, weight="bold")
label = tk.Label(text="Veuillez sélectionner un Fichier Excel", foreground = "black", background = "white")

label = tk.Label(text="Bonjour! Merci de selectionner le fichier excel que vous voulez modifier.", foreground = "black", background = "white")
label['font'] = f_label
label.pack()

#Creation de bouton
bouton = tk.Button(text='Cliquez ici pour charger un fichier', command=traitement)
bouton.place(relx=0.200, rely=0.06, height=500, width=147)
bouton.configure(foreground="black")
bouton['font'] = f_bouton
bouton.pack(expand="yes")

#Message aprés les traitements
label2 = tk.Label(text="Un nouveau fichier excel est géneré. Merci de vérifier dans l'emplacement du fichier initial", foreground = "black", background = "white")
label2['font'] = f_label
label2.pack()

bouton2 = tk.Button(interface, text="J'ai fini merci!", command = interface.destroy)
bouton2.configure(foreground="black")
bouton2['font'] = f_bouton
bouton2.pack()
interface.mainloop()
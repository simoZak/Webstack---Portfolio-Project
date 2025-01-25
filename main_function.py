#******************* imports *****************************************************************
import os
import functions
import tkMessageBox
import re 
import tkFont
import tkFileDialog
from lxml import etree
from Tkinter import Button,Entry,StringVar,Tk
from Tkinter import Text, END, TclError, BOTTOM, X, Y, NONE, DISABLED,NORMAL
from Tkinter import Frame, Label, Scrollbar, HORIZONTAL, Toplevel
from Tkconstants import  LEFT,RIGHT 
import webbrowser
import xlrd
from  Treeview import XML_Viwer
from decimal import Decimal
import csv 
import datetime

#********************************************************************************
#              Interface : mise ef forme
#********************************************************************************
            
root = Tk()                            # interface parent
root.title("XML Tester - DEESSE GMP")  # titre de l'interface
root.wm_iconbitmap('PSA_logo.ico')     # icon
root.geometry("1100x560")              # dimensions de l'interface ("1270x750") 
root.minsize(1000, 700)                # minimum size
root.resizable(1, 0)                   # Don't allow resizing in the  y direction

frame1 =Frame(root,width=1000,height=80) # , bg="green" pour colorer le fond
frame1.pack(side='top', padx=5, pady=5)
frame1.place(x=0, y=0)

frame2 =Frame(root,width=1000,height=80)
frame2.pack(side='top', padx=5, pady=5)
frame2.place(x=300, y=0)

frame3 =Frame(root,width=1000,height=80)
frame3.pack(side='bottom', padx=0, pady=0)
frame3.place(x=830, y=652)

frame4 =Frame(root,width=300,height=440)
frame4.pack(side='right', padx=0 , pady=0)
frame4.place(x=540, y=55)

frame5 =Frame(root,width=80,height=60)
frame5.pack(side='left', padx=5, pady=0)
frame5.place(x=10, y=652)

framex =Frame(root,width=300,height=1000) 
framex.pack(side="left" , padx=0, pady=0) 
framex.place(x=10, y=55) 

w_font = tkFont.Font(family='Helvetica', size=14, weight='bold')  
w = Label(frame4, text="Erreurs trouvées !")
w['font'] = w_font
w.pack() 

xscrollbar = Scrollbar(frame4, orient=HORIZONTAL) 
xscrollbar.pack(side=BOTTOM, fill=X) # axe X

yscrollbar = Scrollbar(frame4)
yscrollbar.pack(side=RIGHT, fill=Y) # axe Y

text = Text(frame4, height = 33 , width= 67, wrap=NONE,xscrollcommand=xscrollbar.set,yscrollcommand=yscrollbar.set)
text.pack() 
      
xscrollbar.config(command=text.xview)
yscrollbar.config(command=text.yview)

w2 = Label(framex, text="      Arborescence      ", width= 42)
w2['font'] = w_font
w2.pack()
text_vide = Text(framex,  height = 33 ,width= 65)
text_vide.insert('1.0', "")
text_vide.pack() 

global j
filename = StringVar(root)
xmltree = StringVar(root)
FILETYPES = [ ("text files", "*.xml") ] # Afficher que les fichiers XML
FILETYPES_config = [ ("text files", "*.xlsx") ] # Afficher que les fichiers Excel

entry = Entry(frame2, textvariable=filename,font = "Helvetica 10 bold",  width=90) 
entry.pack(pady=15)

def set_filename():
    filename.set(tkFileDialog.askopenfilename(filetypes=FILETYPES))
    text.config(state= NORMAL)
    text.delete('1.0', END ) # vider la fenetre des erreurs à chaque click sur le bouton  Tester XML  
    xml_location = entry.get()
    if os.path.exists(xml_location):
        kids = framex.winfo_children()
        kids[1].destroy()  
        text_vide = Text(framex,  height= 33, width=65)
        text_vide.insert('1.0', "")
        text_vide.pack()            
    
def Open_Apropos (): 
    webbrowser.open("a_propos.txt")

# def Open_Config (): 
#     
#     webbrowser.open("config.xlsx")
#     tkMessageBox.showinfo(title="Alerte", message="N'oubliez pas de sauvegarder le fichier de config avant de lancer le Test !")    

def close_window (): 
    root.destroy()



# Fonction priciaple, traitment en appuyant sur le bouton 'Tester XML'    
def on_button():
    winE = Toplevel()
    winE.wm_iconbitmap('PSA_logo.ico')   
    winE.geometry("300x100")               
    Label(winE, text="\nTraitement en cours...",font=("Helvetica", 16)).pack(pady=10)    
#     t1 = int(round(time.time() * 1000)) # pour calculer le temps de traitement du fichier

    text_file = open("Output.txt", "w") # on cree ce fichier pour l'utiliser temporairemnt et le suprrimer directement
    text_file_1 = open("Output_1.txt", "w")# on cree ce fichier pour l'utiliser temporairemnt et le suprrimer directement
    
    #*********************************Read excel config file ************************
    xml_location = entry.get()
    print(entry.get())
    text.config(state= NORMAL)
    text.delete('1.0', END ) # pour vider la fenetre des erreurs à chaque click sur le bouton  Tester XML
    #filename.set(tkFileDialog.askopenfilename(filetypes=FILETYPES_config))
    
    #path = entry.get()
    path = "config.xlsx"
    book = xlrd.open_workbook(path)
    entry.delete(0, 'end')
    entry.insert(END, xml_location)
    #************************** Frame arborescence **********************************
       
    if os.path.exists(xml_location):
        j=0
        kids = framex.winfo_children()
        kids[1].destroy()    # supprimer tous les objets de framex
    try:
        tree = etree.parse(xml_location)
        xml = etree.tostring(tree, pretty_print=True)
        XML_Viwer(framex, xml).pack(fill='both',expand=True) # afficher l'arborescence du fichier XML
        framex.pack(side="left" , padx=10, pady=55 , fill="both")
    

    except etree.XMLSyntaxError as e:
        text.insert('1.0', "Error while parsing XML file : " + str(e))

#----------- liste des Noms normés -------------------------------------- 
#     second_sheet = book.sheet_by_index(1) # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    second_sheet = book.sheet_by_name('Feuil2') # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    
    rows_1 = second_sheet.nrows           # dernière ligne non vide 
    NN_Grandeur_liste = []
    for u in range(2, rows_1): 
        Nom_norme = str((second_sheet.cell(u,2)).value) 
        if Nom_norme <> "":
            NN_Grandeur_liste.append(Nom_norme)  
                  
#----------- liste TypeValeur -------------------------------------- 

#     second_sheet = book.sheet_by_index(1) # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    second_sheet = book.sheet_by_name('Feuil2') # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    liste_TypeValeur = []
    Cell_TypeValeur = int((second_sheet.cell(3,8)).value) 
    for v in range(1, Cell_TypeValeur+1):
        liste_TypeValeur.append(str(v))  
#----------- liste TypeGrandeur -------------------------------------- 

#     second_sheet = book.sheet_by_index(1) # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    second_sheet = book.sheet_by_name('Feuil2') # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    liste_TypeGrandeur = []
    Cell_TypeGrandeur = int((second_sheet.cell(3,4)).value) 
    for v in range(1, Cell_TypeGrandeur+1):
        liste_TypeGrandeur.append(str(v))     
#----------- liste SetType -------------------------------------- 

#     second_sheet = book.sheet_by_index(1) # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    second_sheet = book.sheet_by_name('Feuil2') # 2ème feuille (contenant la liste des noms normés) du fichier Excel 
    liste_SetType = []
    Cell_SetType = int((second_sheet.cell(3,12)).value) 
    for v in range(1, Cell_SetType+1):
        liste_SetType.append(str(v))
          
#---------------------------------------------------------------------------------------------  
    # Tableau des couples qui doivent exister en fonction de la valeur de la grandeur TYPREGUL
    col_1 = second_sheet.col_values(15)
    col_2 = second_sheet.col_values(16)
    col_3 = second_sheet.col_values(17)
    TYPREGUL_values = []
    couple_values_1 = []
    couple_values_2 = []
    liste_OR = []
    liste_OR_couple = []
    for c in range (4 , rows_1):
        if col_1[c]== '' and  col_2[c]== '' and  col_3[c]== '':
            for i in range (4 , c  ): # c dernière valeur non vide 
                TYPREGUL_values.append(int(col_1[i]))
                couple_values_1.append(str(col_2[i]))
                couple_values_2.append(str(col_3[i]))
            break
             
#---------------------------------------------------------------------------------------------  
    # Tableau des valeurs pour l'exeption "n'est pas numérique"
    col_categorie = second_sheet.col_values(19)
    col_valeur = second_sheet.col_values(20)
    liste_categorie = []
    liste_valeur = []
    for c in range (4 , rows_1):
        if col_categorie[c]== '' and  col_valeur[c]== '':
            for i in range (4 , c  ):
                liste_categorie.append(str(col_categorie[i]))
                liste_valeur.append(str(col_valeur[i]))
            break  
#------------------------------------------------------------------------   

#     first_sheet = book.sheet_by_index(0)  # 1ère feuille
    first_sheet = book.sheet_by_name('Feuil1') # 1ere feuille (contenant la liste de règles à vérifier)
    rows_sheet_1 = first_sheet.nrows      # dernière ligne non vide   
    for u in range(1, rows_sheet_1): # 
        element  = str((first_sheet.cell(u,2)).value)
        type_regle = str((first_sheet.cell(u,1)).value)
        

  
    

#   **********************************************************************************



        #Rechercher tous les paths possibles d'un  Element 
        lista = []
        element_tree = element
        for  element_tree in tree.iter():
            paths = tree.getpath(element_tree)
            t = re.sub(r'\[.*?\]', '', paths)
            l = len(element)
            i = t.find(element)
            tt = t[:i+l]
            if element in tt: 
                if tt not in lista:
                    lista.append(tt)
                    
#         print len(list) ," chemins pour : ", element
        
        erreurs =''
        erreurs1=''
        erreurs2=''
        erreurs3=''
        erreurs4=''
        erreurs6=''
        erreurs5=''
        erreurs7=''
        erreurs8='' 
        erreurs9=''  
        erreurs10=''     
        erreurs11=''  
        erreurs12=''
        erreurs13=''      
        erreurs14=''  
        erreurs15=''  
        erreurs16=''    
        erreurs17=''   
        erreurs18=''
        erreurs19=''  
        erreurs20=''
        erreurs21=''  
        erreurs22=''
        erreurs23=''   
        erreurs24=''
        erreurs25='' 
        erreurs26=''  
        erreurs27='' 
        erreurs28=''  
        erreurs29='' 
        erreurs30=''                                                                                 
        for elem in lista:
            print elem  , "***************************************************************"   
                
            if type_regle == "type_1":
                attribut = str((first_sheet.cell(u,3)).value)
                type_attribut = str((first_sheet.cell(u,4)).value)   
                indication =''   
                for toto in tree.xpath(elem): 
                    
                    if 'Categorie' in str(toto.getparent()):
                        if 'Campagne' in str((toto.getparent()).getparent()) :
                            if "nom" in ((toto.getparent()).getparent()).attrib:
                                indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]" 
                            else:
                                indication = "/Campagne/Categorie["+(toto.getparent()).attrib['Description']+"]" 
                            if "Nom" in toto.attrib:   
                                indication = indication + "/"+element + "["+toto.attrib['Nom']+"]"
                            elif "Description" in toto.attrib:   
                                indication = indication + "/" + element + "["+toto.attrib['Description']+"]"                                                                                                             
                        elif 'Point' in  str((toto.getparent()).getparent()):
                            if "nom" in (((toto.getparent()).getparent()).getparent()).attrib:
                                indication = "/Campagne["+(((toto.getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]" 
                            else:
                                indication = "/Campagne/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]"
                            if "Nom" in toto.attrib:
                                indication = indication + "/"+element + "["+toto.attrib['Nom']+"]"
                            elif "Description" in toto.attrib:   
                                indication = indication + "/"+element + "["+toto.attrib['Description']+"]"    
                    if 'Consigne' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib :
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib :
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                        
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                        elif 'ProcedureOnline' in  str((toto.getparent()).getparent()):
                            if 'Point' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib :
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/ProcedureOnline[Description="+((toto.getparent()).getparent()).attrib['Description']+"]" + "/Consigne" 
                                else:
                                    indication = "/Campagne/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/ProcedureOnline[Description="+((toto.getparent()).getparent()).attrib['Description']+"]" + "/Consigne" 
                                     
                                if "Nom" in toto.attrib:
                                    indication = indication + "/"+element + "["+toto.attrib['Nom']+"]"
                            
                    if 'Recopie' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib : 
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                    
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib :
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Mesure' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie[Description="+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie[Description="+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                           
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Activation' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:    
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                    
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib :
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Consignes' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"']"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if "nom" in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                        
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()):
                                if "nom" in (((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib :
                                    indication = "/Campagne["+(((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"+  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consignes"
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"+  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consignes"
                                    
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                            
                    if 'Variation' in str(toto.getparent()):
                        if 'Niveau' in str((toto.getparent()).getparent()) :
                            if 'Variations' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                    
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Point' in str(toto.getparent()):
                        if 'Niveau' in str((toto.getparent()).getparent()) :
                            if 'Variations' in str(((toto.getparent()).getparent()).getparent()) :
                                if "nom" in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                    
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"   
                    indication = indication.encode(encoding='UTF-8',errors='strict')                 
#                     try:
#                         indication = unicode(indication, 'utf-8')
#                     except TypeError:
#                         indication =  indication                          
                            
                    if attribut  in toto.attrib:
                        if  type_attribut == "estnonvide":
                            if toto.attrib[attribut] == '' or  (toto.attrib[attribut]).isspace() == True:
                                erreur1 = type_regle +" :- la balise '" +  indication + " contient un attribut '" + attribut + "', mais il est vide"
                                erreurs1 = erreurs1  + erreur1 + "\n"                               
                        else:    
                            if type_attribut == "isnumeric":
                                test_type = str(toto.attrib[attribut])
                                if  functions.isInt(test_type)== False and functions.isFloat(test_type)== False:
                                    erreur1 = type_regle +" :- l'attribut '"+ attribut +"' de la balise " + indication + "  n'est pas numerique"
                                    erreurs1 = erreurs2 + erreur1 + "\n"
                            if type_attribut == "isalphanumeric":
                                test_type = unicode(toto.attrib[attribut], 'utf-8')
                                if test_type.isalnum()== False :
                                    erreur1 =type_regle +" :- l'attribut '"+ attribut +"' de la balise " +  indication + " n'est pas alphanumerique"
                                    erreurs1 = erreurs1 + erreur1 + "\n" 
                    else:
                        erreur1 = type_regle +" :- la balise "+ indication + "  ne contient pas l'attribut : '" + attribut + "'"
                        erreurs1 = erreurs1  + erreur1 + "\n"      
            #-----------------------------------------------------------------------------------------------------------------------------------------------------            
            if type_regle == "type_2": 
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value) 
                balise_3 = str((first_sheet.cell(u,4)).value) 
                indication = ''
                for toto in tree.xpath(elem): 
                    if 'Point' in str(toto.getparent()):
                        if 'nom' in ((toto.getparent()).getparent()).attrib:
                            indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Point["+(toto.getparent()).attrib['Type']+"]" + "/Variations/" 
                        else:     
                            indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]" + "/Variations/" 
                    indication = indication.encode(encoding='UTF-8',errors='strict') 
                          
                    if balise_2  in str(toto.getchildren()):
                        for titi in toto:
                            if balise_3 not in str(titi.getchildren())  :
                                erreur28 = type_regle +" :- la balise " + indication + " contient la balise '"+  balise_2 + "' mais celle ci  ne contient pas la balise '"+ balise_3 +"'"
                                erreurs28 = erreurs28 + erreur28 + "\n"
            #-----------------------------------------------------------------------------------------------------------------------------------------------------                      
                                     
            if type_regle == "type_3": 
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value) 
                balise_3 = str((first_sheet.cell(u,4)).value)
                indication ='' 
                for toto in tree.xpath(elem): 
                  
                    if balise_2  not in str(toto.getchildren()):
                        erreur3 = type_regle +" :- la balise "  + indication +" ne contient pas la balise "+  balise_2 + "'"
                        erreurs3 = erreurs3 + erreur3 + "\n" 
                        
                    else:
                        for titi in toto : # tree.xpath(elem+"/"+balise_2):
                            toto = titi.getparent()
                            if 'Variation' in str(toto.getparent()):
                                if 'Niveau' in str((toto.getparent()).getparent()) :
                                    if 'Variations' in str(((toto.getparent()).getparent()).getparent()) :
                                        if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                            indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                        else:
                                            indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                            
                                        if "Nom" in toto.attrib :
                                            indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                            if 'Categorie' in str(toto.getparent()):
                                if 'Campagne' in str((toto.getparent()).getparent()) :
                                    if 'nom' in ((toto.getparent()).getparent()).attrib:
                                        indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Categorie[Description='"+(toto.getparent()).attrib['Description']+"]"
                                    else:
                                        indication = "/Campagne/Categorie[Description='"+(toto.getparent()).attrib['Description']+"]"
                                        
                                    if "Nom" in toto.attrib:   
                                        indication = indication + "/"+element + "["+toto.attrib['Nom']+"]"
                            if 'Point' in str(toto.getparent()):
                                if 'nom' in ((toto.getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Point["+(toto.getparent()).attrib['Type']+"]"  + "/ProcedureOnline["+toto.attrib['Description']+"]"
                                else:
                                    indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]"  + "/ProcedureOnline["+toto.attrib['Description']+"]"
                                        
                            indication = indication.encode(encoding='UTF-8',errors='strict')  
                                                
                            if balise_3 not in str(titi.getchildren())  and balise_2 in str(titi) :
                                erreur3 = type_regle +" :- la balise "  + indication +" contient la balise "+  balise_2 + " mais celle ci  ne contient pas la balise "+ balise_3 +"'"
                                erreurs3 = erreurs3 + erreur3 + "\n"
      
            #-----------------------------------------------------------------------------------------------------------------------------------------------------      
                             
            if type_regle == "type_4":
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value) 
                balise_3 = str((first_sheet.cell(u,4)).value)
                indication =''
                
                for toto in tree.xpath(elem): 
                    if "demande" in str(toto):
                            if 'id_utilisateur' and 'utilisateur' in toto.attrib:
                                if 'id_utilisateur' in toto.attrib:
                                    indication = "/demande[id_utilisateur='"+toto.attrib['id_utilisateur']+"' , utilisateur='"+ toto.attrib['utilisateur']+"']" 
                                else:
                                    indication = "/demande[utilisateur='"+ toto.attrib['utilisateur']+"']" 
                                        
                            else:
                                indication = balise_1
                    if 'demande' in str(toto.getparent()):
                            if 'id_utilisateur' in (toto.getparent()).attrib  and 'nom' in toto.attrib :
                                indication = "/demande["+(toto.getparent()).attrib['id_utilisateur']+" , "+ (toto.getparent()).attrib['utilisateur']+"]" +"/Campagne["+toto.attrib['nom']+"]"
                            else:
                                indication = "/demande["+ (toto.getparent()).attrib['utilisateur']+"]" +"/Campagne"
                                      
                    if "Point" in str(toto) and 'Campagne' in str(toto.getparent()):
                        if 'nom' in (toto.getparent()).attrib:
                            indication = "/Campagne["+(toto.getparent()).attrib['nom']+"]" +  "/Point["+toto.attrib['Type']+"]"
                        else:
                            indication = "/Campagne/Point["+toto.attrib['Type']+"]"
                                  
                    if "Niveau" in str(toto) and 'Variations' in str(toto.getparent()):
                            if 'Point' in 'Variations' in str((toto.getparent()).getparent()):
                                if 'nom' in (((toto.getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((toto.getparent()).getparent()).getparent()).attrib['nom']+"]" +  "/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau"
                                else:
                                    indication = "/Campagne/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau"
                                            
                    if "Variation" in str(toto):
                        if 'Point' in  str(toto.getparent()):
                            if 'nom' in ((toto.getparent()).getparent()).attrib:
                                indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" +  "/Point["+(toto.getparent()).attrib['Type']+"]" + "/Variation"  
                            else:
                                indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]" + "/Variation"  
                                
                        if 'Niveau' in  str(toto.getparent()):
                            if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]" +  "/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"
                            else:
                                indication = "/Campagne/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"
                                     
                    if "ProcedureOnline" in str(toto):
                        if 'Point' in  str(toto.getparent()):
                            if 'nom' in  ((toto.getparent()).getparent()).attrib:
                                indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" +  "/Point["+(toto.getparent()).attrib['Type']+"]" +  "/ProcedureOnline[Description="+toto.attrib['Description']+"]"
                            else:
                                indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]" +  "/ProcedureOnline[Description="+toto.attrib['Description']+"]"
                                     
                    if "Categorie" in str(toto):
                        if 'Point' in  str(toto.getparent()):
                                if 'nom' in ((toto.getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" +  "/Point["+(toto.getparent()).attrib['Type']+"']" +  "/Categorie[Description='"+toto.attrib['Description']+"]" 
                                else:
                                    indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"']" +  "/Categorie[Description='"+toto.attrib['Description']+"]" 
                                    
                        if 'Campagne' in  str(toto.getparent()):
                            if 'nom' in (toto.getparent()).attrib:
                                indication = "/Campagne["+(toto.getparent()).attrib['nom']+"]"  +  "/Categorie["+toto.attrib['Description']+"]" 
                            else:
                                indication = "/Campagne/Categorie["+toto.attrib['Description']+"]" 
                                    
                    indication = indication.encode(encoding='UTF-8',errors='strict') 
                    
                    if balise_3 == "":                                                                                                  
                        if balise_2  not in str(toto.getchildren()):
                            erreur9 = type_regle +" :- la balise " + indication  + " n'a pas la balise '"+  balise_2 + "'"
                            erreurs9 = erreurs9 + erreur9 + "\n"    
                    else: 
                        if  balise_2 not in str(toto.getchildren()) and balise_3  not in str(toto.getchildren()):
                            erreur9 = type_regle +" :- la balise '" + balise_1 + indication  + "' n'a ni la balise '"+  balise_2 + "' ni la balise '"+  balise_3  + "'"
                            erreurs9 = erreurs9 + erreur9 + "\n"       

            #-----------------------------------------------------------------------------------------------------------------------------------------------------         

            if type_regle == "type_5":
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value)
                nom_attribut = str((first_sheet.cell(u,5)).value) 
                val_attribut_souhaite = str((first_sheet.cell(u,6)).value) 
                indication = ''
                for toto in tree.xpath(elem): 

                    if 'Point' in str(toto):
                        if 'nom' in (toto.getparent()).attrib:
                            indication = "/Campagne["+(toto.getparent()).attrib['nom']+"]" + "/Point["+toto.attrib['Type']+"]"   
                        else:    
                            indication = "/Campagne/Point["+toto.attrib['Type']+"]"  
                    indication = indication.encode(encoding='UTF-8',errors='strict') 
                    
                    if balise_2  not in str(toto.getchildren()):
                        erreur10 = type_regle +" :- la balise '"  + indication +" ne contient pas  la balise "+  balise_2 + "'"
                        erreurs10 = erreurs10 + erreur10 + "\n"   
                    else:
                        if not any(titi.attrib[nom_attribut] == val_attribut_souhaite for titi in tree.xpath(elem+"/"+balise_2)):
                            erreur11 = type_regle +" :- la balise "  + indication +" ne contient pas la balise "+  balise_2 +" dont [" + nom_attribut +"='" + val_attribut_souhaite + "']'"
                            erreurs11 = erreurs11 + erreur11 + "\n" 
            #----------------------------------------------------------------------------------------------------------------------------------------------------- 

            if type_regle == "type_6":
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value)
                nom_attribut = str((first_sheet.cell(u,5)).value) 
                val_attribut_souhaite = str((first_sheet.cell(u,6)).value) 
                texte_non_souhaite = str((first_sheet.cell(u,7)).value) 
  
                for titi in tree.xpath(elem+"/"+balise_2):  
                    toto = titi.getparent()
                    if 'Categorie' in str(toto.getparent()):
                        if 'Campagne' in str((toto.getparent()).getparent()) :
                            if 'nom' in ((toto.getparent()).getparent()).attrib:
                                indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]"
                            else:
                                indication = "/Campagne/Categorie["+(toto.getparent()).attrib['Description']+"]"
                                    
                            if "Nom" in toto.attrib:   
                                indication = indication + "/"+element + "["+toto.attrib['Nom']+"']"
                            elif "Description" in toto.attrib:   
                                indication = indication + "/" + element + "["+toto.attrib['Description']+"]"                                                                                                             
                        elif 'Point' in  str((toto.getparent()).getparent()):
                            if 'nom' in (((toto.getparent()).getparent()).getparent()).attrib:
                                indication = "/Campagne["+(((toto.getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]"
                            else:
                                indication = "/Campagne/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]"
                                     
                            if "Nom" in toto.attrib:
                                indication = indication + "/"+element + "["+toto.attrib['Nom']+"]"
                            elif "Description" in toto.attrib:   
                                indication = indication + "/"+element + "["+toto.attrib['Description']+"]"    
                    if 'Consigne' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                    
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib :
                                    indication = "/Campagne[nom='"+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"']"+  "/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"']" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                        
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                        elif 'ProcedureOnline' in  str((toto.getparent()).getparent()):
                            if 'Point' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/ProcedureOnline["+((toto.getparent()).getparent()).attrib['Description']+"]" + "/Consigne" 
                                else:
                                    indication = "/Campagne/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/ProcedureOnline["+((toto.getparent()).getparent()).attrib['Description']+"]" + "/Consigne" 
                                    
                                if "Nom" in toto.attrib:
                                    indication = indication + "/"+element + "["+toto.attrib['Nom']+"]"
                            
                    if 'Recopie' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"']" +  "/Point[Type='"+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point[Type='"+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Mesure' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                    
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Activation' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                     
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib :
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                        
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Consignes' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib :
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib :
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                    
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()):
                                if 'nom' in (((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"+  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consignes"
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"+  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consignes"
                                    
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                            
                    if 'Variation' in str(toto.getparent()):
                        if 'Niveau' in str((toto.getparent()).getparent()) :
                            if 'Variations' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" +element + "["+toto.attrib['Nom']+"]"
                    if 'Point' in str(toto.getparent()):
                        if 'Niveau' in str((toto.getparent()).getparent()) :
                            if 'Variations' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                    
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" +element + "[Nom='"+toto.attrib['Nom']+"']"    
                    indication = indication.encode(encoding='UTF-8',errors='strict') 
                                    
                    if nom_attribut  in (titi.getparent()).attrib:           
                        if  (titi.getparent()).attrib[nom_attribut] ==  val_attribut_souhaite :
                              
                            if texte_non_souhaite in titi.text:
                                erreur13 = type_regle +" :- la balise " + indication + " contient '"+  texte_non_souhaite + "' dans sa valeur"
                                erreurs13 = erreurs13 + erreur13 + "\n" 

            #-----------------------------------------------------------------------------------------------------------------------------------------------------                     
                    
            if type_regle == "type_7": 
                balise_2 = str((first_sheet.cell(u,2)).value)
                balise_1 = str((first_sheet.cell(u,3)).value) 
                balise_3 = str((first_sheet.cell(u,4)).value)
                born_min_t = str(int((first_sheet.cell(u,7)).value))
                born_max_t = str(int((first_sheet.cell(u,8)).value)) 
                min_isnumeric = False
                max_isnumeric = False
                indication =''
                
                for toto in tree.xpath(elem): 
                    if 'Activation' in str(toto) and 'Structure' in str(toto.getparent()):
                        if 'Point'  in str(((toto.getparent()).getparent()).getparent()):
                            if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+ "/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+((toto.getparent()).getparent()).attrib['Description']+"]"+ "/Structure["+(toto.getparent()).attrib['Nom']+"]"
                            else:
                                indication = "/Campagne/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+((toto.getparent()).getparent()).attrib['Description']+"]"+ "/Structure["+(toto.getparent()).attrib['Nom']+"]"
                            
                        if 'Campagne'  in str(((toto.getparent()).getparent()).getparent()):
                            if 'nom' in (((toto.getparent()).getparent()).getparent()).attrib:
                                indication = "/Campagne["+(((toto.getparent()).getparent()).getparent()).attrib['nom']+"]" + "/Categorie["+((toto.getparent()).getparent()).attrib['Description']+"]"+ "/Structure["+(toto.getparent()).attrib['Nom']+"]"
                            else:
                                indication = "/Campagne/Categorie["+((toto.getparent()).getparent()).attrib['Description']+"]"+ "/Structure["+(toto.getparent()).attrib['Nom']+"]"
                                    
                        indication = indication.encode(encoding='UTF-8',errors='strict') 
                                                        
                    if balise_1  in str(toto.getparent()) and balise_3 in str(toto.getchildren()):
                        for tete in toto: 
                            for titi in tete:
                                             
                                if  titi.text <> None : 
                                    value = (titi.text).encode('utf-8')
                                    value1 = unicode(value, 'utf-8')
                                    
                                   
                                    born_min_t = born_min_t.encode('utf-8')
                                    born_min_t = unicode(born_min_t, 'utf-8')
                                    born_max_t = born_max_t.encode('utf-8')
                                    born_max_t = unicode(born_max_t, 'utf-8')                                
                                    if value1 <> "" and value1.isspace() == False: 
                                        if born_min_t <> "" and born_max_t <> "" and born_min_t.isspace() == False and born_max_t.isspace() == False: # limites rensignés dans le tableau
                                            
                                            if  functions.isInt(born_min_t):
                                                born_min_tab =  Decimal (born_min_t, 'utf-8')
                                                born_min_tab =  int (born_min_tab)
                                                min_isnumeric = True
                                            elif  functions.isFloat(born_min_t):   
                                                born_min_tab =  Decimal (born_min_t, 'utf-8')
                                                born_min_tab =  float (born_min_tab)
                                                min_isnumeric = True
                                            if    functions.isInt(born_max_t):
                                                born_max_tab =  Decimal (born_max_t, 'utf-8')
                                                born_max_tab =  int (born_max_tab)
                                                max_isnumeric = False
                                            elif  functions.isFloat(born_max_t):
                                                born_max_tab =  Decimal (born_max_t, 'utf-8')
                                                born_max_tab =  float (born_max_tab)  
                                                max_isnumeric = False  
                                                
                                            if  min_isnumeric == True  and    max_isnumeric == True :                                         
                                                if functions.isInt(value1):
                                                    val = int (value1)                                                   
                
                                                    if ( val >= born_min_tab and val <= born_max_tab ) == False:  
                                                        erreur14 = type_regle +" :- la balise  "  + indication + " a une valeur  qui n'est pas comprise entre les bornes min et max : '" + value + "'" 
                                                        erreurs14 = erreurs14 + erreur14 + "\n" 
                                                            
                                                elif functions.isFloat(value1):
                                                    val = float (value1)                                                    
                
                                                    if ( val >= born_min_tab and val <= born_max_tab ) == False:  
                                                        erreur14 = type_regle +" :- la balise  " + indication + " a une valeur  qui n'est pas comprise entre les bornes min et max : '" + value + "'" 
                                                        erreurs14 = erreurs14 + erreur14 + "\n" 
                                                else : 
                                                    erreur14 = type_regle+" :- la balise  " + indication + " a une valeur  qui n'est pas numerique : '" + value + "'" 
                                                    erreurs14 = erreurs14 + erreur14 + "\n"  
                                            else : 
                                                erreur14 = type_regle +" :- les valeurs Min et Max renseignees dans le fichier de config ne sont pas numeriques Min='"+ born_min_t +"', Max='" + born_max_t 
                                                erreurs14 = erreurs14 + erreur14 + "\n"  
                                        else :
                                            erreur14 = type_regle +" :- les valeurs Min et Max renseignées dans le fichier de config sont vides "
                                            erreurs14 = erreurs14 + erreur14 + "\n"                                          
                                    else :
                                        erreur14 = type_regle+" :- la balise  " +  indication + " a une valeur qui est vide"
                                        erreurs14 = erreurs14 + erreur14 + "\n"                    
            #-----------------------------------------------------------------------------------------------------------------------------------------------------                

            if type_regle == "type_8":
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value)
                balise_3 = str((first_sheet.cell(u,4)).value) 
                nom_attribut = str((first_sheet.cell(u,5)).value) 
                val_attribut_souhaite = str((first_sheet.cell(u,6)).value) 
                indication =''             
                for toto in tree.xpath(elem): 
                    
                    if 'Variations' in str(toto) and 'Point' in str(toto.getparent()):
                        if 'nom' in ((toto.getparent()).getparent()).attrib:
                            indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Point["+(toto.getparent()).attrib['Type']+"]"+ "/Variations"
                        else:
                            indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]"+ "/Variations"
                                
                    indication = indication.encode(encoding='UTF-8',errors='strict') 
                                            
                    if balise_3 in toto.getparent() and ((toto.getparent()).attrib[nom_attribut])<>val_attribut_souhaite and balise_2 not in str(toto.getchildren()):
                        erreur16 = type_regle +" :- la balise " + indication + " ne contient pas la balise "+  balise_2  + "'"
                        erreurs16 = erreurs16 + erreur16 + "\n"                  
            #-----------------------------------------------------------------------------------------------------------------------------------------------------
            
            if type_regle == "type_9":  
                balise_1 = str((first_sheet.cell(u,2)).value)
                valeur_non_souhaite = str((first_sheet.cell(u,7)).value)
                indication =''
                for titi in tree.xpath(elem): 
                    
                    toto = titi.getparent()
                    if 'Categorie' in str(toto.getparent()):
                        if 'Campagne' in str((toto.getparent()).getparent()) :
                            if 'nom' in ((toto.getparent()).getparent()).attrib:
                                indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]"
                            else:
                                indication = "/Campagne/Categorie["+(toto.getparent()).attrib['Description']+"]"
                                
                            if "Nom" in toto.attrib:   
                                indication = indication + "/" + "Grandeur["+toto.attrib['Nom']+"]"
                            elif "Description" in toto.attrib:   
                                indication = indication + "/" + "Grandeur["+toto.attrib['Description']+"]"                                                                                                             
                        elif 'Point' in  str((toto.getparent()).getparent()):
                            if 'nom' in (((toto.getparent()).getparent()).getparent()).attrib:
                                indication = "/Campagne["+(((toto.getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]" 
                            else:
                                indication = "/Campagne/Point["+((toto.getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(toto.getparent()).attrib['Description']+"]" 
                                
                            if "Nom" in toto.attrib:
                                indication = indication + "/"+ "Grandeur["+toto.attrib['Nom']+"]"
                            elif "Description" in toto.attrib:   
                                indication = indication + "/"+ "Grandeur["+toto.attrib['Description']+"]"    
                    if 'Consigne' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                
                            if "Nom" in toto.attrib :
                                indication = indication + "/" + "Grandeur["+toto.attrib['Nom']+"]"
                        elif 'ProcedureOnline' in  str((toto.getparent()).getparent()):
                            if 'Point' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/ProcedureOnline["+((toto.getparent()).getparent()).attrib['Description']+"]" + "/Consigne" 
                                else:
                                    indication = "/Campagne/Point["+(((toto.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/ProcedureOnline["+((toto.getparent()).getparent()).attrib['Description']+"]" + "/Consigne" 
                                    
                                if "Nom" in toto.attrib:
                                    indication = indication + "/"+ "Grandeur["+toto.attrib['Nom']+"]"
                            
                    if 'Recopie' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"']"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if "Nom" in toto.attrib :
                                indication = indication + "/" + "Grandeur["+toto.attrib['Nom']+"]"
                    if 'Mesure' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                    
                            if "Nom" in toto.attrib :
                                indication = indication + "/" + "Grandeur["+toto.attrib['Nom']+"]"
                    if 'Activation' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne[nom='"+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"']"+"/Point[Type='"+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"']" + "/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"
                                else:
                                    indication = "/Campagne/Point[Type='"+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"']" + "/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne[nom='"+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"']"+  "/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie[Description='"+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"']" +  "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne[nom='"+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"']" +  "/Point[Type='"+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"']" +  "/Variations/Niveau/Variation" + "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point[Type='"+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"']" +  "/Variations/Niveau/Variation" + "/Structure[Nom='"+((toto.getparent()).getparent()).attrib['Nom']+"']" + "/Consigne"    
                                        
                            if "Nom" in toto.attrib :
                                indication = indication + "/" + "Grandeur[Nom='"+toto.attrib['Nom']+"']"
                    if 'Consignes' in str(toto.getparent()):
                        if 'Structure' in str((toto.getparent()).getparent()) :
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+"/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consignes"
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                        
                            if 'Categorie' in str(((toto.getparent()).getparent()).getparent()) and 'Point' not in str((((toto.getparent()).getparent()).getparent()).getparent()):
                                if 'nom' in ((((toto.getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                else:
                                    indication = "/Campagne/Categorie["+(((toto.getparent()).getparent()).getparent()).attrib['Description']+"]" +  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"
                                            
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in ((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib:
                                    indication = "/Campagne["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent().attrib['nom']+"]" +  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" +  "/Variations/Niveau/Variation" + "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consigne"    
                                            
                            if "Nom" in toto.attrib :
                                indication = indication + "/" +element + "[Nom='"+toto.attrib['Nom']+"']"
                            if 'Variation' in str(((toto.getparent()).getparent()).getparent()):
                                if 'nom' in (((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"+  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consignes"
                                else:
                                    indication = "/Campagne/Point["+((((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"+  "/Structure["+((toto.getparent()).getparent()).attrib['Nom']+"]" + "/Consignes"
                                        
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" + "Grandeur["+toto.attrib['Nom']+"]"
                            
                    if 'Variation' in str(toto.getparent()):
                        if 'Niveau' in str((toto.getparent()).getparent()) :
                            if 'Variations' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                    
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" + "Grandeur["+toto.attrib['Nom']+"]"
                    if 'Point' in str(toto.getparent()):
                        if 'Niveau' in str((toto.getparent()).getparent()) :
                            if 'Variations' in str(((toto.getparent()).getparent()).getparent()) :
                                if 'nom' in (((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib:
                                    indication = "/Campagne["+(((((toto.getparent()).getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]"+  "/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                else:
                                    indication = "/Campagne/Point["+((((toto.getparent()).getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations"+ "/Niveau" + "/Variation" 
                                    
                                if "Nom" in toto.attrib :
                                    indication = indication + "/" + "Grandeur["+toto.attrib['Nom']+"]"    
                                             
                    indication = indication.encode(encoding='UTF-8',errors='strict')

                                                         
                    if (titi.text) is not None : 
                        value = (titi.text).encode('utf-8')              
                        value1 = unicode(value, 'utf-8')
                        langueur = len(value1)
                    
                    if langueur > 32:
                        erreur17 = type_regle +" :- la valeur de la balise " + indication + " a  plus de 32 caractères :'" + value + "'"
                        erreurs17 = erreurs17 + erreur17 + "\n"                     
                    elif  langueur <> 0 :   
                        espace_debut = value1[0]
                        espace_fin = value1[langueur-1]
                    
                    if value1 == "" or value1.isspace() <> False:     
                        erreur17 = type_regle +" :- la valeur de la balise " + indication + " a une donnée vide " 
                        erreurs17 = erreurs17 + erreur17 + "\n" 

                    elif espace_debut == " " or espace_fin == " ":
                        erreur17 = type_regle + " :- la valeur de la balise " + indication + "contient un espace au début ou a la fin : '" + value + "' "
                        erreurs17 = erreurs17 + erreur17 + "\n" 
                        
                    elif  functions.isInt(value1)==False or  functions.isFloat(value1)== False:
#                         listee = []
                        listee = [(a.start()) for a in list(re.finditer(',', value1))]
                        listee.append(len(value1))
                        listee.insert(0,0)
                        listee_count = len(listee)
                        
                         
                        for jj in range(1,listee_count -1  ): # pour tester s'il y'a un espace après chaque virgule
                            if value1[listee[jj]+1] <> " ":
                                erreur17 = type_regle +" :- la valeur de la balise " + indication + " ne contient pas un espace après la virgule : '" + value + "' "
                                erreurs17 = erreurs17 + erreur17 + "\n" 
                        
                        for jj in range(0,listee_count -1 ):   # pour tester le Nbre des points entre 2 virgules successives  
                            s = value1[listee[jj]: listee[jj+1]]
                            if s.count('.') > 1 : 
                                erreur17 = type_regle +" :- une variable dans la liste des valeurs de la balise " + indication + " contient plusieurs '.' : '" + value + "' "
                                erreurs17 = erreurs17 + erreur17 + "\n" 
                            

                    else:
                        try :
                            Typevaleur = (titi.getparent()).attrib['Typevaleur']
                        except:
                            erreur26 = type_regle +" :- absence de l'attribut 'Typevaleur' dans la balise " + indication 
                            erreurs26 = erreurs26 + erreur26 + "\n" 
                        try : 
                            TypeGrandeur = (titi.getparent()).attrib['TypeGrandeur']
                        except:
                            erreur26 = type_regle +" :- absence de l'attribut 'TypeGrandeur' dans la balise " + indication  
                            erreurs26 = erreurs26 + erreur26 + "\n" 
                        try :    
                            Nom_Grandeur = (titi.getparent()).attrib['Nom']
                        except:
                            erreur26 = type_regle +" :- absence de l'attribut 'Nom' dans la balise " + indication
                            erreurs26 = erreurs26 + erreur26 + "\n" 
                        try :    
                            SetType = (titi.getparent()).attrib['SetType']
                        except:
                            erreur26 = type_regle +" :- absence de l'attribut 'SetType' dans la balise " + indication   
                            erreurs26 = erreurs26 + erreur26 + "\n" 
                        
                        if TypeGrandeur not in liste_TypeGrandeur : #TypeGrandeur doit avoir une valeur dans : {"1","2","3","4"}
                                    erreur26 = type_regle +" :- la valeur du 'TypeGrandeur' dans la balise " + indication + " est erroné : '" + TypeGrandeur + "' n'existe pas dans la liste des valeurs normées " + str(liste_TypeGrandeur)
                                    erreurs26 = erreurs26 + erreur26 + "\n"
                                    
                        if TypeGrandeur == "1": #NN, Nom Normé
                            if Nom_Grandeur not in NN_Grandeur_liste:
                                    erreur26 = type_regle +" :- le Nom de la balise  "+ indication +" n'existe pas dans la liste des NN : '" + Nom_Grandeur + "'"
                                    erreurs26 = erreurs26 + erreur26 + "\n"
                                    
                        if TypeGrandeur == "3": # Paramètre, il suffit juste de vérifier que le nom de la grandeur est non-vide
                            if Nom_Grandeur == "" :
                                    erreur26 = type_regle +" :- le Nom de la balise  Grandeur est vide"
                                    erreurs26 = erreurs26 + erreur26 + "\n" 

                        if Typevaleur not in liste_TypeValeur : #Typevaleur doit avoir une valeur dans :  {"1","2","3","4","5","6"}
                                    erreur23 = type_regle +" :- la valeur de l'attribut 'Typevaleur'  dans la balise " + indication + " est erroné : '"+ TypeGrandeur + "' n'existe pas dans la liste des valeurs normées "+ str(liste_TypeValeur)
                                    erreurs23 = erreurs23 + erreur23 + "\n"       
                        if Typevaleur == "1": #NN, Nom Normé
                            if value1 not in NN_Grandeur_liste: #NN_Valeurs_liste:
                                    erreur23 = type_regle +" :- la valeur de la balise " +indication +" n'existe pas dans la liste des NN : '" + value + "'"
                                    erreurs23 = erreurs23 + erreur23 + "\n"                            
                                                    
                        if Typevaleur == "2": #Numérique

                            if value1 <> valeur_non_souhaite:
                                if functions.isInt(value1):
                                    val = int(value1)
                                    attrib_min = str((first_sheet.cell(u,5)).value)
                                    attrib_max = str((first_sheet.cell(u,6)).value)
                                    Min = (titi.getparent()).attrib[attrib_min]
                                    Max = (titi.getparent()).attrib[attrib_max]
                                    
                                    if Min <> "" and Max <> "" and Min.isspace() == False and Max.isspace() == False: # limites rensignés dans le tableau
                                            Min3 =  int(Decimal (Min, 'utf-8'))
                                            Max3 =  int(Decimal (Max, 'utf-8')) 
                                            if val < Min3 or val > Max3:  
                                                if "Matrice" in (((titi.getparent()).getparent()).getparent()).attrib:
                                                    if "PLEX" in str((((titi.getparent()).getparent()).getparent()).attrib['Matrice']):
                                                        erreur5 = type_regle +" :- Commentaire - la  valeur de la balise " +indication +" n'est pas comprise entre les bornes '"+ value +"' : Matrice(cas particulier) "
                                                        erreurs5 = erreurs5 + erreur5 + "\n" 
                                                else:         
                                                    erreur22 = type_regle +" :- la valeur de la balise " +indication +" n'est pas comprise entre les bornes min et max : '" + value + "'"
                                                    erreurs22 = erreurs22 + erreur22 + "\n"                                    
                                elif functions.isFloat(value1):
                                    val = float(value1)
                                    attrib_min = str((first_sheet.cell(u,5)).value)
                                    attrib_max = str((first_sheet.cell(u,6)).value)
                                    Min = (titi.getparent()).attrib[attrib_min]
                                    Max = (titi.getparent()).attrib[attrib_max]
                                    if Min <> "" and Max <> "" and Min.isspace() == False and Max.isspace() == False: # limites rensignés dans le tableau
         
                                        Min3 =  int(Decimal (Min, 'utf-8'))
                                        Max3 =  int(Decimal (Max, 'utf-8')) 
                                        if val < Min3 or val > Max3:  
                                            if "Matrice" in (((titi.getparent()).getparent()).getparent()).attrib:
                                                if "PLEX" in str((((titi.getparent()).getparent()).getparent()).attrib['Matrice']):
                                                    erreur22 = type_regle +" :- Commentaire - la  valeur de la balise " +indication +" n'est pas comprise entre les bornes '"+ value +"' : Matrice(cas particulier) "
                                                    erreurs22 = erreurs22 + erreur22 + "\n" 
                                            else:         
                                                erreur22 = type_regle +" :- la valeur de la balise " +indication +" n'est pas comprise entre les bornes min et max : '" + value + "'"
                                                erreurs22 = erreurs22 + erreur22 + "\n"                                    
                                elif str(value1) in liste_valeur and str (((((titi.getparent()).getparent()).getparent()).getparent()).attrib['Description']) in liste_categorie :
                                    pass # cas LOOP12
                                else:   
                                    erreur23 = type_regle +" :- la valeur de la balise " +indication +" n'est pas numérique : '" + value + "'"
                                    erreurs23 = erreurs23 + erreur23 + "\n"   

                        if Typevaleur == "3": #chaine de caractère 
                            if value1 <> valeur_non_souhaite:
                                if not re.match("^[a-zA-Z0-9_]*$", value1) and value1.isalnum()==False and not any(c.isalpha() for c in value1):
                                    erreur23 = type_regle +" :- la valeur de la balise " + indication +" n'est pas une chaine de caractère : '" + value + "'"
                                    erreurs23 = erreurs23 + erreur23 + "\n"  

                        if SetType not in liste_SetType: #Settype doit avoir une valeur dans : {"1","2","3","4","5","6","7"}
                                    erreur27 = type_regle +" :- la valeur de l'attribut 'SetType' de la balise " + indication +"  est erroné : '" + SetType + "' n'existe pas dans la liste des valeurs normées " + str(liste_SetType)
                                    erreurs27 = erreurs27 + erreur27 + "\n"
                                    
                               
            #-----------------------------------------------------------------------------------------------------------------------------------------------------

            if type_regle == "type_10":
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value) 
                nom_attribut = str((first_sheet.cell(u,5)).value) 
                val_attribut_souhaite = str((first_sheet.cell(u,6)).value) 
                for toto in tree.xpath(elem): 
                    if 'Categorie' in str(toto) and 'Point' in str(toto.getparent()):
                        if 'nom' in ((toto.getparent()).getparent()).attrib:
                            indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                        else:
                            indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                             
                    if 'Categorie' in str(toto) and 'Campagne' in str(toto.getparent()):
                        if 'nom' in (toto.getparent()).attrib : 
                            indication = "/Campagne["+(toto.getparent()).attrib['nom']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                        else:
                            indication = "/Campagne/Categorie["+toto.attrib['Description']+"]"
                                
                    indication = indication.encode(encoding='UTF-8',errors='strict') 
                       
                    if toto.attrib[nom_attribut]==val_attribut_souhaite and balise_2 in str(toto.getchildren()):
                        erreur18 = type_regle +" :- la balise " + indication + " contient la balise '"+  balise_2 + "'"
                        erreurs18 = erreurs18 + erreur18 + "\n" 
            #-----------------------------------------------------------------------------------------------------------------------------------------------------

            if type_regle == "type_11":
                balise_1 = str((first_sheet.cell(u,2)).value)
                balise_2 = str((first_sheet.cell(u,3)).value) 
                nom_attribut_1 = str((first_sheet.cell(u,4)).value) 
                val_attribut_souhaite_1 = str((first_sheet.cell(u,5)).value) 
                nom_attribut_2 = str((first_sheet.cell(u,6)).value) 
                val_attribut_souhaite_2 = str((first_sheet.cell(u,7)).value) 
                for toto in tree.xpath(elem): 
                    if 'Categorie' in str(toto) and 'Point' in str(toto.getparent()):
                        if 'nom' in ((toto.getparent()).getparent()).attrib:
                            indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                        else:
                            indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                            
                    if 'Categorie' in str(toto) and 'Campagne' in str(toto.getparent()):
                        if 'nom' in (toto.getparent()).attrib:
                            indication = "/Campagne["+(toto.getparent()).attrib['nom']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                        else:
                            indication = "/Campagne/Categorie["+toto.attrib['Description']+"]"
                    indication = indication.encode(encoding='UTF-8',errors='strict')
                        
                    if toto.attrib[nom_attribut_1]==val_attribut_souhaite_1 and balise_2 in str(toto.getchildren()):
                        if not any(titi.attrib[nom_attribut_2]==val_attribut_souhaite_2 for titi in toto):
                            erreur24 = type_regle +" :- la balise " + indication + "ne contient pas la balise " + balise_2 + "["+  nom_attribut_2+ "='"+val_attribut_souhaite_2+"']" 
                            erreurs24 = erreurs24 + erreur24 + "\n" 
                    
                if  val_attribut_souhaite_2 =="TYPREGUL": 
                    last_index = len(TYPREGUL_values) - 1 - TYPREGUL_values[::-1].index(7) 
                    nom = "TYPREGUL" 
                    for tata in tree.xpath(elem+"/Grandeur/Valeur"): 
                        if ((tata.getparent()).getparent()).attrib[nom_attribut_1]==val_attribut_souhaite_1:
                            toto = (tata.getparent()).getparent()
                            if 'Categorie' in str(toto) and 'Point' in str(toto.getparent()):
                                if 'nom' in ((toto.getparent()).getparent()).attrib:
                                    indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                                else:
                                    indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                                    
                            if 'Categorie' in str(toto) and 'Campagne' in str(toto.getparent()):
                                if 'nom' in (toto.getparent()).attrib:
                                    indication = "/Campagne["+(toto.getparent()).attrib['nom']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                                else:
                                    indication = "/Campagne/Categorie["+toto.attrib['Description']+"]"
                            indication = indication.encode(encoding='UTF-8',errors='strict') 
                               
                            if nom == (tata.getparent()).attrib['Nom'] :
                                for i in range(0,len(TYPREGUL_values)-1):
                                    if (int(tata.text) ==TYPREGUL_values[i] and int(tata.text) <> 7) :
                                        if not any(couple_values_1[i]  in xoxo.attrib['Nom']  for xoxo in (((tata.getparent()).getparent()).getchildren())) or not any(couple_values_2[i] in xoxo.attrib['Nom']  for xoxo in (((tata.getparent()).getparent()).getchildren())):
                                            erreur24 = type_regle +" :- la balise " + indication + " doit contenir les Grandeurs "+ couple_values_1[i] +" et " + couple_values_2[i] + " car elle contient la Grandeur[Nom='" + nom + "']/Valeur[CDATA='"+ tata.text +"']"
                                            erreurs24 = erreurs24 + erreur24  +"\n"
                                            
                                                  
                                    if (int(tata.text) ==TYPREGUL_values[i]==7) :
                                        if (not any(couple_values_1[i] in xoxo.attrib['Nom']  for xoxo in (((tata.getparent()).getparent()).getchildren())) or not any(couple_values_2[i] in xoxo.attrib['Nom']  for xoxo in (((tata.getparent()).getparent()).getchildren()))) :
                                            liste_OR.append(0)
                                            liste_OR_couple.append(couple_values_1[i]+"/"+couple_values_2[i])
                                        else:
                                            liste_OR.append(1) 
                                        if 1 not in liste_OR and i == last_index:                       
                                            erreur24 = type_regle +" :- la balise " + indication + " doit contenir un de ces couples de Grandeurs "+ str(liste_OR_couple)  + " car elle contient la Grandeur[Nom='" + nom + "']/Valeur[CDATA='"+ tata.text +"']"
                                            erreurs24 = erreurs24 + erreur24 + "\n"  
                                             
                                                           
            #-----------------------------------------------------------------------------------------------------------------------------------------------------
            
            if type_regle == "type_12":
                balise_1 = str((first_sheet.cell(u,2)).value)
                nom_attribut_1 = str((first_sheet.cell(u,3)).value) 
                val_attribut_souhaite_1 = (first_sheet.cell(u,4)).value
 
                balise_2 = str((first_sheet.cell(u,5)).value) 
                nom_attribut_2 = str((first_sheet.cell(u,6)).value) 
                val_attribut_souhaite_2 = str((first_sheet.cell(u,7)).value) 
                val_attribut_souhaite_3 = str((first_sheet.cell(u,8)).value)          
                
                for toto in tree.xpath(elem): 
                    
                    for i in toto.getchildren():
                        toto = i.getparent()
                        if 'Categorie' in str(toto) and 'Point' in str(toto.getparent()):
                            if 'nom' in ((toto.getparent()).getparent()).attrib:
                                indication = "/Campagne["+((toto.getparent()).getparent()).attrib['nom']+"]" + "/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                            else:  
                                indication = "/Campagne/Point["+(toto.getparent()).attrib['Type']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                                  
                        if 'Categorie' in str(toto) and 'Campagne' in str(toto.getparent()):
                            if 'nom' in (toto.getparent()).attrib:
                                indication = "/Campagne["+(toto.getparent()).attrib['nom']+"]" + "/Categorie["+toto.attrib['Description']+"]"
                            else:    
                                indication = "/Campagne/Categorie["+toto.attrib['Description']+"]"
                        indication = indication.encode(encoding='UTF-8',errors='strict') 
                    
                        if toto.attrib[nom_attribut_1]==val_attribut_souhaite_1  and (i.attrib[nom_attribut_2] ==  val_attribut_souhaite_2 or i.attrib[nom_attribut_2]== val_attribut_souhaite_3 ):
                            
                            if (i.attrib[nom_attribut_2]== val_attribut_souhaite_2 and not any(titi.attrib[nom_attribut_2]== val_attribut_souhaite_3 for titi in toto))   or (i.attrib[nom_attribut_2]== val_attribut_souhaite_3 and not any(titi.attrib[nom_attribut_2]== val_attribut_souhaite_2 for titi in toto)) :
                               
                                erreur29 = type_regle +" :- Incoherence couple 'calibration/soft' (absence de l'un de ces deux Grandeurs : EPROMACT/A2L_DESC) dans la balise " +  indication
                                erreurs29 = erreurs29 + erreur29 + "\n" 
                         
            #-----------------------------------------------------------------------------------------------------------------------------------------------------
            
            if type_regle == "type_13":
                balise_1 = str((first_sheet.cell(u,2)).value)
                for titi in tree.xpath(elem):
                      
                    NbPoints = titi.attrib['NbPts'] # voir Liste_règles_main_function_V2 ( 4bis.6 )
                                   
                    for toto in titi:
                        for tete in toto: 
                            if "Consignes" in str(tete):
                                for tutu in tete:
                                    NbValeurs =  len(tutu.getchildren())
                                    
                                    titi = ((tutu.getparent()).getparent()).getparent()
                                    if 'Variation' in str(titi) and 'Niveau' in str(titi.getparent()):
                                        if 'nom' in ((((titi.getparent()).getparent()).getparent()).getparent()).attrib:
                                            indication = "/Campagne["+((((titi.getparent()).getparent()).getparent()).getparent()).attrib['nom']+"]" + "/Point["+(((titi.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"
                                        else:
                                            indication = "/Campagne/Point["+(((titi.getparent()).getparent()).getparent()).attrib['Type']+"]" + "/Variations/Niveau/Variation"
                                            
                                    indication = indication.encode(encoding='UTF-8',errors='strict')                                     
                                    
                                    if int(NbPoints)<> NbValeurs:
                                        erreur21 = type_regle +" :- la balise "  + indication +" a un nombre de valeurs different du nombre de points NbPts" 
                                        erreurs21 = erreurs21 + erreur21 + "\n" 

            #-----------------------------------------------------------------------------------------------------------------------------------------------------            

            if type_regle == "banc_alti": 
                Categorie_banc_alti_flag_1 = True
                Categorie_banc_alti_flag_2 = True
                absence_grandeur_SIMULATI = True
                donne_vaut_1 = True
                absence_grandeur_C_PALTI = True
                donne_hors_bornes = True
                for toto in tree.xpath(elem):
                    try:
                        if str(toto.attrib['banc']) == "C46" or str(toto.attrib['banc']) == "C45" :
                            for tata in toto:
                                for titi in tata:
                                    if "Categorie" in str(titi):
                                        try:
                                            if any (titi.attrib['Description'] == "Instruments"   for titi in tata):
                                                if titi.attrib['Description'] == "Instruments":
                                                    if "Grandeur" in  str(titi.getchildren()):
                                                        for tutu in titi:
                                                            if any ( tutu.attrib['Nom'] == "SIMUALTI"  for tutu in titi):
                                                                for tyty in tutu: 
                                                                    if "ValeurPredefinie"  not in str(tyty):
                                                                        if tyty.text <> '1' and (tyty.getparent()).attrib['Nom'] == "SIMUALTI":
                                                                            donne_vaut_1 = False
                                                            else : 
                                                                absence_grandeur_SIMULATI = False                
                                        except:
                                            Categorie_banc_alti_flag_1 = False                          
                                        try:
                                            if any (titi.attrib['Description'] == "Consigne Altimetrie"  for titi in tata):
                                                if titi.attrib['Description'] == "Consigne Altimetrie":
                                                    if "Grandeur" in  str(titi.getchildren()):
                                                        for tutu in titi:
                                                            if any( tutu.attrib['Nom'] == "C_PALTI" for tutu in titi):
                                                                for tete in tutu: 
                                                                    if "ValeurPredefinie"  not in str(tete):
                                                                        if (int(tete.text) < 540 or  int(tete.text) > 1020) and (tete.getparent()).attrib['Nom'] == "C_PALTI":
                                                                            donne_hors_bornes = False
                                                            else : 
                                                                absence_grandeur_C_PALTI = False          
                                        except:
                                            Categorie_banc_alti_flag_2 = False  
                                                                                                                     
                                if  Categorie_banc_alti_flag_1== False : 
                                    erreur30 = type_regle +" :- absence de la Categorie[Instrument]" 
                                    erreurs30 = erreurs30 + erreur30 + "\n"                                 
                                elif Categorie_banc_alti_flag_1 == True:
                                    if absence_grandeur_SIMULATI == False :  
                                        erreur30 = type_regle +" :- absence de la  Grandeur[SIMUALTI] dans la Categorie[Instrument]" 
                                        erreurs30 = erreurs30 + erreur30 + "\n"    
                                    elif absence_grandeur_SIMULATI == True:                                     
                                        if donne_vaut_1 == False : 
                                            erreur30 = type_regle +" :- la Valeur de la Grandeur[SIMUALTI] dans la Categorie[Instrument] ne vaut pas 1 " 
                                            erreurs30 = erreurs30 + erreur30 + "\n" 
                                    
                                if  Categorie_banc_alti_flag_2== False:
                                    erreur30 = type_regle +" :- absence de la Categorie[Consigne Altimetrie]" 
                                    erreurs30 = erreurs30 + erreur30 + "\n"   
                                elif Categorie_banc_alti_flag_2== True:                                                                    
                                    if absence_grandeur_C_PALTI == False : 
                                        erreur30 = type_regle +" :- absence de la  Grandeur[C_PALTI] dans la Categorie[Consigne Altimetrie]" 
                                        erreurs30 = erreurs30 + erreur30 + "\n"  
                                    elif absence_grandeur_C_PALTI == True :                                        
                                        if donne_hors_bornes == False :   
                                            erreur30 = type_regle +" :- la Valeur de la  Grandeur[C_PALTI] dans la Categorie[Consigne Altimetrie] nest pas comprise entre les bornes 540-1020 " 
                                            erreurs30 = erreurs30 + erreur30 + "\n"  
                    except:
                        erreur30 = " !! ce fichier ne contient aucune Campagne !! " 
                        erreurs30 = erreurs30 + erreur30 + "\n"                        
            #-----------------------------------------------------------------------------------------------------------------------------------------------------
                           
            if type_regle == "banc_clim": 
                Categorie_Regulations_flag = True
                Categorie_Instruments_flag = True
                absence_grandeur_QREGHUCT = True
                QREGHUCT_vaut_1_ou_0 = True
                absence_grandeur_VREGHUCT = True
                VREGHUCT_hors_bornes = True
                absence_grandeur_C_THUCTR = True
                absence_grandeur_CTCCL  = True
                absence_grandeur_TAIRCOMB = True
                absence_grandeur_HAIRCOMB   = True                
                C_THUCTR_hors_bornes = True
                C_THU_hors_bornes = False
                HAIRCOMB__vaut_0  = True
                CTCCL_inf_10 = False
                TAIRCOMB_inf_10 = False
                Regulations = unicode("Régulations", 'utf-8')   
                             
                for toto in tree.xpath(elem):
                    try:
                        if str(toto.attrib['banc']) == "C04" or str(toto.attrib['banc']) == "C55" :
                            for tata in toto:
                                for titi in tata:
                                    if "Categorie" in str(titi):
                                        try:
                                            if any (titi.attrib['Description'] == Regulations  for titi in tata): 
                                                if titi.attrib['Description'] == Regulations:
                                                    if "Grandeur" in  str(titi.getchildren()):
                                                        for tutu in titi:
                                                            if any( tutu.attrib['Nom'] == "C_THU" for tutu in titi):
                                                                for tete in tutu: 
                                                                    if "ValeurPredefinie"  not in str(tete):
                                                                        if (tete.text).isnumeric() == False:
                                                                            C_THU_hors_bornes = False
                                                                        elif (int(tete.text) <= 40 ) and (tete.getparent()).attrib['Nom'] == "C_THU":
                                                                            C_THU_hors_bornes = True
                                                              
                                                            if C_THU_hors_bornes == True:
                                                                if any( tutu.attrib['Nom'] == "C_THUCTR" for tutu in titi):
                                                                    for tete in tutu: 
                                                                        if "ValeurPredefinie"  not in str(tete):
                                                                            if (int(tete.text) < -31 or  int(tete.text) > 50) and (tete.getparent()).attrib['Nom'] == "C_THUCTR" :
                                                                                C_THUCTR_hors_bornes = False
                                                                else : 
                                                                    absence_grandeur_C_THUCTR = False     
                                                                 
                                                            if any( tutu.attrib['Nom'] == "CTCCL" for tutu in titi):
                                                                for tete in tutu: 
                                                                    if "ValeurPredefinie"  not in str(tete):
                                                                        if (int(tete.text) < 10) and (tete.getparent()).attrib['Nom'] == "CTCCL":
                                                                            CTCCL_inf_10 = True
                                                            else : 
                                                                absence_grandeur_CTCCL = False     
                                                            if any( tutu.attrib['Nom'] == "TAIRCOMB" for tutu in titi):
                                                                for tete in tutu: 
                                                                    if "ValeurPredefinie"  not in str(tete):
                                                                        if (int(tete.text) < 10) and (tete.getparent()).attrib['Nom'] == "TAIRCOMB":
                                                                            TAIRCOMB_inf_10 = True
                                                            else : 
                                                                absence_grandeur_TAIRCOMB = False  
                                                                              
                                                            if any( tutu.attrib['Nom'] == "HAIRCOMB" for tutu in titi):
                                                                for tete in tutu: 
                                                                    if "ValeurPredefinie"  not in str(tete):
                                                                        if (int(tete.text) <> 0) and ((tete.getparent()).attrib['Nom'] == "HAIRCOMB") and(CTCCL_inf_10 == True or TAIRCOMB_inf_10 == True):
                                                                            HAIRCOMB__vaut_0= False
                                                            else : 
                                                                absence_grandeur_HAIRCOMB = False   
                                                                                                                                                                               
                                        except:
                                            Categorie_Regulations_flag = False 
                                                                                
                                        try:
                                            if any (titi.attrib['Description'] == "Instruments"   for titi in tata):
                                                if titi.attrib['Description'] == "Instruments":
                                                    if "Grandeur" in  str(titi.getchildren()):
                                                        for tutu in titi:
                                                            if C_THU_hors_bornes == True:
                                                                if any ( tutu.attrib['Nom'] == "QREGHUCT"  for tutu in titi):
                                                                    for tyty in tutu: 
                                                                        if "ValeurPredefinie"  not in str(tyty):
                                                                            if (tyty.text <> '1' or tyty.text <> '0') and (tyty.getparent()).attrib['Nom'] == "QREGHUCT":
                                                                                QREGHUCT_vaut_1_ou_0 = False
                                                                else : 
                                                                    absence_grandeur_QREGHUCT = False         
                                                                        
                                                                if any ( tutu.attrib['Nom'] == "VREGHUCT"  for tutu in titi):
                                                                    for tyty in tutu: 
                                                                        if "ValeurPredefinie"  not in str(tyty):
                                                                            if (int(tyty.text) < -40 or  int(tyty.text) > 50) and (tyty.getparent()).attrib['Nom'] == "VREGHUCT":
                                                                                VREGHUCT_hors_bornes = False
                                                                else : 
                                                                    absence_grandeur_VREGHUCT = False                                                                  
                                        except:
                                            Categorie_Instruments_flag = False                                        
    
                                if  Categorie_Instruments_flag== False : 
                                    erreur30 = type_regle +" :- absence de la Categorie[Instrument]" 
                                    erreurs30 = erreurs30 + erreur30 + "\n" 
                                         
                                elif   Categorie_Instruments_flag== True : 
                                    if C_THU_hors_bornes == True:                                                            
                                        if absence_grandeur_QREGHUCT == False :  
                                            erreur30 = type_regle +" :- absence de la  Grandeur[QREGHUCT]  dans la Categorie[Instrument]"  
                                            erreurs30 = erreurs30 + erreur30 + "\n"  
                                        elif absence_grandeur_QREGHUCT == True :                                         
                                            if QREGHUCT_vaut_1_ou_0 == False : 
                                                erreur30 = type_regle +" :- la Valeur de la Grandeur[QREGHUCT]  dans la Categorie[Instrument] ne vaut pas 1 ou 0  " 
                                                erreurs30 = erreurs30 + erreur30 + "\n"   
                                        if absence_grandeur_VREGHUCT == False : 
                                            erreur30 = type_regle +" :- absence de la Grandeur[VREGHUCT]  dans la la Categorie[Instrument] " 
                                            erreurs30 = erreurs30 + erreur30 + "\n" 
                                        elif  absence_grandeur_VREGHUCT == True :                                         
                                            if VREGHUCT_hors_bornes == False :   
                                                erreur30 = type_regle +" :- la Valeur de la Grandeur[VREGHUCT]  dans la la Categorie[Instrument] n'est pas comprise entre les bornes [-40 ,50] " 
                                                erreurs30 = erreurs30 + erreur30 + "\n" 
    
                                if  Categorie_Regulations_flag== False : 
                                    erreur30 = type_regle +" :- absence de la Categorie[Régulations]"   
                                    erreurs30 = erreurs30 + erreur30 + "\n"  
                                elif Categorie_Regulations_flag== True :   
                                    if C_THU_hors_bornes == True:
                                        if absence_grandeur_C_THUCTR == False : 
                                            erreur30 = type_regle +" :- absence de la  Grandeur[C_THUCTR]  dans la Categorie[Régulations]" 
                                            erreurs30 = erreurs30 + erreur30 + "\n" 
                                        elif absence_grandeur_C_THUCTR == True :                                          
                                            if C_THUCTR_hors_bornes == False :  
                                                erreur30 = type_regle +" :- la Valeur de la Grandeur[C_THUCTR]  dans la Categorie[Régulations] n'est pas comprise entre les bornes [-31 ,50] " 
                                                erreurs30 = erreurs30 + erreur30 + "\n" 
                                                                             
                                    if absence_grandeur_CTCCL == False : 
                                        erreur30 = type_regle +" :- absence de  la Grandeur[CTCCL]  dans la Categorie[Régulations] " 
                                        erreurs30 = erreurs30 + erreur30 + "\n" 
                                    elif absence_grandeur_CTCCL == True :          
                                        if CTCCL_inf_10 == False :  
                                            erreur30 = type_regle +" :- la Valeur de la Grandeur[CTCCL]  dans la Categorie[Régulations] est > 10" 
                                            erreurs30 = erreurs30 + erreur30 + "\n"
                                            
                                    if absence_grandeur_TAIRCOMB == False : 
                                        erreur30 = type_regle +" :- absence de la Grandeur[TAIRCOMB] dans la Categorie[Régulations] " 
                                        erreurs30 = erreurs30 + erreur30 + "\n"  
                                    elif absence_grandeur_TAIRCOMB == True :          
                                        if TAIRCOMB_inf_10 == False :  
                                            
                                            erreur30 = type_regle +" :- la Valeur de la Grandeur[TAIRCOMB] dans la Categorie[Régulations]  est > 10" 
                                            erreurs30 = erreurs30 + erreur30 + "\n"
                                            
                                    if absence_grandeur_HAIRCOMB == False : 
                                        erreur30 = type_regle +" :- absence de la Grandeur[HAIRCOMB] dans la Categorie[Régulations] " 
                                        erreurs30 = erreurs30 + erreur30 + "\n"                                 
                                        absence_grandeur_CTCCL 
                                    elif absence_grandeur_HAIRCOMB == True :          
                                        if HAIRCOMB__vaut_0 == False :  
                                            erreur30 = type_regle +" :- la Valeur de la Grandeur[HAIRCOMB] dans la Categorie[Régulations] ne vaut pas 0" 
                                            erreurs30 = erreurs30 + erreur30 + "\n"     
                    except:
                        erreur30 = " !! ce fichier  ne contient aucune Campagne !! " 
                        erreurs30 = erreurs30 + erreur30 + "\n"                           
            #-----------------------------------------------------------------------------------------------------------------------------------------------------    

        #**********************************************************************************************
        #            affichage des erreurs 
        #********************************************************************************************** 

        erreurs = erreurs1 + erreurs2 + erreurs3 + erreurs4 + erreurs6 + erreurs7 + erreurs8 + erreurs9 + erreurs10 + erreurs11 + erreurs12 + erreurs13 + erreurs14 + erreurs15 + erreurs16 + erreurs17 + erreurs18 + erreurs19 + erreurs20 + erreurs21 + erreurs22+ erreurs23 + erreurs24 + erreurs25 + erreurs26 + erreurs27 + erreurs28+ erreurs29+ erreurs30       
#         text.insert('1.0',erreurs5)
        
        if erreurs5 <> '' and erreurs.isspace() == False:
            text_file_1.write(erreurs5 + "\r")
        
        if erreurs <> '' and erreurs.isspace() == False:
            text_file.write(erreurs + "\r")
        
                    
        xscrollbar.config(command=text.xview)
        yscrollbar.config(command=text.yview)
        text.pack() 
    text_file.close()
    text_file_1.close()
    tfile = open("Output.txt", "r")
    tfile_1 = open("Output_1.txt", "r")
    lines = tfile.readlines()
    lines_1 = tfile_1.readlines()
    Nbr_lignes = len(lines)
    Nbr_lignes_1 = len(lines_1)
    erreurs =''
    erreurs5=''
    for l in  range(0,Nbr_lignes):
        if str(lines[l]) <> '' and str(lines[l]).isspace()==False:
            erreurs = erreurs + str(lines[l]) 

    for l in  range(0,Nbr_lignes_1):
        if str(lines_1[l]) <> ''and str(lines_1[l]).isspace()==False:
            erreurs5 = erreurs5 + str(lines_1[l]) 

    text.tag_add("here", "1.0", "end")
    text.insert( "1.0" ,  erreurs5)
    if erreurs <> '':
        text.tag_config("here", background="white", foreground="black")
    else:    
        text.tag_config("here", background="white", foreground="blue")

    text.tag_add("here",  "1.0", "end")
    text.insert( "1.0" ,  erreurs)
    text.tag_config("here", background="white", foreground="blue")
    
    tfile.close()
    tfile_1.close()  
    os.remove('Output.txt')
    os.remove('Output_1.txt')
    text.pack() 



    #**********************************************************************************************
    #            fichier csv : ajouter le résultat du test dans le fichier 'statistiques.csv'
    #**********************************************************************************************    
    currentDT = datetime.datetime.now()
    time =  (currentDT.strftime("%Y-%m-%d %H:%M:%S")) + " XML File : " + xml_location    
    in_file  = open('statistiques.csv', 'r')
    reader = csv.reader(in_file)
    
    out_file = open('temp.csv', 'wb') 
    Writer = csv.writer(out_file, quoting=csv.QUOTE_NONE, delimiter=',', quotechar='',escapechar=',')
    
    contenu =  text.get(1.0, "end-1c") 
    contenu = contenu.encode(encoding='UTF-8',errors='strict')
    time = time.encode(encoding='UTF-8',errors='strict')

    Writer.writerow([time]) 
    if contenu<>'':
        Writer.writerow([contenu])
    for row in reader:        
        Writer.writerow(row) 
    in_file.close()
    out_file.close()    
    # récupérer le contenu du fichier temp.csv(ancien + nouveau contenu) et l'ajouter dans le fichier statistiques
    in_file  = open('temp.csv', 'rU')
    reader = csv.reader(in_file)
    out_file = open('statistiques.csv', 'wb') 
    Writer = csv.writer(out_file, quoting=csv.QUOTE_NONE, delimiter=',', quotechar='',escapechar=',')
    for row in reader:        
        Writer.writerow(row) 
    in_file.close()
    out_file.close()  
    os.remove('temp.csv')
    
    # empêcher l'utilisateur d'écrire dans la fenêtresdes erreurs
    

    
    a = text.get(1.0, "end-1c") 
    if a=='' or a.isspace():
        text.insert( "1.0" ,  '                          Test sans erreur')
        
    text.config(state= DISABLED)
    text.pack() 
    
   
    winE.destroy()

# ******************************************************** End on_button() ************************************************

#**********************************************************************************************
#   boite de dialogue pour ouvrir le fichier de config via 'Mot de passe'
#**********************************************************************************************
class MyDialog:
    def __init__(self, parent):

        top = self.top = Toplevel(parent)
        top.wm_iconbitmap('PSA_logo.ico')     
        top.geometry("300x100")               
        Label(top, text="Veuillez Entrer le mot de passe").pack(pady=10)

        self.e = Entry(top, show="*")
        self.e.pack(padx=5)

        b = Button(top, text="OK", command=self.ok)
        b.pack(pady=5)

    def ok(self):
        if self.e.get() == "EMAN":
            webbrowser.open("config.xlsx")
            tkMessageBox.showinfo(title="Alerte", message="N'oubliez pas de sauvegarder le fichier avant de lancer le Test !") 
            self.top.destroy()
        else:
            tkMessageBox.showinfo(title="Alerte", message="Mot de passe erroné !")   
            d = MyDialog(root)
            root.wait_window(d.top) 
def Open_Config (): 
    d = MyDialog(root)
    root.wait_window(d.top)  
    
  


 
 
#**********************************************************************************************
#                         les boutons de l'interface 
#**********************************************************************************************
button_font = tkFont.Font(family='Helvetica', size=14, weight='bold')  #style de l'écriture

button_parcourir = Button(frame1, text='Parcourir', command=set_filename,bg = "grey",height =1, width = 10)
button_parcourir['font'] = button_font
button_parcourir.place(x=0, y=0)
button_parcourir.pack(padx=100,pady=5)   

button_tester = Button(frame3, text="Tester XML", command=on_button,bg = "skyblue",height =1, width = 10)
button_tester['font'] = button_font
button_tester.pack(side=LEFT)

button_quitter = Button(frame3, text="Quitter", command=close_window,bg = "tomato",height =1, width = 10)
button_quitter['font'] = button_font
button_quitter.pack(side=RIGHT, padx= 5)

button_apropos = Button(frame5, text="A propos", command=Open_Apropos,bg = "khaki",height =1, width = 10)
button_apropos['font'] = button_font
button_apropos.pack(side=LEFT)

button_config = Button(frame5, text="Config", command=Open_Config,bg = "goldenrod",height =1, width = 10)
button_config['font'] = button_font
button_config.pack(side=RIGHT, padx= 5)

#**********************************************************************************************
#                         fonction main
#**********************************************************************************************
root.mainloop()

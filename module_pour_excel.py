#from xlutils.copy import copy 

import openpyxl 
from openpyxl.styles import PatternFill
from copy import copy
from datetime import date, datetime
from utils import UtilsForFile, UtilsForSheet, Str

class Path():
    def __init__(self,path = 'fichiers_xls/'):
        self.path = path
        
    def act_on_files(self,fonction):
        """
        Fonction qui prend tous les fichiers d'un dossier et qui applique une même action à ces fichiers.
        """
        pass 


class File(UtilsForFile): 
    def __init__(self, name_file, path = 'fichiers_xls/'):
        """L'utilisateur sera invité à mettre son fichier xslx dans un dossier nommé fichiers_xls
        """
        self.name_file = name_file  
        self.path = path
        self.writebook = openpyxl.load_workbook(self.path + self.name_file, data_only=True)
        self.sheets_name = self.writebook.sheetnames

    def sauvegarde(self):
        """
        Fonction qui crée une sauvegarde du fichier name_file et 
        qui l'appelle name_file_time où time est le moment d'enregistrement.

        Exemple d'utilisation : 
            file = File('dataset.xlsx')
            file.sauvegarde()
        """
        file_copy = openpyxl.Workbook()
        del file_copy[file_copy.active.title] #supprimer l'onglet créé

        for onglet in self.sheets_name:
            new_sheet = file_copy.create_sheet(onglet)
            initial_sheet = self.writebook[onglet] 

            for i in range(1,initial_sheet.max_row+1):
                for j in range(1,initial_sheet.max_column+1): 
                    new_sheet.cell(i,j).value = initial_sheet.cell(i,j).value  
                    new_sheet.cell(i,j).fill = copy(initial_sheet.cell(i,j).fill)
                    new_sheet.cell(i,j).font = copy(initial_sheet.cell(i,j).font) 
                    
        name_file_no_extension = Str(self.name_file).del_extension() 
        
        file_copy.save(self.path  + name_file_no_extension + '_date_' + datetime.now().strftime("%Y-%m-%d_%Hh%M") + '.xlsx') 

            

    def create_one_onglet_by_participant(self, onglet_from, column_read, first_line=2):
        """
        Fonction qui prend un onglet dont une colonne contient des chaînes de caractères.
        Chaque chaîne de caractères peut apparaître plusieurs fois dans cette colonne. 
        La fonction retourne un fichier contenant un onglet par chaîne de caractères.
          Chaque onglet contient toutes les lignes correspondant à cette chaîne de caractères.

        Input : 
            onglet_from : onglet de référence.
            column_read : la colonne qui contient les chaînes de caractères.
            first_line : ligne où commencer à parcourir.
            last_line : ligne de fin de parcours
 
        Exemple d'utilisation : 
    
            file = File('dataset.xlsx')
            file.create_one_onglet_by_participant('onglet1', 1) 
        """

        onglets = []
        sheet = self.writebook[onglet_from] 

        for i in range(first_line, sheet.max_row + 1):
            onglet = str(sheet.cell(i,column_read).value)
            if onglet not in onglets:
                self.writebook.create_sheet(onglet)
                self.copy_paste_line(sheet, 1,  self.writebook[onglet], 1)
                onglets.append(onglet) 
            self.add_line_at_bottom(sheet, i, self.writebook[onglet])
        
        self.writebook.save(self.path + self.name_file) 

    def extract_column_from_all_sheets(self,column):
        """
        Fonction qui récupère une colonne dans chaque onglet pour former une nouvelle feuille
        contenant toutes les colonnes. La première cellule de chaque colonne correspond alors 
        au nom de l'onglet. Attention, en l'état, il faut que tous les onglets aient la même structure.

        Inputs : 
            column : int. Le numéro de la colonne à récupérer 
        """ 
        new_sheet = self.writebook.create_sheet(f"gather_{column}")
        num_col = 1
        for name_onglet in self.sheets_name:
            onglet = self.writebook[name_onglet] 
            self.copy_paste_column(onglet,column,new_sheet,num_col)
            num_col = new_sheet.max_column + 1
            new_sheet.cell(1,new_sheet.max_column).value = name_onglet 
            
        self.writebook.save(self.path + self.name_file) 



class Sheet(File,UtilsForSheet): 
    def __init__(self, name_file, name_onglet,path = 'fichiers_xls/'): 
        super().__init__(name_file,path)
        self.name_onglet = name_onglet  
        self.sheet = self.writebook[self.name_onglet]
        del self.sheets_name

    def column_transform_string_in_binary(self,column_read,column_write,*good_answers,line_beginning = 2, line_end = 100, insert = True, security = True):
        """
        Fonction qui prend une colonne de chaîne de caractères et qui renvoie une colonne de 0 ou de 1
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où mettre les 0 ou 1.

        Inputs :
                column_read : la colonne de lecture des réponses.
                colum_write : la colonne d'écriture des 0 et 1. Par défaut, une colonne est insérée à cette position.
                good_answers : une séquence d'un nombre quelconque de bonnes réponses qui valent 1pt. Chaque mot ne doit pas contenir d'espace ni au début ni à la fin.
                line_beggining, line_end : (paramètres optionnels par défaut égaux à 2 et 100) intervalle de ligne dans lequel l'utilisateur veut appliquer sa transformation
                insert : (paramètre optionnel) le mettre à False si on ne veut pas insérer une colonne.
        
        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.

        Exemple d'utilisation : 
        
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.column_transform_string_in_binary(1,2,'reponse1','reponse2',line_end = 1600) 

            #Bien mettre les réponses de good_answers entre ''. 
        """

        if insert == False and security == True and self.column_security(column_write) == False:
            msg = "La colonne n'est pas vide. Si vous voulez vraiment y écrire, mettez security = False en argument."
            print(msg)
            return msg

        if insert == True:
            self.sheet.insert_cols(column_write)

        for i in range(line_beginning,line_end):
            chaine_object = Str(self.sheet.cell(i,column_read).value)  
            bool = chaine_object.clean_string().transform_string_in_binary(*good_answers) 
            self.sheet.cell(i,column_write).value = bool
 
        self.writebook.save(self.path + self.name_file)

    def column_convert_in_minutes(self,column_read,column_write,line_beginning = 2, line_end = 100, insert = True, security = True):
        """
        Fonction qui prend une colonne de chaines de caractères de la forme "10 jour 5 heures" 
        ou "5 heures 10 min" ou "10 min 5s" ou "5s" et qui renvoie le temps en minutes.
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne à remplir.
        Input : column_read : la colonne de lecture des réponses.
                colum_write : la colonne d'écriture. 
                line_beggining, line_end : (optionnel par défaut égaux à 2 et 100) intervalle de ligne dans lequel l'utilisateur veut appliquer sa transformation
                insert : (paramètre optionnel) le mettre à False si on ne veut pas insérer une colonne.

        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.
        
        Exemple d'utilisation : 
        
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.column_convert_in_minutes(1,2,line_beggining = 3,line_end = 1600) 

        """

        if insert == False and security == True and self.column_security(column_write) == False:
            msg = "La colonne n'est pas vide. Si vous voulez vraiment y écrire, mettez security = False en argument."
            print(msg)
            return msg

        if insert == True:
            self.sheet.insert_cols(column_write)

        for i in range(line_beginning,line_end):
            chaine_object = Str(self.sheet.cell(i,column_read).value)  
            bool = chaine_object.clean_string().convert_time_in_minutes() 
            self.sheet.cell(i,column_write).value = bool
 
        self.writebook.save(self.path + self.name_file)

    def column_set_answer_in_group(self,column_read,column_write,groups_of_responses,line_beginning = 2, line_end = 100, insert = True, security = True):
        """
        Dans le cas où il y a des groupes de réponses, cette fonction qui prend une colonne de chaîne de caractères 
        et qui renvoie une colonne remplie de chaînes contenant les groupes associés.
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où écrire.

        Input : 
                column_read : la colonne de lecture des réponses.
                colum_write : la colonne d'écriture des 0 et 1. 
                groups_of_response : dictionnary which keys are response groups and which values are a list of responses 
        associated to this group.
                line_beggining, line_end : (paramètres optionnels par défaut égaux à 2 et 100) intervalle de ligne dans lequel l'utilisateur veut appliquer sa transformation
                insert : (paramètre optionnel) le mettre à False si on ne veut pas insérer une colonne.

        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.
        
        Exemple d'utilisation : 
        
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.column_set_answer_in_group(1, 2, {'groupe1':[1,4,3],'groupe2':[5,2]} ,line_beggining = 3,line_end = 1600) 

        """

        if insert == False and security == True and self.column_security(column_write) == False:
            msg = "La colonne n'est pas vide. Si vous voulez vraiment y écrire, mettez security = False en argument."
            print(msg)
            return msg

        if insert == True:
            self.sheet.insert_cols(column_write)

        for i in range(line_beginning,line_end): 
            chaine_object = Str(self.sheet.cell(i,column_read).value)  
            group = chaine_object.clean_string().set_answer_in_group(groups_of_responses) 
            self.sheet.cell(i,column_write).value = group
            
        self.writebook.save(self.path + self.name_file)
        
    def color_special_cases_in_column(self,column,chainecolor):
        """
        Fonction qui regarde pour une colonne donnée colore les cases correspondant à certaines chaînes de caractères.

        Input : 
            - column : le numéro de la colonne.
            - chainecolor : les chaînes de caractères qui vont être colorées et les couleurs qui correspondent à écrire avec la syntaxe suivante {'vrai':'couleur1','autre':couleur2}. Attention,
                la couleur doit être entrée en hexadécimal et les chaînes de caractères ne doivent pas avoir d'espace au début ou à la fin.
        
        Exemple d'utilisation : 
        
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.color_special_cases_in_column(12, {'vrai': '#FF0000','faux': '#00FF00'}) 

        """
        
        for i in range(1,self.sheet.max_row + 1):
            cellule = self.sheet.cell(i,column) 

            if type(cellule.value) == str:
                key = Str(cellule.value).clean_string().chaine
            else: 
                key = cellule.value

            if key in chainecolor.keys():
                cellule.fill = PatternFill(fill_type = 'solid', start_color = chainecolor[key])

        self.writebook.save(self.path + self.name_file)

    def color_special_cases_in_sheet(self,chainecolor): 
        """
        Fonction qui colore les cases contenant à certaines chaînes de caractères d'une feuille
        
        Input : 
            - column : le numéro de la colonne.
            - chainecolor : les chaînes de caractères qui vont être colorées et les couleurs qui correspondent à écrire avec la syntaxe suivante {'vrai':'couleur1','autre':couleur2}. Attention,
                la couleur doit être entrée en hexadécimal et les chaînes de caractères ne doivent pas avoir d'espace au début ou à la fin.
        
        Exemple d'utilisation : 
    
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.color_special_cases_in_sheet({'vrai': '#FF0000','faux': '#00FF00'}) 
                
        """

        for j in range(1, self.sheet.max_column + 1):
            self.color_special_cases_in_column(j,chainecolor)

    def add_column_in_sheet_differently_sorted(self,column_identifiant, column_insertion, other_sheet):
        """
        Fonction qui insère dans un onglet des colonnes d'un autre onglet de référence. 
        Les deux feuilles ont une colonne d'identifiants communs (exemple : des mails) mais qui peut être
        triés dans des ordres différents.
        La fonction récupère un ou plusieurs éléments d'une ligne déterminée par un identifiant.
        Elle recherche l'identifiant dans la seconde feuille et insère les éléments
        dans la ligne correspondante.

        Inputs :
            - column_identifiant : numéro de la colonne où sont situés les identifiants dans l'onglet qu'on souhaite modifier.
            - column_insertion : numéro de la colonne où on insère les colonnes à récupérer.
            - other_sheet : liste représentant l'onglet duquel on récupère les colonnes  ['namefile','namesheet',numéro de la colonne où sont les identifiants,[numéros des colonnes à récupérer sous forme de liste]]
                namefile doit être au format .xlsx et mis dans le dossier fichier_xls.

        Exemple d'utilisation : 
    
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.add_column_in_sheet_differently_sorted(2,3,['file.xlsx', 'onglet2', 2, [5,6,7,8]]) 
                  
        """
        
        file_to_copy = openpyxl.load_workbook(self.path + other_sheet[0],data_only=True)
        sheet_to_copy = file_to_copy[other_sheet[1]]
        columns_to_copy = other_sheet[3]
        dico = {}

        #Passage en revue les identifiants du premier fichier et création d'un dictionnaire dont les clés sont ces identifiants et les valeurs sont une liste de valeurs à récupérer.
        for i in range(1,sheet_to_copy.max_row + 1):
            value = sheet_to_copy.cell(i,other_sheet[2]).value
            dico[value] = [sheet_to_copy.cell(i,j) for j in columns_to_copy]

        self.sheet.insert_cols(column_insertion,len(columns_to_copy)) 

        #Passage en revue des identifiants du second fichier et insertion des valeurs si les identifiants sont dans les clés du dico
        #. 
        for i in range(1,self.sheet.max_row+1):
            key = self.sheet.cell(i,column_identifiant).value
            if key in dico.keys():
                for j in range(len(columns_to_copy)):
                    self.sheet.cell(i,column_insertion + j).value = dico[key][j].value
                    self.sheet.cell(i,column_insertion + j).fill = copy(dico[key][j].fill)
        
        self.writebook.save(self.path + self.name_file)


    def color_lines_containing_chaines(self,color,*chaines):
        """
        Fonction qui colore les lignes dont une des cases contient une str particulière.

        Input : 
            - color : une couleur indiquée en haxadécimal par l'utilisateur.
            - chaines : des chaines de caractères que l'utilisateur entre et qui entraînent la coloration de la ligne.
        
        Exemple d'utilisation : 
    
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.color_lines_containing_chaine('#FF0000', 'vrai', 'hello', 'heeenri', 'ficheux') 
        
        """

        lines_to_color = []

        for i in range(1, self.sheet.max_row + 1):
            for j in range(1, self.sheet.max_column + 1):
                if self.sheet.cell(i,j).value in chaines:
                    lines_to_color.append(i)
                    break
        
        for row in lines_to_color:
            self.color_line(color, row)
        
        self.writebook.save(self.path + self.name_file)

    def column_cut_string_in_parts(self,column_to_cut,column_insertion,separator):
        """
        Fonction qui prend une colonne dont chaque cellule contient une grande chaîne de
          caractères. Toutes les chaînes sont composés du nombre de morceaux délimités par un séparateur,
        La fonction insère autant de colonnes que de morceaux et place un morceau par colonne dans l'ordre des morceaux.

        Inputs :
            - column_to_cut : colonne contenant les grandes str.
            - column_insertion : où insérer les colonnes
            - separator le séparateur

        Exemple d'utilisation : 
    
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.column_cut_string_in_parts(3, 10, ';') 
        
        """
        
        for i in range(2, self.sheet.max_row + 1):
            value = self.sheet.cell(i,column_to_cut).value
            chaine_object = Str(value)
            parts = chaine_object.cut_string_in_parts(separator)
            if i == 2:
                self.sheet.insert_cols(column_insertion,len(parts))
            for j in range(len(parts)):
                self.sheet.cell(i,column_insertion + j).value = parts[j]

        self.writebook.save(self.path + self.name_file) 

    def delete_lines(self,column,*chaines):
        """
        Fonction qui parcourt une colonne et qui supprime la ligne si celle-ci contient une chaîne particulière.

        Inputs : 
            -column : la colonne à parcourir.
            -chaines : les chaînes de caractères qui doivent engendrer la suppression de la ligne.
        
        Exemple d'utilisation : 
    
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.delete_lines(3, 'chaine1', 'chaine2', 'chaine3', 'chaine4') 
        """

        for i in range(1,self.sheet.max_row + 1):  
            if self.sheet.cell(i,column).value in chaines:
                self.sheet.delete_rows(i)

        self.writebook.save(self.path + self.name_file)


   
      



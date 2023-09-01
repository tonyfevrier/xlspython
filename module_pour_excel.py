#from xlutils.copy import copy 

import openpyxl 
from openpyxl.styles import PatternFill
from copy import copy
from datetime import date, datetime


class Path():
    def __init__(self,path = 'fichiers_xls/'):
        self.path = path
        
    def act_on_files(self,fonction):
        """
        Fonction qui prend tous les fichiers d'un dossier et qui applique une même action à ces fichiers.
        """
        pass 


class File(): 
    def __init__(self,name_file, path = 'fichiers_xls/'):
        """L'utilisateur sera invité à mettre son fichier xslx dans un dossier nommé fichiers_xls"""
        self.name_file = name_file  
        self.path = path
        self.writebook = openpyxl.load_workbook(self.path + self.name_file, data_only=True)
        self.sheets_name = self.writebook.sheetnames

    def sauvegarde(self):
        """
        Fonction qui crée une sauvegarde du fichier name_file et qui l'appelle name_file_numero où le numéro est le premier qui n'a pas été utilisé.
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

class Sheet(File): 
    def __init__(self, name_file, name_onglet,path = 'fichiers_xls/'): 
        super().__init__(name_file,path)
        self.name_onglet = name_onglet  
        self.sheet = self.writebook[self.name_onglet]
        del self.sheets_name

    def column_transform_string_in_binary(self,column_read,column_write,*good_answers,line_beginning = 2, line_end = 100, insert = True, security = True):
        """
        Fonction qui prend une colonne de str et qui renvoie une colonne de 0 ou de 1
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où mettre les 0 ou 1.
        Input : good_answers : une séquence d'un nb quelconque de bonnes réponses qui valent 1pt. Chaque mot ne doit pas contenir d'espace ni au début ni à la fin.
                column_read : la colonne de lecture des réponses.
                colum_write : la colonne d'écriture des 0 et 1. 
                line_beggining, line_end : intervalle de ligne dans lequel l'utilisateur veut appliquer sa transformation
        
        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.
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

    def column_set_answer_in_group(self,column_read,column_write,groups_of_responses,line_beginning = 2, line_end = 100, insert = True, security = True):
        """
        Fonction qui prend une colonne de str et qui renvoie une colonne remplie de str correspondant à des noms de groupes.
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où écrire
        Input : groups_of_response : dictionnary whick keys are response groups and which values are a list of responses 
        associated to this group.
                column_read : la colonne de lecture des réponses.
                colum_write : la colonne d'écriture des 0 et 1. 
                line_beggining, line_end : intervalle de ligne dans lequel l'utilisateur veut appliquer sa transformation
        
        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.
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
    
    def column_security(self,column):
        """
        Fonction qui prend une colonne et regarde si la colonne est vide.
        Input : column
        Output : True si elle ne contient rien, False sinon
        """
        bool = True
        for i in range(1,self.sheet.max_row+1): 
            if self.sheet.cell(i,column).value != None:
                bool = False
                break
        return bool 
        
    def color_special_cases_in_column(self,column,chainecolor):
        """
        Fonction qui regarde pour une colonne donnée colore les cases contenant à certaines chaînes de caractères
        Input : 
            - column : le numéro de la colonne.
            - chainecolor : les str qui vont être colorés et les couleurs qui correspondent à écrire avec la syntaxe suivante {'vrai':'couleur1','autre':couleur2}. Attention,
                la couleur doit être entrée en hexadécimal et les chaînes de caractères ne doivent pas avoir d'espace au début ou à la fin.
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
            - chainecolor : les str qui vont être colorés et les couleurs qui correspondent à écrire avec la syntaxe suivante {'vrai':'couleur1','autre':couleur2}. Attention,
                la couleur doit être entrée en hexadécimal et les chaînes de caractères ne doivent pas avoir d'espace au début ou à la fin.
        """

        for j in range(1, self.sheet.max_column + 1):
            self.color_special_cases_in_column(j,chainecolor)

    def add_column_in_sheet_differently_sorted(self,column_identifiant, column_insertion,other_sheet):
        """
        Fonction qui insère dans une feuille des colonnes d'une autre feuille de référence. Les deux feuilles ont des lignes qui ne sont pas triées dans le même ordre.
        Les deux feuilles ont une colonne d'identifiants (exemple : des mails).
        La fonction récupère un ou plusieurs éléments d'une ligne déterminée par un identifiant, recherche l'identifiant dans la seconde feuille, insère les éléments
        dans la ligne correspondante et dans les colonnes insérées. 

        Je passe en revue dans l'ordre les identifiants du premier fichier et je crée un dictionnaire dont les clés sont ces identifiants et les valeurs sont une liste de valeurs à récupérer.
        Je passe en revue dans l'ordre (qui est différent du premier) les identifiants du second fichier et j'y insère les valeurs si les identifiants sont dans les clés du dico, sinon je laisse les cases vides. 
        Cela évite de parcourir pleins de fois les identifiants en les recherchant.

        Inputs :
            - column_identifiant : numéro de la colonne où sont situés les identifiants dans le fichier qu'on souhaite modifier.
            - column_insertion : numéro de la colonne où on insère les colonnes à récupérer.
            - other_sheet = ['namefile','namesheet',numéro de la colonne où sont les identifiants,[numéros des colonnes à récupérer sous forme de liste]]
                namefile doit être au format .xlsx et mis dans le dossier fichier_xls.
        """
        
        file_to_copy = openpyxl.load_workbook(self.path + other_sheet[0],data_only=True)
        sheet_to_copy = file_to_copy[other_sheet[1]]
        columns_to_copy = other_sheet[3]
        dico = {}

        for i in range(1,sheet_to_copy.max_row + 1):
            value = sheet_to_copy.cell(i,other_sheet[2]).value
            dico[value] = [sheet_to_copy.cell(i,j) for j in columns_to_copy]


        self.sheet.insert_cols(column_insertion,len(columns_to_copy)) 

        for i in range(1,self.sheet.max_row+1):
            key = self.sheet.cell(i,column_identifiant).value
            if key in dico.keys():
                for j in range(len(columns_to_copy)):
                    self.sheet.cell(i,column_insertion + j).value = dico[key][j].value
                    self.sheet.cell(i,column_insertion + j).fill = copy(dico[key][j].fill)
        
        self.writebook.save(self.path + self.name_file)

    def color_line(self, color, row_number):
        """
        Fonction qui colore une ligne spécifique

        Input :
            - color : une couleur indiquée en haxadécimal par l'utilisateur.
            - row_number : le numéro de la ligne à colorer
        """

        for j in range(1, self.sheet.max_column + 1):
            self.sheet.cell(row_number,j).fill = PatternFill(fill_type = 'solid', start_color = color)


    def color_lines_containing_chaines(self,color,*chaines):
        """
        Fonction qui colore les lignes dont une des cases contient une str particulière.

        Input : 
            - color : une couleur indiquée en haxadécimal par l'utilisateur.
            - chaines : des chaines de caractères que l'utilisateur entre et qui entraînent la coloration de la ligne.
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

    def column_cut_str_in_parts(self,column_to_cut,column_insertion,separator):
        """
        Fonction qui prend une colonne contenant dans chaque cellule une grande chaîne de caractères contenant le même nombre de morceaux séparés par un séparateur,
        qui insère autant de colonnes que de morceaux et qui place un morceau par colonne dans l'ordre.

        Inputs :
            - column_to_cut : colonne contenant les grandes str.
            - column_insertion : où insérer les colonnes
            - separator le séparateur
        """
        
        for i in range(2, self.sheet.max_row + 1):
            value = self.sheet.cell(i,column_to_cut).value
            chaine_object = Str(value)
            parts = chaine_object.cut_str_in_parts(separator)
            if i == 2:
                self.sheet.insert_cols(column_insertion,len(parts))
            for j in range(len(parts)):
                self.sheet.cell(i,column_insertion + j).value = parts[j]

        self.writebook.save(self.path + self.name_file) 

    def delete_lines(self,column,*chaines):
        """
        Fonction qui parcourt une colonne et qui supprime la ligne si celle-ci contient un élément particulier dans args.

        Inputs : 
            -column : la colonne à parcourir.
            -chaines : les str qui doivent engendrer la suppression de la ligne.
        """
        for i in range(1,self.sheet.max_row + 1):  
            if self.sheet.cell(i,column).value in chaines:
                self.sheet.delete_rows(i)

        self.writebook.save(self.path + self.name_file)


    def create_dico_from_columns(self, column_keys:int, column_values:int, first_line, last_line):
        """
        Function returning a dictionnary whose keys are elements of a column if they are not empty and values are elements of an other column
        column_keys : column whose elements are the keys of the dictionnary.
        column_vakues : same with values
        first_line : the line we begin to read the file
        """
        dico = {}
        for i in range(first_line,last_line):
            key = self.sheet.cell(i,column_keys).value 
            
            print(key,self.sheet.max_row+1)
            if key != "":
                dico[key] = self.sheet.cell(i,column_values).value
        return dico

class Str():
    def __init__(self,chaine):
        self.chaine = str(chaine)
        

    def transform_string_in_binary(self,*args):
        """
        Fonction qui prend un str et qui le transforme en 0 ou 1

        Inputs : args : des chaînes de caractère devant renvoyer 1 
        Outputs : bool : 0 ou 1.
        """
        bool = 0
        if self.chaine in args:
            bool = 1
        return bool
    
    def set_answer_in_group(self, groups_of_response):
        """
        Function which takes a response and return a string of the group containing the response.
        
        Input : groups_of_response : dictionnary whick keys are response groups and which values are a list of responses 
        associated to this group.
        Output : the string of the group containing the response. 
        """

        """
        for group in groups_of_response.keys():
            if self.chaine in groups_of_response[group] :
                return group
        return ""
        """
        if self.chaine in groups_of_response.keys():
            return groups_of_response[self.chaine]
        else:
            return ""
        
    
    def clean_string(self):
        """
        Fonction qui prend une chaîne de caractère et qui élimine tous les espaces de début et de fin.
        Fonction qui nettoie également les espaces insécables \xa0 par un espace régulier.
        Ceci rendra une chaîne de caractère qui remplacera l'attribut chaine de la classe.  
        On pourra ainsi éviter les erreurs liées à une différence d'un seul espace.      
        """
        depart = 0
        fin = len(self.chaine)
        while self.chaine[depart] == ' ' or self.chaine[fin-1] == ' ':
            if self.chaine[depart] == ' ':
                depart += 1
            if self.chaine[fin-1] == ' ':
                fin -= 1
        self.chaine = self.chaine[depart:fin]
 
        chaine2 = self.chaine.replace('\xa0', ' ')
        self.chaine = chaine2
        return self
    
    def del_extension(self):
        """Fonction 
            - qui enlève l'extension d'un nom de fichier si le nom ne contient pas de date
            - qui ne garde que la partie avant _date_ pour un fichier nommé test_date_****-**-**.xlsx. 
            - qui sert à la sauvegarde et permet ainsi d'éviter des noms à rallonge.
        """
        position = self.chaine.find('_date_')
        if position == -1: 
           position = self.chaine.find('.xlsx')

        return self.chaine[:position]
    
    def cut_str_in_parts(self, separator):
        """
        Fonction qui prend une chaîne de caractères contenant plusieurs sous-chaînes séparées par un séparateur et qui les sépare en plusieurs sous-chaînes.

        Input : separator

        Output : Un tuple contenant les morceaux de chaînes.
        """ 

        parts = ()
        chaine = self.chaine

        debut_part = 0

        for i in range(len(chaine)):
            if chaine[i] == separator:
                parts = parts + (chaine[debut_part:i],) 
                debut_part = i+1

        parts = parts + (chaine[debut_part:],) 
        return parts

    



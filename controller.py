import os
import openpyxl
import json
from openpyxl.utils import column_index_from_string, coordinate_to_tuple
from time import time
from utils import Other, UtilsForFile 
from copy import copy




class FileControler(UtilsForFile):
    def __init__(self, file):
        """
        Input: file (object of class File)
        """
        self.file = file

    def create_one_onglet_by_participant(self, onglet_from, column_read, newfile_name, newfile_path, first_line=2):
        """
        Fonction qui prend un onglet dont une colonne contient des chaînes de caractères comme par exemple un nom.
        Chaque chaîne de caractères peut apparaître plusieurs fois dans cette colonne (exe : quand un participant répond plusieurs fois)
        La fonction retourne un fichier contenant un onglet par chaîne de caractères.
          Chaque onglet contient toutes les lignes correspondant à cette chaîne de caractères.

        Input : 
            onglet_from : onglet de référence.
            column_read : l'étiquette de la colonne qui contient les chaînes de caractères.
            newfile_name (str): name of the newfile
            newfile_path (str): where to write/find the newfile
            first_line : ligne où commencer à parcourir.
            last_line : ligne de fin de parcours 
 
        Exemple d'utilisation : 
    
            file = File('dataset.xlsx')
            file.create_one_onglet_by_participant('onglet1', 'A') 
        """ 
        column_read = column_index_from_string(column_read)  

        sheet = self.file.writebook[onglet_from] 

        # Creation of the file aiming to contain the data if it does not already exists 
        if newfile_name not in os.listdir(newfile_path):
            new_file = openpyxl.Workbook()
        else:
            new_file = openpyxl.load_workbook(newfile_path + newfile_name)

        onglets = new_file.sheetnames

        # Create one tab by identifiant containing all its lines
        for i in range(first_line, sheet.max_row + 1):
            onglet = str(sheet.cell(i,column_read).value)

            # Prepare a new tab
            if onglet not in onglets:
                new_file.create_sheet(onglet)
                self.copy_paste_line(sheet, 1,  new_file[onglet], 1)
                onglets.append(onglet) 

            self.add_line_at_bottom(sheet, i, new_file[onglet]) 
        
        # Deletion of the first tab if the newfile was created
        if newfile_name not in os.listdir(newfile_path):
            del new_file[new_file.sheetnames[0]]
        new_file.save(newfile_path + newfile_name)
        

    def extract_column_from_all_sheets(self, column):
        """
        Fonction qui récupère une colonne dans chaque onglet pour former une nouvelle feuille
        contenant toutes les colonnes. La première cellule de chaque colonne correspond alors 
        au nom de l'onglet. Attention, en l'état, il faut que tous les onglets aient la même structure.

        Input : 
            column : str. L'étiquette de la colonne à récupérer dans chaque onglet 

        Exemple d'utilisation : 
    
            file = File('dataset.xlsx')
            file.extract_column_from_all_sheets('B') 

            Si on veut extraire les formules

            file = File('dataset.xlsx',dataonly = False)
            file.extract_column_from_all_sheets('B') 
        """ 
        column = column_index_from_string(column)
         
        new_sheet = self.writebook.create_sheet(f"gather_{column}")
        column_to = 1

        start = time()
        for name_onglet in self.sheets_name: 
            onglet = self.file.writebook[name_onglet] 
            self.copy_paste_column(onglet,column,new_sheet,column_to)
            column_to = new_sheet.max_column + 1
            new_sheet.cell(1,new_sheet.max_column).value = name_onglet 
            Other.display_running_infos('extract_column_from_all_sheets', name_onglet, self.sheets_name, start)

            
        self.file.writebook.save(self.file.path + self.file.name_file) 
        self.file.sheets_name = self.file.writebook.sheetnames 

    def extract_cells_from_all_sheets(self, *cells):
        """
        Vous avez un fichier avec des onglets de structure identique avec un onglet par participant. Vous souhaitez
        récupérer des cellules identiques dans tous les onglets et créer un onglet avec une ligne par participant,
        qui contient les valeurs de ces cellules. Fonction analogue à gather_multiple_answers mais ne portant pas sur une
        seule feuille.

        Inputs:
            - cells (list[str])
        """ 

        # Recover cells coordinates
        cell_list = []
        for cell in cells: 
            cell_list.append(coordinate_to_tuple(cell)) 
        
        # Create a new tab
        gathered_sheet = self.file.writebook.create_sheet('gathered_data')
        current_line = 2

        start = time()

        # Fill one line by tab
        for name_onglet in self.file.sheets_name:   
            current_onglet = self.file.writebook[name_onglet]
            gathered_sheet.cell(current_line, 1).value = name_onglet
            current_column = 2

            # Fill selected values one by one
            for tuple in cell_list:  
                gathered_sheet.cell(current_line, current_column).value = current_onglet.cell(tuple[0],tuple[1]).value
                current_column += 1
            current_line += 1
            Other.display_running_infos('extract_cells_from_all_sheets', name_onglet, self.file.sheets_name, start)


        self.file.sheets_name = self.file.writebook.sheetnames 
        self.file.writebook.save(self.file.path + self.file.name_file)
        

    def apply_column_formula_on_all_sheets(self, *column_list):
        """
        Fonction qui reproduit les formules d'une colonne ou plusieurs colonnes
          du premier onglet sur toutes les colonnes situées à la même position dans les 
          autres onglets.

        Input : 
            -column_list : int. les positions des colonnes où récupérer et coller.

        Exemples d'utilisation : 

            Bien veiller à mettre dataonly = False sinon il ne copiera pas les formules mais
            les valeurs des cellules. On peut aussi copier les valeurs des cellules : pour cela,
            enlever dataonly = False.

            Sur une colonne
                file = File('dataset.xlsx', dataonly = False)
                file.apply_column_formula_on_all_sheets(2) 

            Sur trois colonnes
                file = File('dataset.xlsx', dataonly = False)
                file.apply_column_formula_on_all_sheets(2,5,10) 

            Sur toutes les colonnes du fichier à partir de la colonne colmin jusque la colonne colmax :
                file = File('dataset.xlsx', dataonly = False)
                file.apply_column_formula_on_all_sheets(*[i for i in range(colmin,colmax + 1)]) 
        """
        column_int_list = []
        for column in column_list: 
            column_int_list.append(column_index_from_string(column))  

        start = time()

        #on applique les copies dans tous les onglets sauf le premier
        for name_onglet in self.file.sheets_name[1:]:
            for column in column_int_list:
                self.copy_paste_column(self.file.writebook[self.file.sheets_name[0]],column,self.file.writebook[name_onglet],column)
            Other.display_running_infos('apply_column_formula_on_all_sheets', name_onglet, self.file.sheets_name[1:], start)
            

        self.file.writebook.save(self.file.path + self.file.name_file)

    def apply_cells_formula_on_all_sheets(self, *cells):
        """
        Fonction qui reproduit les formules d'une cellule ou plusieurs cellules
          du premier onglet sur toutes les cellules situées à la même position dans les 
          autres onglets.

        Input : 
            -cells : string. les positions des cellule où récupérer et coller 

        Exemples d'utilisation : 

            Bien veiller à mettre dataonly = False sinon il ne copiera pas les formules mais
            les valeurs des cellules. On peut aussi copier les valeurs des cellules : pour cela,
            enlever dataonly = False.

            file = File('dataset.xlsx', dataonly = False)
            file.apply_column_formula_on_all_sheets('C5','D6')  
        """

        #obtenir les indices de la cellule 
        cell_list = []
        for cell in cells: 
            cell_list.append(coordinate_to_tuple(cell)) 

        start = time()

        #on applique les copies dans tous les onglets sauf le premier
        for name_onglet in self.sheets_name[1:]:   
            for tuple in cell_list: 
                self.file.writebook[name_onglet].cell(tuple[0],tuple[1]).value = self.file.writebook[self.file.sheets_name[0]].cell(tuple[0],tuple[1]).value  
                self.file.writebook[name_onglet].cell(tuple[0],tuple[1]).fill = copy(self.file.writebook[self.file.sheets_name[0]].cell(tuple[0],tuple[1]).fill)  
                self.file.writebook[name_onglet].cell(tuple[0],tuple[1]).font = copy(self.file.writebook[self.file.sheets_name[0]].cell(tuple[0],tuple[1]).font)  
                self.file.writebook[name_onglet].cell(tuple[0],tuple[1]).border = copy(self.file.writebook[self.sheets_name[0]].cell(tuple[0],tuple[1]).border)  
                self.writebook[name_onglet].cell(tuple[0],tuple[1]).alignment = copy(self.writebook[self.file.sheets_name[0]].cell(tuple[0],tuple[1]).alignment)    
            Other.display_running_infos('apply_cells_formula_on_all_sheets', name_onglet, self.file.sheets_name[1:], start)

        self.file.writebook.save(self.file.path + self.file.name_file)

    def gather_columns_in_one(self, onglet, *column_lists):
        """
        Vous avez des groupes de colonnes de valeurs avec une étiquette en première cellule. Pour chaque groupe, vous souhaitez former deux colonnes de valeurs : l'une qui contient
        les valeurs rassemblées en une colonne, l'autre, à sa gauche, qui indique l'étiquette de la colonne dans laquelle elle a été prise.

        Inputs : 
            - onglet (str) : nom de l'onglet d'où on importe les colonnes.
            - column_lists (list[list[str]]) : liste de groupes de colonnes. Chaque groupe est une liste de colonnes.
        """
         
        for list in column_lists:
            tab_number = len(self.file.sheets_name)
            self.file.writebook.create_sheet(f"onglet {tab_number}")
            target_sheet = self.file.writebook[f"onglet {tab_number}"]
            for column in list: 
                self.copy_column_tags_and_values_at_bottom(self.file.writebook[onglet], column_index_from_string(column), target_sheet)

        self.file.writebook.save(self.file.path + self.file.name_file) 

    def build_file_from_tab(self, tab):
        """
        Fonction qui prend un nom d'onglet dans un fichier et qui crée un fichier associé.

        Input :
            - tab (str) : the name of the tab from which we want to create the file.
        """

        sheet_from = self.file.writebook[tab]
        newfile = openpyxl.Workbook() 
        sheet_to = newfile['Sheet']  
        path = 'multifiles/' 
  
        self.deep_copy_of_a_sheet(sheet_from, sheet_to) 

        namefile = path + tab + '.xlsx'
        newfile.save(namefile) 
        return namefile
            
            
    def one_file_by_tab_sendmail(self, send = False, adressjson = "", objet = "", message = ""):
        """
        Vous souhaitez fabriquer un fichier par onglet. Chaque fichier aura le nom de l'onglet. 
        Vous souhaitez éventuellement envoyer chaque fichier à la personne associée.
        Attention, pour utiliser cette fonction, les onglets doivent être de la forme "prenom nom" sans caractère spéciaux. 

        Inputs : 
            send(optional boolean) : True si on veut envoyer le mail, False si on veut juste couper en fichiers.
            adressjson(str) : nom du fichier xlsx qui contient deux colonnes la première avec les noms des onglets, la seconde avec l'adresse mail. Ce fichier doit être mis dans le dossier fichier_xls. 
            objet(optional str) : Objet du message.
            message (optional str) : Contenu du message.
        """ 
        if adressjson != "":
            file = open(self.file.path + adressjson, 'r')
            mailinglist = json.load(file)
            file.close()

        start = time()

        for tab in self.file.sheets_name: 

            file_to_send = self.build_file_from_tab(tab)
            if send:
                if adressjson == "":
                    prenom = tab.split(" ")[0]
                    nom = tab.split(" ")[1]
                    self.envoi_mail(prenom + "." + nom + "@universite-paris-saclay.fr", file_to_send, "tony.fevrier62@gmail.com", "qkxqzhlvsgdssboh", objet, message)
                else: 
                    self.envoi_mail(mailinglist[tab], file_to_send, "tony.fevrier62@gmail.com", "qkxqzhlvsgdssboh", objet, message) 
            Other.display_running_infos('one_file_by_tab_sendmail', tab, self.file.sheets_name, start)
             

    def merge_cells_on_all_tabs(self, start_column, end_column, start_row, end_row):
        """
        Fonction qui merge les mêmes cellules sur tous les onglets d'un fichier 

        Inputs :
            - start_column (string): Letter of the column where to start the merging
            - end_column (string): Letter of the column where to end the merging
            - start_row (int): Index of the row where to start the merging
            - end_row (int): Index of the row where to start the merging

        """
        
        start_column = column_index_from_string(start_column)
        end_column = column_index_from_string(end_column)

        start = time()

        for tab in self.file.sheets_name: 
            sheet = self.file.writebook[tab] 
            sheet.merge_cells(start_row=start_row, start_column=start_column, end_row=end_row, end_column=end_column)
            Other.display_running_infos('merge_cells_on_all_tabs', tab, self.file.sheets_name, start)

        self.file.writebook.save(self.file.path + self.file.name_file)

    def check_linenumber_of_tabs(self, line_number):
        """
        Fonction qui prend un fichier et qui contrôle si tous les onglets ont un nombre de lignes égal à l'argument
        """
        wrong_tabs = []
        for tab in self.file.sheets_name:
            if self.file.writebook[tab].max_row != line_number:
                wrong_tabs.append(tab)
        return wrong_tabs
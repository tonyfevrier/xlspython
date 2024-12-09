"""Handle methods reading and modifying a unique tab of a file."""

# Ce que j'ai fait pr les classes. 
# Attributs : initiaux, mais aussi transitoires pour éviter de créer des arguments dans les fonctions

import os
import openpyxl
import json
import re

from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter, coordinate_to_tuple
from time import time
from utils.utils import Other, String, UtilsForFile, Str, DisplayRunningInfos, TabsCopy, GetIndex, TabUpdate, ColumnDelete, ColumnInsert, LineDelete, LineInsert
from copy import copy
from model.model_factorise import Cell 


def create_empty_workbook():
    workbook = openpyxl.Workbook()
    del workbook[workbook.active.title]
    return workbook


class ColorTabController(String):
    """Handle methods coloring a unique tab of a file."""

    def __init__(self, file_object, tab_name=None, tab_options=None, color=None, first_line=2):
        """
        Attributs:
            - file_object (file object)
            - tab_name (str)
            - tab (openpyxl.workbook.tab)            
            - tab_options (TabOptions object)
            - first_line (optional int)
        """ 
        self.file_object = file_object 
        if tab_name is not None:
            self.tab = self.file_object.get_tab_by_name(tab_name)
        self.tab_options = tab_options 
        self.first_line = first_line 
        self.color = color

    def reinitialize_storing_attributes(self):
        pass

    def color_cases_in_column(self, map_string_to_color):
        """
        Fonction qui pour une colonne donnée colore les cases correspondant à certaines chaînes de caractères.
        """  
        column_index = column_index_from_string(self.tab_options.column_to_read) 
        
        for i in range(self.first_line, self.tab.max_row + 1):
            cell = self.tab.cell(i, column_index) 
            self._color_cell_if_contains_string(cell, map_string_to_color) 

    def _color_cell_if_contains_string(self, cell, map_string_to_color):
        cleaned_cell_value = self._get_cleaned_cell_value(cell)
        if cleaned_cell_value in map_string_to_color.keys():
            cell.fill = PatternFill(fill_type = 'solid', start_color = map_string_to_color[cleaned_cell_value])

    def _get_cleaned_cell_value(self, cell): 
        if cell.value is str:
            cleaned_cell_value = self.clean_string_from_spaces(cell.value) 
        else: 
            cleaned_cell_value = cell.value
        return cleaned_cell_value

    def color_cases_in_sheet(self, map_string_to_color): 
        """
        Fonction qui colore les cases contenant à certaines chaînes de caractères d'une feuille 
        """  

        for j in range(1, self.tab.max_column + 1):
            self.tab_options.column_to_read = get_column_letter(j) 
            self.color_cases_in_column(map_string_to_color)

    def color_lines_containing_strings(self, *strings):
        """
        Fonction qui colore les lignes dont une des cases contient une str particulière.
        """ 

        lines_indexes = self._list_lines_containing_strings(*strings)
        self.color_lines(lines_indexes) 

    def _list_lines_containing_strings(self, *strings):
        lines_indexes = []
        for line_index in range(self.first_line, self.tab.max_row + 1):
            lines_indexes = self._add_line_containing_strings_to_list(strings, line_index, lines_indexes)
        return lines_indexes
    
    def _color_line(self, line_index): 
        for column_index in range(1, self.tab.max_column + 1):
            self.tab.cell(line_index, column_index).fill = PatternFill(fill_type = 'solid', start_color = self.color)

    def _add_line_containing_strings_to_list(self, strings, line_index, lines_indexes):
        for column_index in range(1, self.tab.max_column + 1):
            if self.file_object.get_compiled_cell_value(self.tab, Cell(line_index, column_index)) in strings:
                lines_indexes.append(line_index)
                break
        return lines_indexes
    
    def color_lines(self, list_of_lines):
        for line_index in list_of_lines:
            self._color_line(line_index) 
    

class DeleteController(String):
    """Handle methods deleting lines or columns of a tab"""

    def __init__(self, file_object, tab_name=None, first_line=2):
        self.file_object = file_object
        self.tab_name = tab_name
        self.tab = None
        if tab_name is not None:
            self.tab = self.file_object.get_tab_by_name(tab_name)  
        self.tab_update = TabUpdate()
        self.columns_to_delete = []
        self.lines_to_delete = []
        self.first_line = first_line  
    
    def reinitialize_storing_attributes(self):
        self.columns_to_delete = []
        self.lines_to_delete = []

    def update_cell_formulas(self, modification_object): 
        self.tab_update.choose_modifications_to_apply(modification_object)  
        self.tab_update.update_cells_formulas(self.tab) 
    
    def delete_columns(self, string_of_columns):
        """
        Prend une séquence de colonnes sous la forme 'C-J,K,L-N,Z' qu'on souhaite supprimer. 
        """  

        # Réordonner par les lettres les plus grandes pour supprimer de la droite vers la gauche dans l'excel  
        self.columns_to_delete = self.get_columns_from(string_of_columns)
        self.columns_to_delete.sort(reverse = True)  

        for column_letter in self.columns_to_delete:  
            self.tab.delete_cols(column_index_from_string(column_letter)) 

        self.update_cell_formulas(ColumnDelete(self.columns_to_delete)) 

    def delete_other_columns(self, string_of_columns):
        """
        Prend une séquence de colonnes sous la forme 'C-J,K,L-N,Z' et supprime les autres 
        """  
        columns_to_keep = self.get_columns_from(string_of_columns)
        list_all_columns = self._get_list_of_columns()

        #Réordonner par les lettres les plus grandes pour supprimer de la droite vers la gauche
        list_all_columns.sort(reverse=True) 

        for column_letter in list_all_columns: 
            self._delete_column_not_to_keep(column_letter, columns_to_keep)
       
        self.update_cell_formulas(ColumnDelete(self.columns_to_delete))
        self.file_object.save_file() 

    def _delete_column_not_to_keep(self, column_letter, columns_to_keep):
        if column_letter not in columns_to_keep:
            self.columns_to_delete.append(column_letter)
            self.tab.delete_cols(column_index_from_string(column_letter)) 

    def _get_list_of_columns(self):
        return [get_column_letter(column_index) for column_index in range(1, self.tab.max_column + 1)]

    def delete_lines_containing_strings_in_given_column(self, column_letter, *strings):
        """
        Fonction qui parcourt une colonne et qui supprime la ligne si celle-ci contient une chaîne particulière.

        """ 

        column_index = column_index_from_string(column_letter)  

        # On part de la plus grande ligne pour éviter qu'une suppression ne change la position d'une ligne à supprimer après
        for line_index in range(self.tab.max_row, 0, -1):
            cell_value = self.file_object.get_compiled_cell_value(self.tab, Cell(line_index,column_index)) 
            self._delete_line_containing_strings(line_index, cell_value, *strings)
 
        self.update_cell_formulas(LineDelete(self.lines_to_delete))   

    def _delete_line_containing_strings(self, line_index, cell_value, *strings):
        if str(cell_value) in strings:  
            self.tab.delete_rows(line_index) 
            self.lines_to_delete.append(str(line_index))     

    def delete_twins_lines_and_color_last_twin(self, column_identifier, color = 'FFFFFF00'):
        """
        Certains participants répondent plusieurs fois à une étude. Cette fonction supprime les premières lignes réponses
        des participants dans ce cas. Elle ne garde que leur dernière réponse. On repère les participants
        par leur identifiant unique donné dans colum_identifiant.
        """

        column_identifier = column_index_from_string(column_identifier) 

        map_identifier_to_line = {}  

        #On parcourt dans le sens inverse afin d'éviter que la suppression progressive impacte la position des lignes étudiées ensuite. 
        for line_index in range(self.tab.max_row, 0, -1):
            cell_identifier = Cell(line_index, column_identifier)
            self._delete_line_and_color_last_twin(cell_identifier, map_identifier_to_line, color)
             
        self.update_cell_formulas(LineDelete(self.lines_to_delete))       

    def _delete_line_and_color_last_twin(self, cell_identifier, map_identifier_to_line, color):
        """
        map_identifier_to_line stocke tous les identifiants et la ligne associée. Si un identifiant I1
        est déjà dans le dictionnaire, on va supprimer sa ligne et donc décaler toutes les lignes en-dessous 
        de -1, il faut donc baisser la ligne de I1 de 1 car si I1 vient une troisième fois, on ne colorerait pas 
        la bonne ligne 
        """
        cell_value = self.file_object.get_compiled_cell_value(self.tab, cell_identifier)
        identifier = self.clean_string_from_spaces(cell_value) 

        if identifier in map_identifier_to_line.keys(): 
            self._color_line(map_identifier_to_line[identifier], color)
            self.tab.delete_rows(cell_identifier.line_index)
            self.lines_to_delete.append(str(cell_identifier.line_index))
            map_identifier_to_line[identifier] -= 1    
        else:
            map_identifier_to_line[identifier] = cell_identifier.line_index 

        return map_identifier_to_line

    def _color_line(self, line_index, color):  
        for column_index in range(1, self.tab.max_column + 1):
            self.tab.cell(line_index, column_index).fill = PatternFill(fill_type = 'solid', start_color = color)  
        

class InsertController():
    """Handle methods inserting columns in a tab""" 

    def __init__(self, file_object, tab_name=None, tab_options=None, first_line=2):
        """
        Attributs:
            - file_object (file object)
            - tab_name (str)
            - tab (openpyxl.workbook.tab)            
            - tab_options (TabOptions object)
            - first_line (optional int)
        """ 
        self.file_object = file_object
        self.tab_name = tab_name
        if tab_name is not None:
            self.tab = self.file_object.get_tab_by_name(tab_name)
        self.tab_update = TabUpdate()
        self.tab_options = tab_options 
        self.first_line = first_line 

    def reinitialize_storing_attributes(self):
        pass

    def update_cell_formulas(self, modification_object): 
        self.tab_update.choose_modifications_to_apply(modification_object)  
        self.tab_update.update_cells_formulas(self.tab) 
 
    def insert_splitted_strings_of(self, column_to_split, separator):
        """
        Fonction qui prend une colonne dont chaque cellule contient une grande chaîne de
          caractères. Toutes les chaînes sont composés du nombre de morceaux délimités par un séparateur,
        La fonction insère autant de colonnes que de morceaux et place un morceau par colonne dans l'ordre des morceaux.
        """ 
        column_to_split = column_index_from_string(column_to_split) 
        self.tab_options.column_to_write = column_index_from_string(self.tab_options.column_to_write)
        
        for line_index in range(self.first_line, self.tab.max_row + 1): 
            parts = self._get_string_and_split_it(Cell(line_index, column_to_split), separator)  
            self._insert_splitted_string(line_index, parts)

        modifications = [get_column_letter(column_to_split + index) for index in range(len(parts))]
        self.update_cell_formulas(ColumnInsert(modifications))   

    def _get_string_and_split_it(self, cell, separator):
        cell_value = self.file_object.get_compiled_cell_value(self.tab, cell) 
        return cell_value.split(separator)

    def _insert_splitted_string(self, line_index, parts):
        if line_index == self.first_line:
            self.tab.insert_cols(self.tab_options.column_to_write, len(parts))

        for part_index in range(len(parts)):
            self.tab.cell(line_index, self.tab_options.column_to_write + part_index).value = parts[part_index]
    
    def fill_one_column_by_QCM_answer(self, *answers):
        """
        Fonction qui recoit des réponses et crée une colonne par réponse. Elle regarde ensuite dans une cellule si la réponse 
        y est contenue. Si oui elle l'indique dans la colonne correspondante.
        """ 
        self.tab_options.column_to_read = column_index_from_string(self.tab_options.column_to_read) 
        self.tab_options.column_to_write = column_index_from_string(self.tab_options.column_to_write) 
        self._insert_answers_columns(answers)

        for line_index in range(self.first_line, self.tab.max_row + 1):
            self._check_columns_of_answers_contained(Cell(line_index, self.tab_options.column_to_read), answers)    

        modifications = [get_column_letter(self.tab_options.column_to_write + index) for index in range(len(answers))]
        self.update_cell_formulas(ColumnInsert(modifications))  

    def _insert_answers_columns(self, answers):
        self.tab.insert_cols(self.tab_options.column_to_write, len(answers))
        for index in range(0, len(answers)):
            self.tab.cell(1, index + self.tab_options.column_to_write).value = answers[index]

    def _check_columns_of_answers_contained(self, cell, answers):
        try:
            for index in range(0, len(answers)):  
                self._check_column_if_answer_contained(cell, (index, answers[index]))
        except TypeError:
            pass

    def _check_column_if_answer_contained(self, cell, answer):
        if answer[1] in self.tab.cell(cell.line_index, cell.column_index).value:
            self.tab.cell(cell.line_index, answer[0] + self.tab_options.column_to_write).value = 'X'

    #♥ ARRIVE ICI : A tester
    def gather_multiple_answers(self, sheet_name, column_read, column_store, line_beggining = 2):
        """
        Dans un onglet, nous avons les réponses de participants qui ont pu répondre plusieurs fois à un questionnaire.
        Cette fonction parcourt les noms et met dans un autre onglet. La ligne du participant est alors constituée des différentes valeurs
         d'une même donnée récupérée.
        
        Inputs :
            - column_read (str) : la colonne avec les identifiants des participants.
            - column_store (str) : lettre de la colonne contenant la donnée qu'on veut stocker.
            - line_beggining (int) : ligne où débute la recherche. 
        """ 
        sheet = self.file.writebook[sheet_name]

        column_read = column_index_from_string(column_read) 
        column_store = column_index_from_string(column_store) 

        #we create a dictionary whose keys are the identifiers (of participants) and values are their number of answers and a list containing
        #the data we want to store for each answer.
        dico = self.create_dico_to_store_multiple_answers_of_participants(sheet, column_read,column_store,line_beggining)
        
        #we create the new sheet where we store participants answering multiple times and their data.
        storesheet = self.file.writebook.create_sheet('severalAnswers')
        self.create_newsheet_storing_multiple_answers(storesheet, dico)

        self.file.writebook.save(self.file.path + self.file.name_file) 
    
    def act_on_columns(function):
        """
        Décorateur qui en plus d'appliquer la fonction, transforme les lettres de colonnes en index, met à jour les formules 
        après l'insertion de la colonne et sauvegarde le fichier.
        """
        def wrapper(self, *args, **kwargs):
            """
            - args[0] (str): sheet name
            - args[1] (list[str] or str): letters of columns to read
            - args[2] (str): letter of column in which to write
            """
            # Transform all args corresponding to columns in indexes 
            modifications = [args[2]]
            if isinstance(args[1], list):
                columns_read = [column_index_from_string(column) for column in args[1]]
            else:
                columns_read = column_index_from_string(args[1]) 
            column_insertion = column_index_from_string(args[2]) 

            # Apply the function on column indexes
            sheet = self.file.writebook[args[0]]
            sheet.insert_cols(column_insertion)
            function(self, args[0], columns_read, column_insertion, *args[3:], **kwargs)

            # Update eventual formulas and save
            self.updateCellFormulas(sheet, True, 'column', modifications)         
            #self.file.writebook.save(self.file.path + self.file.name_file)
        return wrapper
    
    @act_on_columns
    def map_two_columns_to_a_third_column(self, sheet_name, columns_read, column_insertion, mapping, line_beginning=2):
        """
        Vous avez deux colonnes de lecture, suivant ce qui est écrit sur une ligne, vous voulez ou non insérer quelque chose 
        dans une nouvelle colonne.

        Inputs:
            - columns_read (list[str]): liste de deux lettres contenant les colonnes de lecture.
            - column_insertion (str): lettre de la colonne où l'insertion doit avoir lieu.
            - mapping (dict): dictionnaire dont les clés sont les chaînes à écrire. Les valeurs sont dans l'ordre les 
            str qui si elles sont présentes, entraînent l'écriture de ces chaînes.
            - line_beggining (int) : ligne où débute la recherche.
        """ 
        sheet = self.file.writebook[sheet_name]

        for i in range(line_beginning, sheet.max_row + 1):
            # Fill the new column if columns read contain some expected values
            for key, value in mapping.items():
                value1 = str(sheet.cell(i, columns_read[0]).value)
                value2 = str(sheet.cell(i, columns_read[1]).value) 
                if [value1, value2] == value:
                    sheet.cell(i, column_insertion).value = key
                    break
    
    @act_on_columns
    def column_get_part_of_str(self, sheet_name, column_read, column_insertion, separator, piece_number, line_beginning=2):
        """
        Vous avez une colonne qui contient une chaîne dont vous voulez prendre le début jusqu'à un certain séparateur.
        Ce mot est inséré dans une nouvelle colonne.
        
        Inputs:
            - column_read (str): lettre de la colonne de lecture.
            - column_insertion (str): lettre de la colonne où l'insertion doit avoir lieu.
            - separator (str): le symbole délimitant le début du mot
            - piece_number (int): l'index du morceau à prendre (début : 0)
            - line_beggining (int) : ligne où débute la recherche.
        """
        sheet = self.file.writebook[sheet_name] 

        # Fill cells of the new columns 
        for i in range(line_beginning,sheet.max_row + 1): 
            if sheet.cell(i, column_read).value is not None: 
                sheet.cell(i, column_insertion).value = sheet.cell(i, column_read).value.split(separator)[piece_number]  
             
    @act_on_columns
    def column_for_prime_probe_congruence(self, sheet_name, columns_read, column_insertion, line_beginning=2):
        """
        Vous avez trois colonnes l'une contient des chaines de caractères particulières qui sont prime, probe, croix de fixation ...
          Les deux autres contiennent des chaines de la forme MOTnb_.jpg où MOT peut 
        être congruent, neutre, incongruent et nb est un nombre. Vous souhaitez insérer une colonne contenant soit rien, soit prime
        suivi du MOT de la deuxième colonne si la chaîne de la première colonne est prime, soit probe suivi du MOT de la troisième 
        colonne si la chaîne de la première colonne est probe.

        Inputs:
            - columns_read (list[str]): the three columns, the two lasts contains MOTnb_.jpg and the first contains prime, probe.
            - column_insertion (str): lettre de la colonne où l'insertion doit avoir lieu. 
            - line_beggining (int) : ligne où débute la recherche.
        """ 
        sheet = self.file.writebook[sheet_name] 

        for i in range(line_beginning, sheet.max_row + 1):

            # Adjonction de la chaine de first_column à MOT
            if sheet.cell(i, columns_read[0]).value in ["prime", "Prime"]:
                mot = re.sub(r'([A-Z-a-z]+)\d+_[A-Z-a-z].jpg', r'\1', sheet.cell(i, columns_read[1]).value)
                sheet.cell(i,column_insertion).value = sheet.cell(i, columns_read[0]).value + "_" + mot 

            elif sheet.cell(i, columns_read[0]).value in ["probe", "Probe"]:
                mot = re.sub(r'([A-Z-a-z]+)\d+_[A-Z-a-z].jpg', r'\1', sheet.cell(i, columns_read[2]).value)
                sheet.cell(i,column_insertion).value = sheet.cell(i, columns_read[0]).value + "_" + mot 

    @act_on_columns
    def give_names_of_maximum(self, sheet_name, column_list, column_insertion, line_beggining = 2):
        """
        Vous avez une liste de colonnes avec des chiffres, chaque colonne a un nom dans sa première cellule. 
        Cette fonction crée une colonne dans laquelle on entre pour chaque ligne le nom de la colonne ou des colonnes qui contient le max.

        Inputs : 
            - column_insertion : 
            - columnlist :
        """ 
        sheet = self.file.writebook[sheet_name] 
        sheet.cell(1, column_insertion).value = "Colonne de(s) maximum(s)"

        #dico qui à une colonne associe le nom de la colonne
        dico = {}
        for column in column_list:
            dico[column] = sheet.cell(1,column).value
 
        for line in range(line_beggining, sheet.max_row + 1):
            #pour une ligne donnée, on récupère le nom de la colonne associé aux maximum(s).
            maximum = -1
            chaine = ""
            for column in column_list:
                cellvalue = sheet.cell(line, column).value
                if cellvalue > maximum:
                    maximum = cellvalue
                    chaine = dico[column]
                elif cellvalue == maximum:
                    chaine += "_" + dico[column]
            sheet.cell(line, column_insertion).value = chaine 

    @act_on_columns 
    def column_transform_string_in_binary(self, sheet_name, column_read, column_write,*good_answers,line_beginning = 2):
        """
        Fonction qui prend une colonne de chaîne de caractères et qui renvoie une colonne de 0 ou de 1
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où mettre les 0 ou 1.

        Inputs :
                column_read : l'étiquette de la colonne de lecture des réponses.
                colum_write : l'étiquette de la colonne d'écriture des 0 et 1. Par défaut, une colonne est insérée à cette position.
                good_answers : une séquence d'un nombre quelconque de bonnes réponses qui valent 1pt. Chaque mot ne doit pas contenir d'espace ni au début ni à la fin.
                line_beggining: (optionnel par défaut égaux à 2) : ligne où débute l'application de la fonction. 

        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.

        Exemple d'utilisation : 
        
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.column_transform_string_in_binary('A','B','reponse1','reponse2') 

            #Bien mettre les réponses de good_answers entre ''. 
        """  
        sheet = self.file.writebook[sheet_name] 

        for i in range(line_beginning, sheet.max_row + 1):
            chaine_object = Str(sheet.cell(i,column_read).value)   
            bool = chaine_object.clean_string().transform_string_in_binary(*good_answers) 
            sheet.cell(i,column_write).value = bool

    @act_on_columns 
    def column_convert_in_minutes(self, sheet_name, column_read,column_write,line_beginning = 2):
        """
        Fonction qui prend une colonne de chaines de caractères de la forme "10 jours 5 heures" 
        ou "5 heures 10 min" ou "10 min 5s" ou "5s" et qui renvoie le temps en minutes.
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne à remplir.
        Input : column_read : l'étiquette de la colonne de lecture des réponses.
                colum_write : l'étiquette de la colonne d'écriture. 
                line_beggining: (optionnel par défaut égaux à 2) : ligne où débute l'application de la fonction. 

        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.
        
        Exemple d'utilisation : 
        
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.column_convert_in_minutes('A','B',line_beggining = 3) 

        """  
        sheet = self.file.writebook[sheet_name] 

        for i in range(line_beginning, sheet.max_row + 1):
            chaine_object = Str(sheet.cell(i,column_read).value) 
            if chaine_object.chaine != "None": 
                bool = chaine_object.clean_string().convert_time_in_minutes() 
                sheet.cell(i,column_write).value = bool

    @act_on_columns 
    def column_set_answer_in_group(self, sheet_name, column_read,column_write,groups_of_responses,line_beginning = 2):
        """
        Dans le cas où il y a des groupes de réponses, cette fonction qui prend une colonne de chaîne de caractères 
        et qui renvoie une colonne remplie de chaînes contenant pour chaque cellule le groupe associé.
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où écrire.

        Input : 
                column_read : l'étiquette de la colonne de lecture des réponses.
                colum_write : l'étiquette de la colonne d'écriture. 
                groups_of_response : dictionnary which keys are response groups and which values are a list of responses 
        associated to this group.
                line_beggining: (optionnel par défaut égaux à 2) : ligne où débute l'application de la fonction. 

        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.
        
        Exemple d'utilisation : 
        
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.column_set_answer_in_group('A', 'B', {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']} ,line_beggining = 3) 

        """ 
        sheet = self.file.writebook[sheet_name] 
        reversed_group_of_responses = Other.reverse_dico_for_set_answer_in_group(groups_of_responses)

        for i in range(line_beginning,sheet.max_row + 1): 
            chaine_object = Str(sheet.cell(i,column_read).value)  
            group = chaine_object.clean_string().set_answer_in_group(reversed_group_of_responses) 
            sheet.cell(i,column_write).value = group 


# class PathControler(FileControler):
#     def __init__(self, path):
#         """Input : path (object of the class Path)"""
#         self.path = path
    
    # DANS LA FONCTION CI-DESSOUS, IL NE RESTE QU A ECRIRE LENVOI DES FICHIERS PAR MAIL.
    #  def create_one_file_by_tab_and_send_by_mail(self, send = False, adressjson = "", objet = "", message = ""):
    #     """
    #     Vous souhaitez fabriquer un fichier par onglet. Chaque fichier aura le nom de l'onglet. 
    #     Vous souhaitez éventuellement envoyer chaque fichier à la personne associée.
    #     Attention, pour utiliser cette fonction, les onglets doivent être de la forme "prenom nom" sans caractère spéciaux. 

    #     Inputs : 
    #         send(optional boolean) : True si on veut envoyer le mail, False si on veut juste couper en fichiers.
    #         adressjson(str) : nom du fichier xlsx qui contient deux colonnes la première avec les noms des onglets, la seconde avec l'adresse mail. Ce fichier doit être mis dans le dossier fichier_xls. 
    #         objet(optional str) : Objet du message.
    #         message (optional str) : Contenu du message.
    #     """ 
    #     if adressjson != "":
    #         file = open(self.file.path + adressjson, 'r')
    #         mailinglist = json.load(file)
    #         file.close()

    #     start = time()

    #     for tab in self.file.sheets_name: 

    #         file_to_send = self.build_file_from_tab(tab)
    #         if send:
    #             if adressjson == "":
    #                 prenom = tab.split(" ")[0]
    #                 nom = tab.split(" ")[1]
    #                 self.envoi_mail(prenom + "." + nom + "@universite-paris-saclay.fr", file_to_send, "tony.fevrier62@gmail.com", "qkxqzhlvsgdssboh", objet, message)
    #             else: 
    #                 self.envoi_mail(mailinglist[tab], file_to_send, "tony.fevrier62@gmail.com", "qkxqzhlvsgdssboh", objet, message) 
    #         Other.display_running_infos('one_file_by_tab_sendmail', tab, self.file.sheets_name, start)

#     def apply_method_on_homononymous_files(self, filename, method_name, *args, **kwargs):
#         """ 
#         Vous avez plusieurs dossiers contenant un fichier ayant le même nom.
#         Fonction qui prend tous les fichiers d'un même nom et qui lui applique une même méthode.  

#         Inputs:
#             - filename (str)
#             - method_name (str): the name of the method to execute 
#             - *args, **kwargs : arguments of the method associated with method_name
#         """
#         start = time()

#         # Récupérer tous les dossiers d'un dossier  
#         for directory in self.path.directories:
#             file = File(filename, self.path.pathname + directory + '/')
#             controler = FileControler(file)
#             method = getattr(controler, method_name)
#             method(*args, **kwargs) 
#             Other.display_running_infos(method_name, directory, self.path.directories, start)

#     def apply_method_on_homononymous_sheets(self, filename, sheetname, method_name, *args, **kwargs):
#         """ 
#         Vous avez plusieurs dossiers contenant un fichier ayant le même nom.
#         Fonction qui prend tous les fichiers d'un même nom et qui lui applique une même méthode.  

#         Inputs:
#             - filename (str)
#             - method_name (str): the name of the method to execute 
#             - *args, **kwargs : arguments of the method associated with method_name
#         """
#         start = time()

#         # Récupérer tous les dossiers d'un dossier  
#         for directory in self.path.directories: 
#             file = File(filename, self.path.pathname + directory + '/')
#             controler = FileControler(file) 
#             method = getattr(controler, method_name)
#             method(sheetname, *args, **kwargs) 
#             Other.display_running_infos(method_name, directory, self.path.directories, start)
           
#     def gather_files_in_different_directories(self, name_file, name_sheet, values_only=False):
#         """
#         Vous avez plusieurs dossiers contenant un fichier ayant le même nom. Vous souhaitez créer un seul fichier regroupant 
#         toutes les lignes de ces fichiers.

#         Inputs:
#             - name_file(str)
#             - name_sheet(str)
#             - values_only(bool): to decide whether or not copying only the values and not formulas
#         """
#         # Récupérer tous les dossiers d'un dossier
#         directories = [f for f in os.listdir(self.path.pathname) if os.path.isdir(os.path.join(self.path.pathname, f))]

#         # Créer un nouveau fichier
#         new_file = openpyxl.Workbook() 
#         new_sheet = new_file.worksheets[0] 

#         start = time()

#         # Récupérer le fichier dans chacun des dossiers
#         for directory in directories: 
#             sheet_to_copy = File(name_file, self.path.pathname + directory + '/').writebook[name_sheet]

#             # Copier une fois la première ligne
#             if directory == directories[0]:
#                 self.copy_paste_line(sheet_to_copy, 1, new_sheet, 1, values_only=values_only)

#             # Copier son contenu à la suite du fichier
#             for line in range(2, sheet_to_copy.max_row + 1): 
#                 if line % 200 == 0:
#                     print(line, sheet_to_copy.max_row + 1)
#                 self.add_line_at_bottom(sheet_to_copy, line, new_sheet, values_only=values_only)

#             # save at the end of each directory not to use too much memory
#             new_file.save(self.path.pathname  + "gathered_" + name_file)
#             Other.display_running_infos('gather_files_in_different_directories', directory, directories, start)

#     def create_one_onglet_by_participant(self, name_file, onglet_from, column_read, first_line=2):
#         """
#         VERSION ALTERNATIVE A APPLYHOMOGENEOUSFILES DOC OBSOLETE
#         Fonction qui prend un onglet dont une colonne contient des chaînes de caractères comme par exemple un nom.
#         Chaque chaîne de caractères peut apparaître plusieurs fois dans cette colonne (exe : quand un participant répond plusieurs fois)
#         La fonction retourne un fichier contenant un onglet par chaîne de caractères.
#           Chaque onglet contient toutes les lignes correspondant à cette chaîne de caractères.

#         Input : 
#             name_file (str): name of the file to divide
#             onglet_from : onglet de référence.
#             column_read : l'étiquette de la colonne qui contient les chaînes de caractères.
#             first_line : ligne où commencer à parcourir.
#             last_line : ligne de fin de parcours 
 
#         Exemple d'utilisation : 
    
#             file = File('dataset.xlsx')
#             file.create_one_onglet_by_participant('onglet1', 'A') 
#         """ 
#         directories = [f for f in os.listdir(self.path.pathname) if os.path.isdir(os.path.join(self.path.pathname, f))]

#         # Créer un nouveau fichier
#         new_file = openpyxl.Workbook()  
#         onglets = new_file.sheetnames
#         column_read = column_index_from_string(column_read)  
#         start = time()

#         for directory in directories:
#             file = File(name_file, self.path.pathname + directory + '/')
#             sheet = file.writebook[onglet_from] 

#             # Create one tab by identifiant containing all its lines
#             for i in range(first_line, sheet.max_row + 1):
#                 onglet = str(sheet.cell(i,column_read).value)

#                 # Prepare a new tab
#                 if onglet not in onglets:
#                     new_file.create_sheet(onglet)
#                     self.copy_paste_line(sheet, 1,  new_file[onglet], 1)
#                     onglets.append(onglet) 

#                 self.add_line_at_bottom(sheet, i, new_file[onglet]) 
#             Other.display_running_infos('create_one_onglet_by_participant', directory, directories, start)
            
#         # Deletion of the first tab 
#         del new_file[new_file.sheetnames[0]]
#         new_file.save(self.path.pathname + f'divided_{name_file}')
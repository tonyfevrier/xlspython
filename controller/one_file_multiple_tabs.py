from utils.utils import GetIndex, TabsCopy, DisplayRunningInfos
from openpyxl.utils import column_index_from_string  
from time import time 


class MultipleSameTabController():

    def __init__(self, file_object, tab_controller, optional_names_of_file=None):
        """
        Attributs: 
            - file_object (object of class File) 
            - optional_names_of_file (OptionalNamesOfFile object)
            - first_line (optional int)  
            - tabs_copy (TabsCopy object): object to apply copy method from a tab to a new tab
            - display (DisplayRunningInfos object): to display the current state of the run
        """
        self.file_object = file_object
        self.tab_controller = tab_controller 
        self.optional_names_of_file = optional_names_of_file  
        self.display = DisplayRunningInfos() 

    def reinitialize_tab_controller(self, tab_name):
        self.tab_controller.tab = self.file_object.get_tab_by_name(tab_name)
        self.tab_controller.reinitialize_storing_attributes()

    def apply_method_on_some_tabs(self, method_name, *args, **kwargs):
        """ 
        Vous avez un fichier contenant plusieurs onglets et vous souhaitez appliquer une même méthode de la 
        classe Sheet sur une liste de ces onglets du fichier. On s'attend à ce que tous les onglets aient une structure identique.

        Inputs:
            - method_name (str): the name of the method to execute 
            - *args, **kwargs : arguments of the method associated with method_name
        """  
        self.display.start_time = time()
        for tab_name in self.optional_names_of_file.names_of_tabs_to_modify:    
            # Get the method from its name and apply it  
            self.reinitialize_tab_controller(tab_name)
            method = getattr(self.tab_controller, method_name)
            method(*args, **kwargs) 

            self._update_display_infos(method_name, tab_name, self.optional_names_of_file.names_of_tabs_to_modify) 
            self.display.display_running_infos() 

        self.file_object.save_file() 
    
    def _update_display_infos(self, method_name, current_running_part, list_of_running_parts):
        self.display.method_name = method_name
        self.display.current_running_part = current_running_part
        self.display.list_of_running_parts = list_of_running_parts


class OneFileMultipleTabsController(GetIndex):
    """
    Handle methods involving multiple tabs of a file.
    """
    def __init__(self, file_object, optional_names_of_file=None, first_line=2):
        """
        Attributs: 
            - file_object (object of class File) 
            - optional_names_of_file (OptionalNamesOfFile object)
            - first_line (optional int)  
            - tabs_copy (TabsCopy object): object to apply copy method from a tab to a new tab
            - display (DisplayRunningInfos object): to display the current state of the run
        """
        self.file_object = file_object
        self.optional_names_of_file = optional_names_of_file  
        self.first_line = first_line 
        self.tabs_copy = TabsCopy()
        self.display = DisplayRunningInfos() 

    def _update_display_infos(self, method_name, current_running_part, list_of_running_parts):
        self.display.method_name = method_name
        self.display.current_running_part = current_running_part
        self.display.list_of_running_parts = list_of_running_parts

    def extract_a_column_from_all_tabs(self):
        """
        Fonction qui récupère une colonne dans chaque onglet pour former une nouvelle feuille
        contenant toutes les colonnes. La première cellule de chaque colonne correspond alors 
        au nom de l'onglet. Attention, en l'état, il faut que tous les onglets aient la même structure.
        """ 
        self.optional_names_of_file.column_to_read = column_index_from_string(self.optional_names_of_file.column_to_read)
        new_tab = self.file_object.create_and_return_new_tab(f"gather_{self.optional_names_of_file.column_to_read}")
        self.tabs_copy._choose_the_tab_to_write_in(new_tab)
        self.optional_names_of_file.column_to_write = 1

        self.display.start_time = time() 
        for tab_name in self.file_object.sheets_name: 
            self.tabs_copy._choose_the_tab_to_read(self.file_object.get_tab_by_name(tab_name))
            self._copy_column_from_a_tab_in_the_next_new_tab_column(tab_name)

            self._update_display_infos('extract_column_from_all_sheets', tab_name, self.file_object.sheets_name)
            self.display.display_running_infos()

        self.file_object.save_file()  
        self.file_object.update_sheet_names()
    
    def _copy_column_from_a_tab_in_the_next_new_tab_column(self, tab_name): 
        self.tabs_copy.copy_paste_column(self.optional_names_of_file.column_to_read, self.optional_names_of_file.column_to_write) 
        self._choose_the_new_column_title(tab_name)  
        self.optional_names_of_file.column_to_write = self.tabs_copy.tab_to.max_column + 1

    def _choose_the_new_column_title(self, title):
        self.tabs_copy.tab_to.cell(1, self.tabs_copy.tab_to.max_column).value = title    

    def apply_columns_formula_on_all_tabs(self):
        """
        Fonction qui reproduit les formules d'une colonne ou plusieurs colonnes
          du premier onglet sur toutes les colonnes situées à la même position dans les 
          autres onglets.

        Input : 
            -column_list : int. les positions des colonnes où récupérer et coller
        """
        columns_int_list = self.get_list_of_columns_indexes(self.optional_names_of_file.columns_to_read)

        self.display.start_time = time()
        self.tabs_copy._choose_the_tab_to_read(self.file_object.get_tab_by_name(self.file_object.sheets_name[0]))

        # on applique les copies des formules dans tous les onglets sauf le premier duquel viennent ces formules
        for tab_name in self.file_object.sheets_name[1:]:
            self.tabs_copy._choose_the_tab_to_write_in(self.file_object.get_tab_by_name(tab_name))
            self.tabs_copy.copy_paste_multiple_columns(columns_int_list) 

            self._update_display_infos('apply_column_formula_on_all_sheets', tab_name, self.file_object.sheets_name[1:])
            self.display.display_running_infos()
            
        self.file_object.save_file() 

    def apply_cells_formula_on_all_tabs(self, *cells):
        """
        Fonction qui reproduit les formules d'une cellule ou plusieurs cellules
          du premier onglet sur toutes les cellules situées à la même position dans les 
          autres onglets.

        Input : 
            -cells : string. les positions des cellule où récupérer et coller 
        """
        cells_list = GetIndex.get_list_of_cells_coordinates(cells) 

        self.display.start_time = time()
        tab_to_read = self.file_object.get_tab_by_name(self.file_object.sheets_name[0])
        self.tabs_copy._choose_the_tab_to_read(tab_to_read)

        # on applique les copies de formules dans tous les onglets sauf le premier duquel proviennent les formules
        for tab_name in self.file_object.sheets_name[1:]: 
            self.tabs_copy._choose_the_tab_to_write_in(self.file_object.get_tab_by_name(tab_name))  
            self.tabs_copy.deep_copy_multiple_cells(cells_list)  
            
            self._update_display_infos('apply_cells_formula_on_all_sheets', tab_name, self.file_object.sheets_name[1:])
            self.display.display_running_infos()

        self.file_object.save_file() 

    def gather_groups_of_multiple_columns_in_tabs_of_two_columns_containing_tags_and_values(self, *lists_of_columns):
        """
        Vous avez des groupes de colonnes de valeurs avec une étiquette en première cellule. Pour chaque groupe, vous souhaitez former deux colonnes de valeurs : l'une qui contient
        les valeurs rassemblées en une colonne, l'autre, à sa gauche, qui indique l'étiquette de la colonne dans laquelle elle a été prise.

        Inputs : 
            - onglet (str) : nom de l'onglet d'où on importe les colonnes.
            - column_lists (list[list[str]]) : liste de groupes de colonnes. Chaque groupe est une liste de colonnes.
        """ 
        tab_to_read = self.file_object.get_tab_by_name(self.optional_names_of_file.name_of_tab_to_read)
        self.tabs_copy._choose_the_tab_to_read(tab_to_read)

        for list_of_columns in lists_of_columns: 
            self.optional_names_of_file.columns_to_read = list_of_columns
            tab_to = self._create_a_tab_for_a_list_of_columns()
            self.tabs_copy._choose_the_tab_to_write_in(tab_to) 
            self.copy_tags_and_values_of_a_list_of_columns()

        self.file_object.save_file() 
    
    def _create_a_tab_for_a_list_of_columns(self):
        string_of_columns = ''.join(self.optional_names_of_file.columns_to_read)
        tab_name = f"tab_column_gathered_{string_of_columns}"
        return self.file_object.writebook.create_sheet(tab_name) 

    def copy_tags_and_values_of_a_list_of_columns(self):
        for column in self.optional_names_of_file.columns_to_read: 
            self.tabs_copy.copy_tag_and_values_of_a_column_at_tab_bottom(column_index_from_string(column))         

    def merge_cells_on_all_tabs(self, merged_cells_range):
        """
        Fonction qui merge les mêmes cellules sur tous les onglets d'un fichier 
        """
        
        merged_cells_range.start_column = column_index_from_string(merged_cells_range.start_column)
        merged_cells_range.end_column = column_index_from_string(merged_cells_range.end_column)

        self.display.start_time = time()

        for tab_name in self.file_object.sheets_name: 
            self.tabs_copy._choose_the_tab_to_write_in(self.file_object.get_tab_by_name(tab_name)) 
            self.tabs_copy.tab_to.merge_cells(start_row=merged_cells_range.start_line, 
                                              start_column=merged_cells_range.start_column, 
                                              end_row=merged_cells_range.end_line, 
                                              end_column=merged_cells_range.end_column)
            self._update_display_infos('merge_cells_on_all_tabs', tab_name, self.file_object.sheets_name)
            self.display.display_running_infos()

        self.file_object.save_file() 

    def list_tabs_with_different_number_of_lines(self, number_of_lines):
        """
        Fonction qui prend un fichier et qui contrôle si tous les onglets ont un nombre de lignes égal à l'argument
        """
        list_of_tab_names = []
        for tab_name in self.file_object.sheets_name:
            list_of_tab_names = self.add_tab_to_list_if_different_number_of_lines(tab_name, list_of_tab_names, number_of_lines)
        return list_of_tab_names
    
    def add_tab_to_list_if_different_number_of_lines(self, tab_name, list_of_tab_names, number_of_lines):
        tab = self.file_object.get_tab_by_name(tab_name)
        if tab.max_row != number_of_lines:
            list_of_tab_names.append(tab_name)
        return list_of_tab_names
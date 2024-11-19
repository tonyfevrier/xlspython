import openpyxl
import os 
from utils.utils import UtilsForFile, Other, Str 
from pycel import ExcelCompiler  
from time import time
from datetime import datetime 
from copy import copy

class Path():
    def __init__(self, pathname = 'fichiers_xls/'):
        self.pathname = pathname
        self.directories = [f for f in os.listdir(self.pathname) if os.path.isdir(os.path.join(self.pathname, f))]
        

class File(UtilsForFile): 
    def __init__(self, name_file, path = 'fichiers_xls/', dataonly = False):
        """L'utilisateur sera invité à mettre son fichier xlxx dans un dossier nommé fichiers_xls
        """ 
        self.name_file = name_file  
        self.path = path
        self.dataonly = dataonly 
        self.writebook = openpyxl.load_workbook(self.path + self.name_file, data_only = dataonly)
        self.sheets_name = self.writebook.sheetnames 
        self.writebook_copy = None
        
    def create_excel_compiler(self):
        return ExcelCompiler(self.path + self.name_file) 
    
    def make_file_horodated_copy(self):
        self.writebook_copy = create_empty_workbook()
        self._copy_tabs_in_new_workbook()
        self._save_file()            
                    
    def _copy_tabs_in_new_workbook(self): 
        start = time()
        for tab_name in self.sheets_name:
            self.writebook_copy.create_sheet(tab_name)
            self._copy_old_file_tab_in_new_file_tab(tab_name)
            Other.display_running_infos('sauvegarde', tab_name, self.sheets_name, start)

    def _copy_old_file_tab_in_new_file_tab(self, tab_name):
        original_tab = self.writebook[tab_name]
        for i in range(1, original_tab.max_row + 1):
            for j in range(1, original_tab.max_column + 1): 
                self._copy_old_file_cell_in_new_file_cell(tab_name, Cell(i, j))

    def _copy_old_file_cell_in_new_file_cell(self, tab_name, cell):  
        new_tab = self.writebook_copy[tab_name]
        original_tab = self.writebook[tab_name]
        new_tab.cell(cell.i, cell.j).value = original_tab.cell(cell.i, cell.j).value  
        new_tab.cell(cell.i, cell.j).fill = copy(original_tab.cell(cell.i, cell.j).fill)
        new_tab.cell(cell.i, cell.j).font = copy(original_tab.cell(cell.i, cell.j).font)  

    def _save_file(self):
        name_file_no_extension = Str(self.name_file).del_extension() 
        self.writebook_copy.save(self.path  + name_file_no_extension + '_date_' + datetime.now().strftime("%Y-%m-%d_%Hh%M") + '.xlsx') 

            
class Cell():
    def __init__(self, i, j):
        self.i = i
        self.j = j


def create_empty_workbook():
    workbook = openpyxl.Workbook()
    del workbook[workbook.active.title]
    return workbook




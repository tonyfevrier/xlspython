import openpyxl
import os 
from utils.utils import UtilsForFile, Other, Str 
# from pycel import ExcelCompiler  
# from time import time
# from datetime import datetime 
# from copy import copy


""" def display_run():
    def wrapper(method, *args, **kwargs):
        start = time()
        method(*args, **kwargs)
        Other.display_running_infos('sauvegarde', tab_name, self.sheets_name, start)
    return wrapper """

class Path():
    def __init__(self, pathname = 'fichiers_xls/'):
        self.pathname = pathname
        self.directories = [f for f in os.listdir(self.pathname) if os.path.isdir(os.path.join(self.pathname, f))]
        

class File(): 
    def __init__(self, name_file, path = 'fichiers_xls/', dataonly = False):
        """
        Handle methods modifying a tab
        """ 
        self.name_file = name_file  
        self.path = path
        self.dataonly = dataonly 
        self.writebook = openpyxl.load_workbook(self.path + self.name_file, data_only = dataonly)
        self.sheets_name = self.writebook.sheetnames 
        
    def get_tab_by_name(self, tab_name):
        return self.writebook[tab_name]

class Tab():
    """Handle methods modifying a tab"""
    def __init__(self, tab_name): #file_object, tab_name):
        #self.file_object = file_object
        #self.tab = file_object.writebook[tab_name]  
        self.tab_name = tab_name

    # def _copy_old_file_tab_in_new_file_tab(self): 
    #     for i in range(1, self.tab.max_row + 1):
    #         for j in range(1, self.tab.max_column + 1): 
    #             new_writebook = self.file_object.writebook_copy
    #             Cell(self, i, j)._copy_old_file_cell_in_new_file_cell(new_writebook)


class Line():
    def __init__(self, tab, line_index):
        #self.tab_object = tab_object
        self.tab = tab
        self.line_index = line_index


class Column():
    def __init__(self, tab, letter):
        #self.tab_object = tab_object
        self.tab = tab
        self.letter = letter
            

class Cell():
    """Handle methods modifying a cell"""
    def __init__(self, i, j): 
        #self.tab_object = tab_object 
        self.i = i
        self.j = j

    # def _copy_old_file_cell_in_new_file_cell(self, new_writebook):   
    #     new_tab = new_writebook[self.tab_object.tab_name] 
    #     new_tab.cell(self.i, self.j).value = self.tab_object.tab.cell(self.i, self.j).value  
    #     new_tab.cell(self.i, self.j).fill = copy(self.tab_object.tab.cell(self.i, self.j).fill)
    #     new_tab.cell(self.i, self.j).font = copy(self.tab_object.tab.cell(self.i, self.j).font) 




# def create_empty_workbook():
#     workbook = openpyxl.Workbook()
#     del workbook[workbook.active.title]
#     return workbook






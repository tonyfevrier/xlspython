import openpyxl
import os   

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
        
    def create_and_return_new_tab(self, tab_name):
        return self.writebook.create_sheet(tab_name)

    def get_tab_by_name(self, tab_name):
        return self.writebook[tab_name]
    
    def get_cell_value_from_a_tab(self, tab, cell):
        return str(tab.cell(cell.line_index, cell.column_index).value)
    
    def update_sheet_names(self):
        self.sheets_name = self.writebook.sheetnames 

    def save_file(self):
        self.writebook.save(self.path + self.name_file) 


class Tab():
    """Handle methods modifying a tab"""
    def __init__(self, tab_name):   
        self.tab_name = tab_name 


# class Line():
#     def __init__(self, tab, line_index): 
#         self.tab = tab
#         self.line_index = line_index


# class Column():
#     def __init__(self, tab, letter): 
#         self.tab = tab
#         self.letter = letter
            

class Cell():
    """Handle methods modifying a cell"""
    def __init__(self, line_index, column_index):  
        self.line_index = line_index
        self.column_index = column_index
 






from time import time
from copy import deepcopy
from model.model_factorise import File
from utils.utils_factorise import DisplayRunningInfos, TabsCopy

class PathController():
    def __init__(self, path, name_file, controller):
        """Input : path (object of the class Path)"""
        self.path = path 
        self.name_file = name_file
        self.controller = controller
        self.display = DisplayRunningInfos() 
        self.tabs_copy = TabsCopy

    def _store_controller_attributes(self):
        return deepcopy(self.controller)
    
    def _reinitialize_controller_attributes(self, controller_copy):
        self.controller = deepcopy(controller_copy) 

    def apply_method_on_homononymous_files(self, method_name, *args, **kwargs):
        """ 
        Vous avez plusieurs dossiers contenant un fichier ayant le même nom. Vous leur appliquez la même méthode. 
        """
        self.display.start_time = time()
        controller_copy = self._store_controller_attributes()

        for directory in self.path.directories:
            self._reinitialize_controller_attributes(controller_copy)
            self.controller.file_object = File(self.name_file, self.path.pathname + directory + '/') 
            method = getattr(self.controller, method_name)
            method(*args, **kwargs) 

            self.display._update_display_infos(method_name, directory, self.path.directories) 
            self.display.display_running_infos()  

    def apply_method_on_homononymous_tabs(self, method_name, *args, **kwargs):
        """ 
        Vous avez plusieurs dossiers contenant un fichier ayant le même nom. Vous leur appliquez la même méthode. 
        """
        self.display.start_time = time() 
        controller_copy = self._store_controller_attributes() 

        for directory in self.path.directories: 
            self._reinitialize_controller_attributes(controller_copy)  
            self.controller.file_object = File(self.name_file, self.path.pathname + directory + '/') 
            self.controller.tab = self.controller.file_object.get_tab_by_name(self.controller.tab_name)
            method = getattr(self.controller, method_name)
            method(*args, **kwargs) 

            self.display._update_display_infos(method_name, directory, self.path.directories) 
            self.display.display_running_infos()           
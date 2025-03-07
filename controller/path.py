import json
import re

from time import time
from copy import deepcopy
from model.model import File, Mail
from utils.utils import DisplayRunningInfos, TabsCopy

class SeveralFoldersOneFileController():
    """Apply a method for all files having save name but stored in different directories of a given folder"""
    def __init__(self, path, name_file, controller, dataonly=False):
        """Input : path (object of the class Path)"""
        self.path = path 
        self.name_file = name_file
        self.controller = controller
        self.dataonly = dataonly
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
            self.controller.file_object = File(self.name_file, self.path.pathname + directory + '/', dataonly=self.dataonly) 
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
            self.controller.file_object = File(self.name_file, self.path.pathname + directory + '/', dataonly=self.dataonly) 
            self.controller.tab = self.controller.file_object.get_tab_by_name(self.controller.tab_name)
            method = getattr(self.controller, method_name)
            method(*args, **kwargs) 

            self.display._update_display_infos(method_name, directory, self.path.directories) 
            self.display.display_running_infos()     


class PathMailSender():
    """Apply a method on all files of a folder"""
    def __init__(self, path, mail, receiver_domain):
        self.path = path
        self.mail = mail
        self.receiver_domain = receiver_domain
        self.display = DisplayRunningInfos()

    def send_files_by_mail_using_same_domain(self):
        # Files must be named prenom nom.xlsx 
        # mail is an instance of Mail class    

        self.display.start_time = time() 

        for file_name in self.path.files: 
            self._send_file_corresponding_to(file_name)

            self.display._update_display_infos("send_files_by_mail_using_same_domain", file_name, self.path.files) 
            self.display.display_running_infos()  

    def _send_file_corresponding_to(self, file_name):
        name = re.sub(r'(.+).xlsx',r'\1', file_name) 
        prenom = name.split(" ")[0]
        nom = name.split(" ")[1]
        receiver_mail = prenom + "." + nom + self.receiver_domain
        self.mail.receiver_mail = receiver_mail
        self.mail.joint_file = self.path.pathname + file_name#name_file_to_send
        self.mail.send()

    def send_files_by_mail_using_(self, json_file):
        # json file must contain people's name as key and corresponding mail as value

        file = open(self.path.pathname + json_file, 'r')
        mailing_list = json.load(file)
        file.close()

        self.display.start_time = time() 

        for file_name in self.path.files: 
            self.mail.receiver_mail = mailing_list[file_name] 
            self.mail.joint_file = self.path.pathname + file_name
            self.mail.send()
            
            self.display._update_display_infos("send_files_by_mail_using_same_domain", file_name, self.path.files) 
            self.display.display_running_infos()  
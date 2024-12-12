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
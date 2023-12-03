from openpyxl.styles import PatternFill
import typer

class UtilsForFile():
    def copy_paste_line(self,onglet_from,row_from, onglet_to, row_to ):
        """
        Fonction qui prend une ligne de la feuille et qui la copie dans un autre onglet.

        Inputs : 
            - onglet_from : onglet d'où on copie
            - row_from : ligne de l'onglet d'origine.
            - onglet_to : onglet où coller.
            - row_to : la ligne où il faut coller dans l'onglet à modifier.

        Exemple d'utilisation : 
      
            file = File('dataset.xlsx')
            file.copy_paste_line('onglet1', 1, 'onglet2', 1)
        """

        for j in range(1, onglet_from.max_column + 1): 
            onglet_to.cell(row_to,j).value = onglet_from.cell(row_from, j).value 

    def copy_paste_column(self ,onglet_from ,column_from , onglet_to, column_to ):
        """
        Fonction qui prend une colonne de la feuille et qui la copie dans un autre onglet.
        """

        for i in range(1, onglet_from.max_row + 1): 
            onglet_to.cell(i,column_to).value = onglet_from.cell(i,column_from).value 

    def add_line_at_bottom(self, onglet_from, row_from, onglet_to):
        """
        Fonction qui copie une ligne spécifique de la feuille à la fin d'un autre onglet.

        Input : 
            - row_origin : ligne de l'onglet d'origine.
            - onglet : l'onglet à modifier où on copie la ligne.

        Exemple d'utilisation : 
     
            file = File('dataset.xlsx')
            file.copy_paste_line('onglet1', 1, 'onglet2')
        """ 
        self.copy_paste_line(onglet_from, row_from, onglet_to, onglet_to.max_row + 1)  

class UtilsForSheet():
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
    
    def color_line(self, color, row_number):
        """
        Fonction qui colore une ligne spécifique

        Input :
            - color : une couleur indiquée en haxadécimal par l'utilisateur.
            - row_number : le numéro de la ligne à colorer

        Exemple d'utilisation : 
    
            sheet = Sheet('dataset.xlsx','onglet1')
            sheet.color_line('#FF0000', 3) 
        
        """

        for j in range(1, self.sheet.max_column + 1):
            self.sheet.cell(row_number,j).fill = PatternFill(fill_type = 'solid', start_color = color)

    def create_dico_from_columns(self, column_keys:int, column_values:int, first_line, last_line):
        """
        Function returning a dictionnary whose keys are elements of a column
          if they are not empty and values are elements of an other column
        
        Inputs :
            column_keys : column whose elements are the keys of the dictionnary.
            column_values : same with values
            first_line : the line we begin to read the file

        Output : 
            dico : dictionary. 
        """

        dico = {}
        for i in range(first_line,last_line):
            key = self.sheet.cell(i,column_keys).value 
            
            print(key,self.sheet.max_row+1)
            if key != "":
                dico[key] = self.sheet.cell(i,column_values).value
        return dico
    
    def create_dico_to_store_multiple_answers_of_participants(self,column_read, column_store, line_beggining):
        """
        Fonction qui parcourt les identifiants d'un fichier et qui crée un dictionnaire contenant le nombre de réponses de chaque participant
        et les valeurs d'une même donnée lors des différentes réponses.

        Inputs : 
            - column_read (str) : la colonne avec les identifiants des participants.
            - column_store (str) : lettre de la colonne contenant la donnée qu'on veut stocker.
            - line_beggining (int) : ligne où débute la recherche. 
        """
        dico = {}
        for line in range(line_beggining,self.sheet.max_row + 1):
            identifier = self.sheet.cell(line,column_read).value
            value_to_store = self.sheet.cell(line,column_store).value 
            if identifier in dico.keys():
                dico[identifier][0] += 1
                dico[identifier][1].append(value_to_store)
            else: 
                dico[identifier] = [1, [value_to_store]]
        return dico
    
    def create_newsheet_storing_multiple_answers(self,storesheet,dico):
        """
        Fonction qui prend un dico issu de la fonction create_dico_to_store_multiple_answers_of_participants et qui crée une feuille 
        présentant les participants ayant répondu plusieurs fois et les valeurs de la donnée stockée lors des différentes réponses.

        Inputs:
            - storesheet (openpyxl sheet) : feuille de stockage.
            - dico (dict) : issu de create_dico_to_store_multiple_answers_of_participants.
        """
        storesheet.cell(1,1).value = 'Identifiers'
        
        for key,list in dico.items(): 
            if list[0] >= 2:
                firstline = storesheet.max_row + 1
                storesheet.cell(firstline, 1).value = key
                for i in range(len(list[1])):
                    storesheet.cell(firstline, i+2).value = list[1][i]



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
    
    def cut_string_in_parts(self, separator):
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
    
    def convert_time_in_minutes(self):
        """
        Function which takes a str of the form "10 jour 5 heures" and return a string giving the conversion in unity.

        Output : str
        """
        parts = self.cut_string_in_parts(" ")
        
        if parts[1] in ["jour","jours"]:
            duration = 24 * 60 * float(parts[0]) 
            if len(parts) > 2:
                if parts[3] in ['heure', 'heures']:
                    duration += float(parts[2]) * 60
                elif parts[3] == 'min':
                    duration +=  float(parts[2])
                else:
                    duration += round(float(parts[2])/60,2)
        elif parts[1] in ['heure', 'heures']:
            duration = float(parts[0]) * 60
            if len(parts) > 2:
                if parts[3] == "min":
                    duration += float(parts[2])
                else:
                    duration += round(float(parts[2])/60,2)
        elif parts[1] == "min":
            duration = float(parts[0])
            if len(parts) > 2:
                duration += round(float(parts[2])/60,2)
        else:
            duration = round(float(parts[0])/60,2)
    
        conversion = str(duration)
        return conversion
    
class UtilsForcommands():
    def askArgumentUntilNone(args, message):
        """
        Fonction qui permet de demander en ligne de commande un nombre d'arguments indéterminé à l'utilisateur. S'arrête quand l'utilisateur ne rentre rien.
        """
        if not args:
            args = []
            while True:
                user_input = typer.prompt(message, default="")
                if not user_input:
                    break
                args.append(user_input)
        return args
    
    def createDictByCmd(message1, message2):
        """
        Fonction qui demande à l'utilisateur de créer un dictionaire une clé, une valeur.
        """
        dictionary = {}
        while True: 
            key = typer.prompt(message1, default = "")
            if not key:
                break 
            value = typer.prompt(message2, default = "")
            dictionary[key] = value
        return dictionary

        
    @classmethod
    def createDictListValueByCmd(cls,message):
        """
        Fonction qui demande en ligne de commande à l'utilisateur de rentrer un par un les groupes de réponses et, pour chaque groupe, de rentrer
        les réponses qui lui sont associées.
        """ 
        answers = None
        dictionary = {} 
        while True:
            user_input = typer.prompt(message,default = "")
            if not user_input:
                break
            else: 
                values = cls.askArgumentUntilNone(answers,"") 
            dictionary[user_input] = values
        return dictionary
            
    def insertOrOverwrite(column_write):
        """
        Fonction qui propose à l'utilisateur d'insérer ou écraser les données de la colonne dans laquelle il écrit.
        """
        bool = typer.prompt(f"By default, we insert a new column {column_write} to write data. If you agree, press enter. If you want to overwrite the column {column_write}, write False" , default = True)
        return bool
    
class Other():
    def reverse_dico_for_set_answer_in_group(self,dictionary):
        """
        Function taking a dictionary of the form {'group1':['a','b'],'group2':['c','d','e']} and returning the dictionary
        {'a':'group1','b':'group1','c':'group2','d':'group2','e':'group2'}
        """
        reverse_dico = {}
        for key, value in dictionary.items():
            for reponse in value:
                reverse_dico[reponse] = key
        return reverse_dico

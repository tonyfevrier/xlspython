"""
Fonctions utiles pour travailler automatiquement sur des tableurs excel

* pour les fonctions qui seraient les références à lancer pour l'utilisateur.

Classe File qui prend un fichier et qui possède des méthodes.
    1) (A préciser : quel but) Fonction sur le modèle de "cherche" ou plutôt "recherche_chaine_et_retourne_ligne" qui recherche une donnée dans une colonne donnée et qui renvoie une autre donnée d'une autre colonne ainsi que la ligne.
    2) * Fonction qui parcourt une colonne C et qui crée (ou insère pour éviter l'écrasement de données) une nouvelle colonne à une position donnée, cette nouvelle colonne étant le résultat d'une fonction appliquée à la colonne C et passée en argument.
        a Fonction qui parcourt une colonne qui contient plusieurs types de réponses et qui crée une nouvelle colonne à une position donnée qui contient 1 ou 0. Pourrait prendre en argument deux listes de réponses associées par le prog à 0 ou 1. A mon avis vu la fonction de la ligne précédente, il suffit de créer une fonction qui transforme un str en 0 ou 1 et de l'appliquer à la précédente fonction.
        b Fonction style xlsparse de dataset2 qui parcourt une colonne qui contient une chaîne à séparer et qui écrit les morceaux séparés en insérant des colonnes (autant que le nb de morceaux de la chaîne) à partir d'une colonne fixée en argument.  
        c Avec la fonction globale, il ne resterait qu'à écrire une fonction spécifique décrivant une action sur la chaîne de chaque cellule de cette colonne (exe : les deux ci-dessous) voir d'autres.
        d Fonction qui sous conditions d'une colonne colore une case.
        e Fonction qui si il y a une couleur insère une colonne et y met qqch.
    3) Même chose qu'en ligne 8 mais cette fois en remplaçant la même colonne (juste appeler la fonction ligne 6 et bien choisir la position de la nouvelle colonne = à l'ancienne)
    4) *Fonction qui parcourt une colonne C et qui supprime une ligne si la cellule contient qqch.
    5) *Fonction qui parcourt plusieurs colonnes d'un fichier et qui crée une nouvelle colonne contenant des choses dépendant du contenu des cellules (même style qu'en ligne 6 mais avec plusieurs colonnes au départ) : on aurait aussi une fonction générique en argument.
        a Fonction gén 1 : si on a ça et ça, on met un 1 dans la nouvelle colonne.
        b Fonction gén 2 : on fait la somme, la moyenne de colonnes chiffrées.
    6) *Fonction ajout_colonne_autre_fichier(file1, file2,column): qui parcourt les mails ou un élément caractérisant les participants d'un fichier et ajoute une des caractéristiques dans un second fichier si les mails ou la caractéristique est présent dans ce fichier (les mails sont dans un ordre différent du fichier de départ). Il faut passer en arg les onglets et les colonnes de travail des deux fichiers. Idem peut sûrement se baser sur celle 2 lignes au-dessus
    6bis) *Améliorer la fonction précédente et qui fait copie non pas une mais plusieurs colonnes (l'idée est que si on doit copier plusieurs colonnes, on ne fasse pas plusieurs fois la recherche des mails dans le fichier d'arrivée car c'est coûteux).
    6ter) Même chose mais cette fois en créant une colonne (et pas en copiant).
    7) *Fonction ajout ligne_autre_fichier : qui fait comme ajout colonne.  
    8) *Fonction qui prend tous les fichiers d'un dossier et qui fait la même action sur chacun de ces fichiers.
    9) En combinant les deux précédentes fonctions, on peut créer un fichier de data à partir n fichiers individuels.
    10) Fonction qui regarde si une colonne contient ou non des choses : on pourra s'en servir afin d'éviter d'écraser des données déjà écrite.
    11) (- urgent) Fonction qui trie les lignes suivant un ou plusieurs critères avec des ordres de priorité suivant les critères. Par exe, pour le recrutement de l'institut, on veut trier les femmes en premier critère puis par handicap puis...
    12) (- urgent) Fonction qui filtre les lignes suivant un critère.
    13) (- urgent) Fonction qui copie des lignes dans un autre fichier. On pourra la combiner à 12) pour classer pour le recrutt Charpak.
    14) Fonction qui si qqch est écrit dans une case la colore en une couleur choisie par l'utilisateur. A décliner sur une colonne ou sur l'ensemble d'une feuille.

Tests : 
    pour 1)
    2a créer une fonction testant si de 
    
    Fonction test_correspondance_entre_deux_fichiers (à débugger) qui prend deux fichiers. Dans chaque fichier, on a des clés (ici un mail) mais l'ordre des lignes est différent.
    Ces deux fichiers sont censés avoir chacun une colonne avec les mêmes valeurs associées à chaque clé. On vérifie que c'est le cas.
    Créer un fonction test pour chacune de ces fonctions et un excel jouet court pour voir si le test passe.
    
Classe chaine :
    Fonction qui prend une str et qui la sépare en plusieurs chaines, la sparation étant donnée par un séparateur.
    Fonction qui enlève les guillemets ou un symbole qcq autour d'une chaine si ce symbole est là.

Pour la programmation par classe, la logique voudrait une classe File parent, une classe enfant onglet, puis une classe petit enfant colonne

Version ++ : on fait une interface graphique ou web permettant d'entrer un excel et faire ces opérations sans code.
"""
 
from xlutils.copy import copy 

import openpyxl 
from openpyxl.styles import PatternFill
from copy import copy
from datetime import date, datetime

class File(): 
    def __init__(self,name_file, path = 'fichiers_xls/'):
        """L'utilisateur sera invité à mettre son fichier xslx dans un dossier nommé fichiers_xls"""
        self.name_file = name_file  
        self.path = path
        self.writebook = openpyxl.load_workbook(self.path + self.name_file, data_only=True)
        self.sheets_name = self.writebook.sheetnames

    def sauvegarde(self):
        """
        Fonction qui crée une sauvegarde du fichier name_file et qui l'appelle name_file_numero où le numéro est le premier qui n'a pas été utilisé.
        """
        file_copy = openpyxl.Workbook()
        del file_copy[file_copy.active.title] #supprimer l'onglet créé

        for onglet in self.sheets_name:
            new_sheet = file_copy.create_sheet(onglet)
            initial_sheet = self.writebook[onglet] 

            for i in range(1,initial_sheet.max_row+1):
                for j in range(1,initial_sheet.max_column+1): 
                    new_sheet.cell(i,j).value = initial_sheet.cell(i,j).value  
                    new_sheet.cell(i,j).fill = copy(initial_sheet.cell(i,j).fill)
                    new_sheet.cell(i,j).font = copy(initial_sheet.cell(i,j).font) 
                    
        name_file_no_extension = Str(self.name_file).del_extension() 
        #file_copy.save(self.path  + name_file_no_extension + '_date_' + str(date.today()) + '.xlsx') 
        file_copy.save(self.path  + name_file_no_extension + '_date_' + datetime.now().strftime("%Y-%m-%d_%Hh%M") + '.xlsx') 

class Sheet(File): 
    def __init__(self, name_file, name_onglet,path = 'fichiers_xls/'): 
        super().__init__(name_file,path)
        self.name_onglet = name_onglet  
        self.sheet = self.writebook[self.name_onglet]
        del self.sheets_name

    def column_transform_string_in_binary(self,column_read,column_write,*good_answers,line_beginning = 2, line_end = 100, insert = True, security = True):
        """
        Fonction qui prend une colonne de str et qui renvoie une colonne de 0 ou de 1
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où mettre les 0 ou 1.
        Input : good_answers : une séquence d'un nb quelconque de bonnes réponses qui valent 1pt. Chaque mot ne doit pas contenir d'espace ni au début ni à la fin.
                column_read : la colonne de lecture des réponses.
                colum_write : la colonne d'écriture des 0 et 1. 
                line_beggining, line_end : intervalle de ligne dans lequel l'utilisateur veut appliquer sa transformation
        
        Output : rien sauf si la security est enclenchée et que l'on écrit dans une colonne déjà remplie.
        """

        if insert == False and security == True and self.column_security(column_write) == False:
            msg = "La colonne n'est pas vide. Si vous voulez vraiment y écrire, mettez security = False en argument."
            print(msg)
            return msg

        if insert == True:
            self.sheet.insert_cols(column_write)

        for i in range(line_beginning,line_end):
            chaine_object = Str(self.sheet.cell(i,column_read).value)  
            bool = chaine_object.clean_string().transform_string_in_binary(*good_answers) 
            self.sheet.cell(i,column_write).value = bool
 
        self.writebook.save(self.path + self.name_file)
    
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
        
    def column_color(self,column,chainecolor):
        """
        Fonction qui regarde pour une colonne donnée colore les cases contenant à certaines chaînes de caractères
        Input : 
            - column : le numéro de la colonne.
            - chainecolor : les str qui vont être colorés et les couleurs qui correspondent à écrire avec la syntaxe suivante {'vrai':'couleur1','autre':couleur2}. Attention,
                la couleur doit être entrée en hexadécimal et les chaînes de caractères ne doivent pas avoir d'espace au début ou à la fin.
        """
        
        for i in range(1,self.sheet.max_row + 1):
            cellule = self.sheet.cell(i,column) 

            if type(cellule.value) == str:
                key = Str(cellule.value).clean_string().chaine
            else: 
                key = cellule.value

            if key in chainecolor.keys():
                cellule.fill = PatternFill(fill_type = 'solid', start_color = chainecolor[key])

        self.writebook.save(self.path + self.name_file)

    def sheet_color(self,chainecolor): 
        """
        Fonction qui colore les cases contenant à certaines chaînes de caractères d'une feuille
        Input : 
            - column : le numéro de la colonne.
            - chainecolor : les str qui vont être colorés et les couleurs qui correspondent à écrire avec la syntaxe suivante {'vrai':'couleur1','autre':couleur2}. Attention,
                la couleur doit être entrée en hexadécimal et les chaînes de caractères ne doivent pas avoir d'espace au début ou à la fin.
        """

        for j in range(1, self.sheet.max_column + 1):
            self.column_color(j,chainecolor)
        
    


class Str():
    def __init__(self,chaine):
        self.chaine = chaine
        

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

    

"""
Déroulé et prochaines étapes :

    FAIT Imaginer la strcuture par classe 
    FAIT Fabriquer un excel jouet puis un micro test pour la fonction column transform string. 
    FAIT Tester l'ouverture de l'attribut sheet (bonne page) de la classe sheet. 
    FAIT Programmer la fonction et la tester : ajouter l'intervalle des lignes où l'opération a lieu.
    FAIT Factoriser : Nettoyer le fichier des commentaires inutiles
    FAIT : Factoriser: Voir les méthodes qui doivent renvoyer l'objet complet. 
    FAIT : Factoriser si c'est possible : notamment voir la page openclassroom sur la poo : normalement sheet devrait avoir un attribut readbook puisqu'elle hérite.
    FAIT : Factoriser : voir aussi comment utiliser args, kwargs.
    FAIT : Poo : voir s'il ne vaut pas mieux créer une classe Files avec deux noms : celui du fichier à lire et celui à écrire.
    FAIT : Factoriser : Certains arguments des méthode ne seraient-ils pas mieux comme attributs de classe?
    FAIT : Créer un repository git (j'aurais dû le faire bien avant).
    FAIT : Passer à openpyxl : modifier avec les nouvelles commandes.
    FAIT : Faire et retester une fonction sécurité qui empêche d'écrire dans une colonne contenant des choses. Pour cela ajouter dans les fonctions un paramètre security = True qui mis à False permettra d'écrire dans une colonne déjà remplie.
    FAIT : Ajouter dans la classe File une méthode permettant de créer une sauvegarde du fichier de départ 
    FAIT : Ecrire la fonction test_files_identical
    FAIT : Améliorer la fonction copy afin de conserver aussi le format des cellules, les couleurs de fond et de texte.
    FAIT : Voir aussi pour obtenir un nom plus pertinent pour le fichier copié. Mettre test_2023_04_25 pour avoir un historique des copies. Il faudrait alors changer ma fonction del_extension pour supprimer aussi la date si on sauve un fichier déjà daté.
    FAIT : Modifier mes classes de sorte que les modifications se fassent sur le même fichier (en ayant bien vérifié que la sauvegarde fonctionne avant).
    FAIT : Ajouter l'heure au nom du fichier sauvegardé.
    FAIT : Modifier ma fonction 2a avec un paramètre insert = True qui choisit si on insère ou non une colonne à la position column_write. Si on n'insère pas, le paramètre security permet alors d'éviter d'écraser.
    FAIT : Tester ma fonction dans les deux cas : insert = True ou False.    
    FAIT : Débug : comprendre pourquoi dans column_color il ne rentre jamais dans la condition.
    Débugger le test de sheet_color.
    Fonction 6 : imaginer un test avec un fichier d'arrivée déjà écrit à la main (avec les colonnes séparées).
    Programmer le test.
    Programmer la fonction.
    Voir si on ne peut pas faire une seule fonction pour 2a et 2b qui utilise en argument les ss fonctions transform_string_in_binary et ...


Bug intéressants : 
    -oubli de sauvegarder la feuille en fin de fonction : le prog ne fait alors rien.     
    -certaines str sortant d'excel ont des espaces insécables \xa0 différents des espaces réguliers. Python voit ainsi parfois des str qui semblent identiques différemment.
"""



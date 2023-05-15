"""
Fonctions utiles pour travailler automatiquement sur des tableurs excel

* pour les fonctions qui seraient les références à lancer pour l'utilisateur.

Classe File qui prend un fichier et qui possède des méthodes.
    1) (A préciser : quel but) Fonction sur le modèle de "cherche" ou plutôt "recherche_chaine_et_retourne_ligne" qui recherche une donnée dans une colonne donnée et qui renvoie une autre donnée d'une autre colonne ainsi que la ligne.
    2) * Fonction qui parcourt une colonne C et qui crée (ou insère pour éviter l'écrasement de données) une nouvelle colonne à une position donnée, cette nouvelle colonne étant le résultat d'une fonction appliquée à la colonne C et passée en argument.
        a Fonction qui parcourt une colonne qui contient plusieurs types de réponses et qui crée une nouvelle colonne à une position donnée qui contient 1 ou 0. Pourrait prendre en argument deux listes de réponses associées par le prog à 0 ou 1. A mon avis vu la fonction de la ligne précédente, il suffit de créer une fonction qui transforme un str en 0 ou 1 et de l'appliquer à la précédente fonction.
        b Fonction style xlsparse de dataset2 qui parcourt une colonne qui contient une chaîne à séparer et qui écrit les morceaux séparés dans des colonnes préalablement choisies.
        c Avec la fonction globale, il ne resterait qu'à écrire une fonction spécifique décrivant une action sur la chaîne de chaque cellule de cette colonne (exe : les deux ci-dessous) voir d'autres.
        d Fonction qui sous conditions d'une colonne colore une case.
        e Fonction qui si il y a une couleur insère une colonne et y met qqch.
    3) Même chose qu'en ligne 8 mais cette fois en remplaçant la même colonne (juste appeler la fonction ligne 6 et bien choisir la position de la nouvelle colonne = à l'ancienne)
    4) *Fonction qui parcourt une colonne C et qui supprime une ligne si la cellule contient qqch.
    5) *Fonction qui parcourt plusieurs colonnes d'un fichier et qui crée une nouvelle colonne contenant des choses dépendant du contenu des cellules (même style qu'en ligne 6 mais avec plusieurs colonnes au départ) : on aurait aussi une fonction générique en argument.
        a Fonction gén 1 : si on a ça et ça, on met un 1 dans la nouvelle colonne.
        b Fonction gén 2 : on fait la somme, la moyenne de colonnes chiffrées.
    6) *Fonction ajout_colonne_autre_fichier(file1, file2,column): qui parcourt les mails ou un élément caractérisant les participants d'un fichier et ajoute une des caractéristiques dans un second fichier si les mails ou la caractéristique est présent dans ce fichier. Il faut passer en arg les onglets et les colonnes de travail des deux fichiers. Idem peut sûrement se baser sur celle 2 lignes au-dessus
    7) *Fonction ajout ligne_autre_fichier : qui fait comme ajout colonne.  
    8) *Fonction qui prend tous les fichiers d'un dossier et qui fait la même action sur chacun de ces fichiers.
    9) En combinant les deux précédentes fonctions, on peut créer un fichier de data à partir n fichiers individuels.
    10) Fonction qui regarde si une colonne contient ou non des choses : on pourra s'en servir afin d'éviter d'écraser des données déjà écrite.
    
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

class File():
    def __init__(self,name_file,name_file_generated='test_generated.xls', path = 'fichiers_xls/'):
        """L'utilisateur sera invité à mettre son fichier dans un dossier nommé fichiers_xls"""
        self.name_file = name_file 
        self.name_file_generated = name_file_generated
        self.path = path
        self.writebook = openpyxl.load_workbook(self.path + self.name_file, data_only=True)
        self.sheets_name = self.writebook.sheetnames
    

class Sheet(File):

    def __init__(self, name_file, name_onglet, name_file_generated='test_generated.xlsx',path = 'fichiers_xls/'):
        super().__init__(name_file,name_file_generated,path)
        self.name_onglet = name_onglet  
        self.sheet = self.writebook[self.name_onglet]
        del self.sheets_name

    def column_transform_string_in_binary(self,column_read,column_write,*good_answers,line_beginning = 2, line_end = 100):
        """
        Fonction qui prend une colonne de str et qui renvoie une colonne de 0 ou de 1
        L'utilisateur doit indiquer un numéro de colonne de lecture et un numéro de colonne où mettre les 0 ou 1. Si les numéros de colonne sont identiques il renvoie un message d'erreur.
        Input : good_answers : une séquence d'un nb quelconque de bonnes réponses qui valent 1pt. Chaque mot ne doit pas contenir d'espace ni au début ni à la fin.
                column_read : la colonne de lecture des réponses.
                colum_write : la colonne d'écriture des 0 et 1. 
                line_beggining, line_end : intervalle de ligne dans lequel l'utilisateur veut appliquer sa transformation
        """
        for i in range(line_beginning,line_end):
            chaine_object = Str(self.sheet.cell(i,column_read).value)  
            bool = chaine_object.clean_string().transform_string_in_binary(*good_answers) 
            self.sheet.cell(i,column_write).value = bool

        self.writebook.save(self.path + self.name_file_generated)

    
    """
    def column_security(self,column):
        bool = True
        for i in range(self.writebook.nrows):
            if self.sheet_write.cell_value(i,column) != '':
                bool = False
                break
        return bool
    """


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
        return self


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
    Faire et retester une fonction sécurité qui empêche d'écrire dans une colonne contenant des choses.
    Ajouter dans la classe File une méthode permettant de créer une sauvegarde du fichier de départ.
    Factoriser : enlever les deux noms de fichiers (name file generated et name file) de sorte d'avoir création d'une copie mais manipulation d'un seul nom de fichier.
    Modifier mes classes de sorte que les modifications se fassent sur le même fichier (en ayant bien vérifié que la sauvegarde fonctionne avant).
"""



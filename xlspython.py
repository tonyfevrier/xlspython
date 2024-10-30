import typer
from typing import Optional, List, Tuple 
from model import File, Path, Str
from controller import PathControler, FileControler
from typing_extensions import Annotated
from utils.utils import UtilsForcommands as Ufc

app = typer.Typer()


def apply_sheet_method_on_a_file(file_name, sheets, method_name, *args, **kwargs):
        """
        Create a file object and apply a given method on some or all sheets of this file.

        Inputs : 
            - file_name(str)
            - sheets (list[str]): list of sheet names to apply the method on
            - method_name(str)
            - *args, **kwargs: arguments associated with the method
        """
        fileobject = File(file_name)
        controler = FileControler(fileobject)

        if sheets:
            controler.apply_method_on_some_sheets(sheets, method_name, *args, **kwargs)
        else:
            # when sheets is empty, the method applies on all sheets
            controler.apply_method_on_some_sheets(fileobject.sheets_name, method_name, *args, **kwargs)


# Path commands

@app.command()
def gatherfiles(directory : Annotated[str, typer.Option(prompt = 'Enter the name of the directory containing all directories')],
                file : Annotated[str, typer.Option(prompt = 'Enter the name of the file')],
                sheet : Annotated[str, typer.Option(prompt = 'Enter the name of the sheet to copy')],
                values : Annotated[bool, typer.Option(prompt = 'Do you want to copy only values (y) or to keep forumla (n)')],):
    
    """
    Fonction agissant sur un dossier. Fonction qui prend plusieurs fichiers de structures identiques et 
    qui crée un fichier contenant l'ensemble des lignes de ces fichiers.

    Commande : 

        Version guidée : python xlspython.py gatherfiles

        Version complète : python xlspython.py gatherfiles --directory name --file name.xlsx --sheet name
    
    """ 
    pathobject = Path(directory + '/')
    controler = PathControler(pathobject)
    controler.gather_files_in_different_directories(file, sheet, values_only=values)


@app.command()
def multidelcols(directory : Annotated[str, typer.Option(prompt = 'Enter the name of the directory containing all directories')],
                 file : Annotated[str, typer.Option(prompt = 'Enter the name of the file. This name must be present in each directory')],
                 sheet : Annotated[str, typer.Option(prompt = 'Enter the name of the sheet')],
                 columns : Annotated[str, typer.Option(prompt = "Enter the group of column you want to KEEP. Respect the form A-D,E,G,H-J,Z without introducing any space.")]):
    """
    Fonction agissant sur un dossier. Fonction qui prend plusieurs fichiers de structures identiques et qui ne garde qu'un ensemble
    de colonnes prédéfinies. 

    Commande : 

        Version guidée : python xlspython.py multidelcols 
    """  
    pathobject = Path(directory + '/') 
    controler = PathControler(pathobject)
    controler.apply_method_on_homononymous_sheets(file, sheet, 'delete_other_columns', columns)


# File commands

@app.command()
def filesave(file : Annotated[str, typer.Option(prompt = 'Enter the file you want to save ')]):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. 
    Fonction qui crée une sauvegarde du fichier entré (file) et qui l'appelle name_file_time où time est le moment d'enregistrement.

    Commande : 

        Version guidée : python xlspython.py filesave

        Version complète : python xlspython.py filesave --file name.xlsx
    
    """
    fileobject = File(file) 
    fileobject.sauvegarde()


@app.command()
def multipletabs(file : Annotated[str, typer.Option(prompt = 'Enter the name of the xlsx file(s) you want to cut')],
                 sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name ')],
                 colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing strings ')],
                 newfilepath : Annotated[Optional[str], typer.Option(prompt = 'If you want to divide a single file in tabs, press enter, otherwise your files must be included in folders themselves included in a bigger folder whose name must be written now.')] = '',
                 line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un fichier ou sur un même fichier dans plusieurs dossiers différents. Si vous souhaitez l'utiliser sur un
     seul fichier, pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez un ou plusieurs fichiers xlsx dont une colonne (colread) contient des participants qui ont pu répondre plusieurs fois à un questionnaire. 
    Vous souhaitez créer un onglet par participant avec toutes les lignes qui correspondent. 

    Commande : 

        Version guidée : python xlspython.py multipletabs

        Version complète : python xlspython.py multipletabs --file name.xlsx --sheet nametab --colread columnletter --newfilepath name --line linenumber
    
    """
    # Apply command to same name files contained in folders
    if newfilepath:
        pathobject = Path(newfilepath + '/')
        controler = PathControler(pathobject)
        #controler.create_one_onglet_by_participant(file, sheet, colread, first_line=line)
        controler.apply_method_on_homononymous_files(file, 'create_one_onglet_by_participant', sheet, colread, f'divided_{file}', newfilepath + '/', first_line=line)
        path = pathobject.pathname

    # Apply command to a single file
    else:
        fileobject = File(file)
        controler = FileControler(fileobject)
        controler.create_one_onglet_by_participant(sheet, colread, f'divided_{file}', 'fichiers_xls/', first_line=line)
        path = fileobject.path
    
    # Eventually verify if each tab has the same number of lines
    check = typer.prompt('Do you want to check if all tabs have the same number of lines? If yes, write the number of lines else press enter', default="")
    if check:
        wrong_tabs = FileControler(File(f'divided_{file}', path)).check_linenumber_of_tabs(int(check))
        if wrong_tabs:
            print('The following tabs have a different number of lines :' + ",".join(wrong_tabs))


@app.command()
def extractcolsheets(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                     colread : Annotated[str, typer.Option(prompt = 'Enter the column letter ')]):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui récupère une même colonne (colread) dans chaque onglet pour former une nouvelle feuille contenant toutes les colonnes.
    La première cellule de chaque colonne correspond alors au nom de l'onglet.

    Commande : 

        Version guidée : python xlspython.py extractcolsheets

        Version complète : python xlspython.py extractcolsheets --file name.xlsx --colread columnletter
    
    """
    fileobject = File(file) 
    controler = FileControler(fileobject)
    controler.extract_column_from_all_sheets(colread)

@app.command()
def cutsendmail(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                sendmail : Annotated[Optional[str], typer.Option(prompt = '(Optional) Do you want to send files by mail? Press y (yes) or enter (no)')] = 'n'):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez fabriquer un fichier par onglet. Chaque fichier aura le nom de l'onglet. Vous souhaitez éventuellement envoyer chaque fichier à la personne associée.
    Attention, pour utiliser cette fonction, les onglets doivent être de la forme "prenom nom" sans caractère spéciaux. 

    Commande : 

        Version guidée : python xlspython.py cutsendmail

        Version complète : python xlspython.py cutsendmail 
    """
    fileobject = File(file, dataonly=True)
    controler = FileControler(fileobject)


    if sendmail ==  'n':
        controler.one_file_by_tab_sendmail()
    else:
        objet = typer.prompt('Please enter the object of your email',default="")
        message = typer.prompt('Please enter the message of your email',default="") 
        jsonfile = typer.prompt('Please enter the name of the json file containing mail adresses. If you want to send to the mail paris-saclay, just press enter',default="") 
        controler.one_file_by_tab_sendmail(send = True, adressjson = jsonfile, objet = objet, message = message)


@app.command()
def gathercolumn(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                 sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                 columnlists : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez des groupes de colonnes de valeurs avec une étiquette en première cellule. Pour chaque groupe, vous souhaitez former deux colonnes de valeurs : l'une qui contient
        les valeurs rassemblées en une colonne, l'autre, à sa gauche, qui indique l'étiquette de la colonne dans laquelle elle a été prise.
    
    Commande : 

        Version guidée : python xlspython.py gathercolumn

        Version complète : python xlspython.py gathercolumn --file nom --sheet onglet --columnlists A-D,E,G,H-J,Z
    """
 
    group = Ufc.askArgumentUntilNone(columnlists, "Enter a group of column of the form A-D,E,G,H-J,Z")
    column_lists = Str.listFromColumnsStrings(*group) 

    fileobject = File(file)
    controler = FileControler(fileobject)
    controler.gather_columns_in_one(sheet,*column_lists)


@app.command()
def extractcellsheets(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                 cells : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez un fichier avec des onglets de structure identique avec un onglet par participant. Vous souhaitez
    récupérer des cellules identiques dans tous les onglets et créer un onglet avec une ligne par participant,
    qui contient les valeurs de ces cellules. Fonction analogue à gather_multiple_answers mais ne portant pas sur une
    seule feuille.

    Commande :

        Version guidée: python xlspython.py extractcellsheets
    """
    
    cells = Ufc.askArgumentUntilNone(cells, "Entrez une cellule que vous souhaitez copier : ")
    fileobject = File(file)
    controler = FileControler(fileobject)
    controler.extract_cells_from_all_sheets(*cells)


@app.command()
def stringinbinary(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   sheets : Annotated[Optional[List[str]], typer.Option()] = None,
                   answers : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un ou plusieurs onglets d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Cette fonction lit les cellules d'une colonne (colread) et crée une nouvelle colonne (colwrite) contenant 1 si la valeur de la cellule est dans les bonnes réponses (answers)
    0 sinon.

    Commande : 

        Version guidée : python xlspython.py stringinbinary

        Version complète : python xlspython.py stringinbinary --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --answers chaine1 --answers chaine2
    
    """ 
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    answers = Ufc.askArgumentUntilNone(answers,"Enter one good answer and then press enter. Press directly enter if you have entered all the good answers")
     
    apply_sheet_method_on_a_file(file, sheets, 'column_transform_string_in_binary', colread, colwrite, *answers)


# Créer un fichier test pour tester cette fonction.
@app.command()
def cpcolumnonsheets(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                     column : Annotated[List[str], typer.Option()] = None):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui reproduit les formules d'une ou plusieurs colonnes (column) du premier onglet sur toutes les colonnes situées à la même position dans les 
          autres onglets.
    
    Commande : 

        Version guidée : python xlspython.py cpcolumnonsheets

        Version complète : python xlspython.py cpcolumnonsheets --file name.xlsx --column columnletter
    
    """
    columns = Ufc.askArgumentUntilNone(column,'Enter one column whose you want to reproduce the formula')

    fileobject = File(file, dataonly = False)
    controler = FileControler(fileobject)
    controler.apply_column_formula_on_all_sheets(*columns) 

@app.command()
def cpcellonsheets(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   cell : Annotated[List[str], typer.Option()] = None):
    
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui reproduit les formules d'une ou plusieurs cellules (cell) du premier onglet sur toutes les cellules situées à la même position dans les 
          autres onglets.
    
    Commande : 

        Version guidée : python xlspython.py cpcellonsheets

        Version complète : python xlspython.py cpcellonsheets --file name.xlsx --cell C5 --cell C17
    
    """

    cells = Ufc.askArgumentUntilNone(cell,'Enter one cell whose you want to reproduce the formula')

    fileobject = File(file)
    controler = FileControler(fileobject)
    controler.apply_cells_formula_on_all_sheets(*cells) 
    
@app.command()
def mergecells(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
               start_column : Annotated[str, typer.Option(prompt = 'Enter the first column of  cells to merge :')],
               end_column : Annotated[str, typer.Option(prompt = 'Enter the last column of cells to merge :')],
               start_row : Annotated[str, typer.Option(prompt = 'Enter the first row of  cells to merge :')],
               end_row : Annotated[str, typer.Option(prompt = 'Enter the last row of cells to merge :')]):
    
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui fusionne les cellules de start_column à end_column et de start_row à end_row.
    
    Commande : 

        Version guidée : python xlspython.py mergecells

        Version complète : python xlspython.py mergecells --file name.xlsx --start_column columnletter --end_column columnletter --start_row rowindex --end_column rowindex 
    
    """
    
    fileobject = File(file)
    controler = FileControler(fileobject)
    controler.merge_cells_on_all_tabs(start_column,end_column,start_row,end_row)

 
@app.command()
def convertminutes(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   sheets : Annotated[Optional[List[str]], typer.Option()] = None,
                   line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous avez une colonne (colread) contenant des temps de la forme xx jours 5 heures 10 min 5 s.
     Vous souhaitez convertir dans une colonne (colwrite) les temps en minutes.

    Commande : 

        Version guidée : python xlspython.py convertminutes

        Version complète : python xlspython.py convertminutes --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --line linenumber
    
    """
    sheets = Ufc.askArgumentUntilNone(sheets, "If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    apply_sheet_method_on_a_file(file, sheets, 'column_convert_in_minutes', colread, colwrite, line_beginning=line)


@app.command()
def groupofanswers(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   sheets : Annotated[Optional[List[str]], typer.Option()] = None,
                   line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous avez une colonne (colread) contenant des réponses. Chacune de ses réponses appartient à un groupe.
    Vous souhaitez afficher dans une colonne (colwrite) le groupe d'appartenance de la réponse. On vous demandera d'entrer des noms de groupes et dans la foulée, les réponses qui appartiennent au groupe.

    Commande : 

        Version guidée : python xlspython.py groupofanswers

        Version complète : python xlspython.py groupofanswers --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --line linenumber
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")

    #Creation of the groups of answers dictionary 
    groups_of_responses = Ufc.createDictListValueByCmd("Enter the name of one group of answers", "")

    apply_sheet_method_on_a_file(file, sheets, 'column_set_answer_in_group', colread, colwrite, groups_of_responses, line_beginning=line)
    

@app.command()
def colorcasescolumn(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                     column : Annotated[str, typer.Option(prompt = 'Enter the column')],
                     sheets : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir une colonne (column) et colorer certaines chaînes de caractères dans cette colonne. Ces chaînes et les couleurs associées vous seront demandées
    durant l'exécution de la commande.

    Commande : 

        Version guidée : python xlspython.py colorcasescolumn 

        Version complète : python xlspython.py colorcasescolumn --file name.xlsx --sheet nametab --column columnletter
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")

    #Creation of the dictionary with the strings to color and their color
    color = Ufc.createDictByCmd("Please enter a string which will be colored", "Please enter the color in hexadecimal type")
    
    apply_sheet_method_on_a_file(file, sheets, 'color_special_cases_in_column', column, color)
    

@app.command()
def colorcasestab(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                  sheets : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir un onglet (sheet) et colorer certaines chaînes de caractères. Ces chaînes et les couleurs associées vous seront demandées
    durant l'exécution de la commande.

    Commande : 

        Version guidée : python xlspython.py colorcasestab 

        Version complète : python xlspython.py colorcasestab --file name.xlsx --sheet nametab
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")

    #Creation of the dictionary with the strings to color and their color
    color = Ufc.createDictByCmd("Please enter a string which will be colored", "Please enter the color in hexadecimal type")

    apply_sheet_method_on_a_file(file, sheets, 'color_special_cases_in_sheet', color)


@app.command()
def addcolumn(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file in which you want to write ')],
            sheet : Annotated[str, typer.Option(prompt = 'Enter the corresponding sheet name')],
            colread : Annotated[str, typer.Option(prompt = 'Enter the column of this sheet containing the identifiers')],
            colwrite : Annotated[str, typer.Option(prompt = 'Enter the column from which you want to write')],
            file2 : Annotated[str, typer.Option(prompt = 'Enter the xlsx file from which you import data ')],
            sheet2 : Annotated[str, typer.Option(prompt = 'Enter the corresponding sheet name')],
            colread2 : Annotated[str, typer.Option(prompt = 'Enter the column of this sheet containing the identifiers')],
            colimport : Annotated[Optional[List[str]], typer.Option()] = None):                    
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
      Vous souhaitez importer des colonnes (colimport) d'un fichier (file2, sheet2) dans un autre fichier (file,sheet).
    La mise en correspondance entre les deux onglets s'effectue via une colonne d'identifiants dans chaque fichier (colread, colread2). L'important s'effectue 
    dans file à partir de la colonne colwrite.    

    Commande : 

        Version guidée : python xlspython.py addcolumn 

        Version complète : python xlspython.py addcolumn --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --file2 name.xlsx --sheet2 nametab --colread2 columnletter --colimport col1 --colimport col2
    
    """ 
    colimport = Ufc.askArgumentUntilNone(colimport,"Enter one column to import and then press enter. Press directly enter if you have entered all the columns to import.")
    fileobject = File(file)
    controler = FileControler(fileobject)

    controler.add_column_in_sheet_differently_sorted(sheet,colread,colwrite,[file2,sheet2,colread2,colimport])

@app.command()
def colorlines(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
               color : Annotated[str, typer.Option(prompt = 'Enter the color in a hexadecimal format')],
               sheets : Annotated[Optional[List[str]], typer.Option()] = None,
               strings : Annotated[Optional[List[str]], typer.Option()] = None):                    

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir un onglet (sheet) et colorer (color) les lignes contenant certaines chaînes de caractères (strings).

    Commande : 

        Version guidée : python xlspython.py colorlines 

        Version complète : python xlspython.py colorlines --file name.xlsx --sheet nametab --color colorinhexadecimal --strings chaine1 --strings chaine2
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    strings = Ufc.askArgumentUntilNone(strings,"Enter one string leading to the coloration of the line and then press enter. Press directly enter if you have entered all the strings.")
    
    apply_sheet_method_on_a_file(file, sheets, 'color_lines_containing_chaines', color, *strings)
    

@app.command()
def cutstring(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
              colcut : Annotated[str, typer.Option(prompt = 'Enter the column containing strings to cut')],
              colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
              sheets : Annotated[Optional[List[str]], typer.Option()] = None,
              separator : Annotated[Optional[str], typer.Option(prompt = '(Optional) Enter the separator or press enter')] = ','):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. 
    Une colonne (colcut) contient des chaînes de caractères séparées par un symbole (separator). Vous souhaitez les couper en morceaux 
    et créer des colonnes (à partir de colwrite) pour chacun de ces morceaux.

    Commande : 

        Version guidée : python xlspython.py cutstring 

        Version complète : python xlspython.py cutstring --file name.xlsx --sheet nametab --colcut columnletter --colwrite columnletter --separator symbol
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    apply_sheet_method_on_a_file(file, sheets, 'column_cut_string_in_parts', colcut,colwrite,separator)
    

@app.command()
def deletecols(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
               columns : Annotated[str, typer.Option(prompt = 'Enter a group of column to delete of the form A-D,E,G,H-J,Z')],
               sheets : Annotated[Optional[List[str]], typer.Option()] = None):

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Prend une séquence de colonnes et les supprime.

    Commande : 

        Version guidée : python xlspython.py deletecols  

        Version complète : python xlspython.py deletecols --file name.xlsx --sheet nametab --columns A --columns D 
    
    """ 
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    apply_sheet_method_on_a_file(file, sheets, 'delete_columns', columns)


@app.command()
def keepcols(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
               columns : Annotated[str, typer.Option(prompt = 'Enter a group of column to delete of the form A-D,E,G,H-J,Z')],
               sheets : Annotated[Optional[List[str]], typer.Option()] = None):

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous entrez les colonnes que vous souhaitez garder. Toutes les autres seront supprimées.

    Commande : 

        Version guidée : python xlspython.py deletecols  

        Version complète : python xlspython.py deletecols --file name.xlsx --sheet nametab --columns A --columns D 
    
    """ 
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    apply_sheet_method_on_a_file(file, sheets, 'delete_other_columns', columns)


@app.command()
def deletelinesstr(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                sheets : Annotated[Optional[List[str]], typer.Option()] = None,
                strings : Annotated[Optional[List[str]], typer.Option()] = None):                    
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir une colonne (colread) et si une chaîne (strings) apparaît dans cette colonne, supprimer la ligne associée.

    Commande : 

        Version guidée : python xlspython.py deletelines 

        Version complète : python xlspython.py deletelines --file name.xlsx --sheet nametab --colread columnletter --strings chaine1 --strings chaine2
    
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    strings = Ufc.askArgumentUntilNone(strings,"Enter one string leading to the delation of the line and then press enter. Press directly enter if you have entered all the strings.")

    apply_sheet_method_on_a_file(file, sheets, 'delete_lines_containing_str', colread,*strings) 


@app.command()
def deletetwins(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                sheets : Annotated[Optional[List[str]], typer.Option()] = None,
                line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
      Certains participants à un questionnaire répondent plusieurs fois. Vous souhaitez parcourir une colonne (colread)
     qui les identifie et ne garder que leur dernière réponse à ce questionnaire.

    Commande : 

        Version guidée : python xlspython.py deletetwins 

        Version complète : python xlspython.py deletetwins --file name.xlsx --sheet nametab --colread columnletter --line linenumber
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")

    apply_sheet_method_on_a_file(file, sheets, 'delete_doublons', colread, line_beginning=line)  


@app.command()
def columnbyqcmanswer(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                     colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                     colwrite : Annotated[str, typer.Option(prompt = 'Enter the column from which you want to write')], 
                     sheets : Annotated[Optional[List[str]], typer.Option()] = None,
                     answers : Annotated[Optional[List[str]], typer.Option()] = None,
                     liste : Annotated[Tuple[str, str], typer.Option(prompt = 'Enter what you want to write in the cells or press enter')] = ('oui', 'non'),
                     line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
                   

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. 
    Une colonne (colread) contient toutes les réponses d'un participant à une question de QCM. Vous souhaitez créer autant de colonnes que
    de réponses (answers) à la question et mettre dans chaque colonne (à partir de colwrite) si les participants l'ont coché ou non (list).

    Commande : 

        Version guidée : python xlspython.py columnbyqcmanswer 

        Version complète : python xlspython.py columnbyqcmanswer --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --answers chaine1 --answers chaine2 --list oui non
    
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    answers = Ufc.askArgumentUntilNone(answers,"Enter one QCM answer and then press enter. Press directly enter if you have entered all the answers.")

    apply_sheet_method_on_a_file(file, sheets, 'create_one_column_by_QCM_answer', colread, colwrite, liste, *answers, line_beggining=line)  


@app.command()
def gathermultianswers(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                       sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                       colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                       colstore : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the data to store ')],
                       line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
      Certains participants à un questionnaire répondent plusieurs fois. Vous souhaitez parcourir une colonne (colread)
     qui les identifie et créer, dans un autre onglet, une ligne par participant ayant répondu plusieurs fois. Cette ligne contient les différentes
     réponses de ce participant contenues dans une colonne (colstore) donnée.

    Commande : 

        Version guidée : python xlspython.py gathermultianswers 

        Version complète : python xlspython.py gathermultianswers --file name.xlsx --sheet nametab --colread columnletter --colstore columnletter --line linenumber
    
    """
    fileobject = File(file)
    controler = FileControler(fileobject)
    controler.gather_multiple_answers(sheet,colread,colstore,line_beggining=line)

@app.command()
def maxnames(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
             colstore : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the data to store ')],
             sheets : Annotated[Optional[List[str]], typer.Option()] = None,
             columnlist : Annotated[Optional[List[str]], typer.Option()] = None,
             line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez une liste de colonnes avec des chiffres, chaque colonne a un nom dans sa première cellule.
      Cette fonction crée une colonne dans laquelle on entre pour chaque ligne le nom de la colonne ou des colonnes qui contient le max.

    Commande : 

        Version guidée : python xlspython.py maxnames 

        Version complète : python xlspython.py maxnames --file name.xlsx --sheet nametab --colstore columnletter --columnlist A --columnlist C
    
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    columnlist = Ufc.askArgumentUntilNone(columnlist, "Enter the letter of a column you want to read")

    apply_sheet_method_on_a_file(file, sheets, 'give_names_of_maximum', columnlist, colstore, line_beggining=line)  


@app.command()
def colcongruent(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                firstcol : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the word (example: prime, probe)')],
                secondcol : Annotated[str, typer.Option(prompt = 'Enter the column letter corresponding to prime')],
                thirdcol : Annotated[str, typer.Option(prompt = 'Enter the column letter corresponding to probe')],
                colwrite : Annotated[str, typer.Option(prompt = 'Enter the column letter in which you want to write')],
                sheets : Annotated[Optional[List[str]], typer.Option()] = None):
    
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    On a deux colonnes, une contenant des mots, l'autre contenant une chaine de la forme motnb_o.jpg (congruent1_o.jpg), on veut créer
    une colonne qui contient le mot de la première colonne + mot (congruent) si le mot de la première colonne fait partie d'une liste de mots prédéfinie.

    Commande : 

        Version guidée : python xlspython.py colcongruent
 
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    apply_sheet_method_on_a_file(file, sheets, 'column_for_prime_probe_congruence', [firstcol, secondcol, thirdcol], colwrite)  


@app.command()
def colgetpart(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                colread : Annotated[str, typer.Option(prompt = 'Enter the column letter to read ')],
                colwrite : Annotated[str, typer.Option(prompt = 'Enter the column letter where to write')],
                separator : Annotated[str, typer.Option(prompt = 'Enter the separator')],
                piece : Annotated[int, typer.Option(prompt = 'Enter the number of the part you want to get. For example, tap 1 to get the begin of the word')],
                sheets : Annotated[Optional[List[str]], typer.Option()] = None,
                line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez une colonne qui contient une chaîne qui contient un séparateur dont vous voulez prendre une partie. 
    Ce mot est inséré dans une nouvelle colonne. 

    Commande : 

        Version guidée : python xlspython.py colgetbegin
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    apply_sheet_method_on_a_file(file, sheets, 'column_get_part_of_str', colread, colwrite, separator, piece - 1, line_beginning=line)  


@app.command()
def maptwocols(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
               colread : Annotated[str, typer.Option(prompt = 'Enter the column letters to read separated by a comma (C,E)')],
               colwrite : Annotated[str, typer.Option(prompt = 'Enter the column letter where to write')],
               sheets : Annotated[Optional[List[str]], typer.Option()] = None,
               line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez deux colonnes de lecture, suivant ce qui est écrit sur une ligne, vous voulez ou non insérer quelque chose 
    dans une nouvelle colonne.

    Commande : 

        Version guidée : python xlspython.py maptwocols
    """
    sheets = Ufc.askArgumentUntilNone(sheets,"If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter")
    mapping = Ufc.createDictListValueByCmd("Enter a value you want to put in the new column", "Enter the two strings which should lead to this new value. You must enter it with the same order as the order you entered the columns to read")
    apply_sheet_method_on_a_file(file, sheets, 'map_two_columns_to_a_third_column', colread.split(","), colwrite, mapping, line_beginning=line)  


if __name__ == "__main__":
    app()
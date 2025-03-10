import typer
from typing import Optional, List
from model.model import File, Path, FileOptions, TabOptions, MergedCellsRange, Mail
from controller.path import SeveralFoldersOneFileController, PathMailSender
from controller.one_file_multiple_tabs import EvenTabsController, OneTabCreatedController, MultipleSameTabController
from controller.one_file_one_tab import ColorTabController, InsertController, DeleteController
from controller.two_files import TwoFilesController, OneFileCreatedController
from typing_extensions import Annotated 
from utils.prompts import *
from utils.utils import InputStore, MapStore, String

app = typer.Typer()

# Path commands

@app.command()
def gatherfiles(directory : Annotated[str, typer.Option(prompt = directory_prompt)],
                file_name : Annotated[str, typer.Option(prompt = multiple_file_prompt)],
                tab_name : Annotated[str, typer.Option(prompt = sheet_prompt)],
                values : Annotated[bool, typer.Option(prompt = 'Do you want to copy only values (y) or to keep forumla (n)')],):
    
    """
    Fonction agissant sur un dossier. Fonction qui prend plusieurs fichiers de structures identiques et 
    qui crée un fichier contenant l'ensemble des lignes de ces fichiers.

    Commande : 

        Version guidée : python xlspython.py gatherfiles

        Version complète : python xlspython.py gatherfiles --directory name --file name.xlsx --sheet name
    
    """ 
    path_object = Path(directory + '/')
    file_controller = OneFileCreatedController(new_path=path_object.pathname)
    controller = SeveralFoldersOneFileController(path_object, file_name, file_controller, dataonly=values)
    controller.apply_method_on_homononymous_files('copy_a_tab_at_tab_bottom', tab_name) 
    
@app.command()
def multidelcols(directory : Annotated[str, typer.Option(prompt = directory_prompt)],
                 file_name : Annotated[str, typer.Option(prompt = multiple_file_prompt)],
                 tab_name : Annotated[str, typer.Option(prompt = sheet_prompt)],
                 columns : Annotated[str, typer.Option(prompt = "Enter the group of column you want to KEEP. Respect the form A-D,E,G,H-J,Z without introducing any space.")]):
    """
    Fonction agissant sur un dossier. Fonction qui prend plusieurs fichiers de structures identiques et qui ne garde qu'un ensemble
    de colonnes prédéfinies. 

    Commande : 

        Version guidée : python xlspython.py multidelcols 
    """  
    pathobject = Path(directory + '/')  
    tab_controller = DeleteController(tab_name=tab_name)
    controller = SeveralFoldersOneFileController(pathobject, file_name, tab_controller)
    controller.apply_method_on_homononymous_tabs('delete_other_columns', columns)

# File commands

@app.command()
def filesave(file : Annotated[str, typer.Option(prompt = file_prompt)],
             dataonly: Annotated[bool, typer.Option(prompt = 'Do you want to save only values?')]):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. 
    Fonction qui crée une sauvegarde du fichier entré (file) et qui l'appelle name_file_time où time est le moment d'enregistrement.

    Commande : 

        Version guidée : python xlspython.py filesave

        Version complète : python xlspython.py filesave --file name.xlsx
    
    """
    file_object = File(file, dataonly=dataonly) 
    controller = OneFileCreatedController(file_object)
    controller.make_horodated_copy_of_a_file()

@app.command()
def multipletabs(file : Annotated[str, typer.Option(prompt = file_prompt)],
                 tab : Annotated[str, typer.Option(prompt = sheet_prompt)],
                 colread : Annotated[str, typer.Option(prompt = column_read_prompt)],
                 newfilepath : Annotated[Optional[str], typer.Option(prompt = 'If you want to divide a single file in tabs, press enter, otherwise your files must be included in folders themselves included in a bigger folder whose name must be written now.')] = '',
                 line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
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
        path_object = Path(newfilepath + '/')
        file_options = FileOptions(name_of_tab_to_read=tab, column_to_read=colread)
        controller = OneFileCreatedController(file_options=file_options, new_path=path_object.pathname, first_line=line)
        path_controller = SeveralFoldersOneFileController(path_object, file, controller)
        path_controller.apply_method_on_homononymous_files('split_one_tab_in_multiple_tabs')
        path = path_object.pathname

    # Apply command to a single file
    else:
        file_object = File(file)
        file_options = FileOptions(name_of_tab_to_read=tab, column_to_read=colread)
        controller = OneFileCreatedController(file_object, file_options=file_options, first_line=line)
        controller.split_one_tab_in_multiple_tabs()
        path = file_object.path
    
    # Eventually verify if each tab has the same number of lines
    check = typer.prompt('Do you want to check if all tabs have the same number of lines? If yes, write the number of lines else press enter', default="")
    if check:
        wrong_tabs = EvenTabsController(File(f'divided_{file}', path)).list_tabs_with_different_number_of_lines(int(check))
        if wrong_tabs:
            print('The following tabs have a different number of lines :' + ",".join(wrong_tabs))


@app.command()
def checklinenumber(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                    number : Annotated[int, typer.Option(prompt = 'Enter the expected number of lines in each tab')]):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui regarde si chaque onglet a un nombre de lignes égal au nombre attendu. Renvoie la liste des onglets qui n'ont pas le bon nombre
    de lignes

    Commande : 

        Version guidée : python xlspython.py checklinenumber

        Version complète : python xlspython.py checklinenumber --file name.xlsx --number 1
    
    """
    file_object = File(file)  
    controller = EvenTabsController(file_object)
    wrong_tabs = controller.list_tabs_with_different_number_of_lines(number)
    print(f'The tabs {wrong_tabs} does not have the expected number of lines.')

@app.command()
def extractcoltabs(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                   colread : Annotated[str, typer.Option(prompt = column_read_prompt)]):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui récupère une même colonne (colread) dans chaque onglet pour former une nouvelle feuille contenant toutes les colonnes.
    La première cellule de chaque colonne correspond alors au nom de l'onglet.

    Commande : 

        Version guidée : python xlspython.py extractcolsheets

        Version complète : python xlspython.py extractcolsheets --file name.xlsx --colread columnletter
    
    """
    file_object = File(file) 
    file_options = FileOptions(column_to_read=colread)
    controller = OneTabCreatedController(file_object, file_options)
    controller.extract_a_column_from_all_tabs(colread)

@app.command()
def cutsendmail(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                send : Annotated[bool, typer.Option(prompt = 'Do you want to send files by mail? Press y (yes) or enter (no)')]):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez fabriquer un fichier par onglet. Chaque fichier aura le nom de l'onglet. Vous souhaitez éventuellement envoyer chaque fichier à la personne associée.
    Attention, pour utiliser cette fonction, les onglets doivent être de la forme "prenom nom" sans caractère spéciaux. 

    Commande : 

        Version guidée : python xlspython.py cutsendmail

        Version complète : python xlspython.py cutsendmail 
    """
    file_object = File(file, dataonly=True)
    controller = OneFileCreatedController(file_object)
    controller.create_one_file_by_tab()

    if send:
        _write_mails_and_send_it()

def _write_mails_and_send_it():
    objet = typer.prompt('Please enter the object of your email',default="")
    message = typer.prompt('Please enter the message of your email',default="") 
    mail = Mail(sender_mail="tony.fevrier62@gmail.com", subject=objet, message=message, password="qkxqzhlvsgdssboh")
    path = Path('multifiles/')
    mail_controller = PathMailSender(path, mail, "@universite-paris-saclay.fr" )#"@etu-upsaclay.fr")
    json_file = typer.prompt('Please enter the name of the json file containing mail adresses. If you want to send to the mail paris-saclay, just press enter',default="") 
    if json_file:
        mail_controller.send_files_by_mail_using_(json_file)
    else:
        mail_controller.send_files_by_mail_using_same_domain()

@app.command()
def gathercolumn(file : Annotated[str, typer.Option(prompt = file_prompt)],
                 tab : Annotated[str, typer.Option(prompt = sheet_prompt)],
                 column_lists : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez des groupes de colonnes de valeurs avec une étiquette en première cellule. Pour chaque groupe, vous souhaitez former deux colonnes de valeurs : l'une qui contient
        les valeurs rassemblées en une colonne, l'autre, à sa gauche, qui indique l'étiquette de la colonne dans laquelle elle a été prise.
    
    Commande : 

        Version guidée : python xlspython.py gathercolumn

        Version complète : python xlspython.py gathercolumn --file nom --sheet onglet --columnlists A-D,E,G,H-J,Z
    """
    store = InputStore(column_lists, group_column_prompt)
    strings = store.ask_argument_until_none()
    column_lists = String.get_columns_from_several(*strings)

    file_object = File(file)
    file_options = FileOptions(name_of_tab_to_read=tab)
    controller = OneTabCreatedController(file_object, file_options)
    controller.gather_groups_of_multiple_columns_in_tabs_of_two_columns_containing_tags_and_values(*column_lists)

@app.command()
def extractcelltabs(file : Annotated[str, typer.Option(prompt = file_prompt)], 
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
    store = InputStore(cells, ask_argument_prompt('cell'))
    cells = store.ask_argument_until_none() 

    file_object = File(file)
    controller = OneFileCreatedController(file_object)
    controller.extract_cells_from_all_tabs(*cells)


@app.command()
def stringinbinary(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                   colread : Annotated[str, typer.Option(prompt = column_read_prompt)],
                   colwrite : Annotated[str, typer.Option(prompt = column_store_prompt)], 
                   tabs : Annotated[Optional[List[str]], typer.Option()] = None,
                   answers : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un ou plusieurs onglets d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Cette fonction lit les cellules d'une colonne (colread) et crée une nouvelle colonne (colwrite) contenant 1 si la valeur de la cellule est dans les bonnes réponses (answers)
    0 sinon.

    Commande : 

        Version guidée : python xlspython.py stringinbinary

        Version complète : python xlspython.py stringinbinary --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --answers chaine1 --answers chaine2
    
    """ 
    file_object = File(file)
    tab_options = TabOptions(column_to_read=colread, column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)
    
    store_answers = InputStore(answers, ask_argument_prompt('answer'))
    answers = store_answers.ask_argument_until_none()
      
    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('transform_string_in_binary_in_column', *answers) 

# Créer un fichier test pour tester cette fonction.
@app.command()
def cpcolumnontabs(file : Annotated[str, typer.Option(prompt = file_prompt)],
                   column : Annotated[List[str], typer.Option()] = None):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui reproduit les formules d'une ou plusieurs colonnes (column) du premier onglet sur toutes les colonnes situées à la même position dans les 
          autres onglets.
    
    Commande : 

        Version guidée : python xlspython.py cpcolumnonsheets
        Version complète : python xlspython.py cpcolumnonsheets --file name.xlsx --column columnletter
    """
    store = InputStore(column, ask_argument_prompt('column'))
    columns = store.ask_argument_until_none()

    fileobject = File(file, dataonly = False)
    file_options = FileOptions(columns_to_read=columns)
    controller = EvenTabsController(fileobject, file_options)
    controller.apply_columns_formula_on_all_tabs() 

@app.command()
def cpcellontabs(file : Annotated[str, typer.Option(prompt = file_prompt)],
                 cell : Annotated[List[str], typer.Option()] = None):
    
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui reproduit les formules d'une ou plusieurs cellules (cell) du premier onglet sur toutes les cellules situées à la même position dans les 
          autres onglets.
    
    Commande : 

        Version guidée : python xlspython.py cpcellonsheets
        Version complète : python xlspython.py cpcellonsheets --file name.xlsx --cell C5 --cell C17
    """
    store = InputStore(cell, ask_argument_prompt('cell'))
    cells = store.ask_argument_until_none()

    fileobject = File(file)
    controller = EvenTabsController(fileobject)
    controller.apply_cells_formula_on_all_tabs(*cells) 
    
@app.command()
def mergecells(file : Annotated[str, typer.Option(prompt = file_prompt)],
               start_column : Annotated[str, typer.Option(prompt = 'Enter the first column of cells to merge :')],
               end_column : Annotated[str, typer.Option(prompt = 'Enter the last column of cells to merge :')],
               start_line : Annotated[str, typer.Option(prompt = 'Enter the first row of  cells to merge :')],
               end_line : Annotated[str, typer.Option(prompt = 'Enter the last row of cells to merge :')]):
    
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Fonction qui fusionne les cellules de start_column à end_column et de start_row à end_row.
    
    Commande : 

        Version guidée : python xlspython.py mergecells
        Version complète : python xlspython.py mergecells --file name.xlsx --start_column columnletter --end_column columnletter --start_row rowindex --end_column rowindex 
    """
    
    fileobject = File(file)
    controller = EvenTabsController(fileobject)
    merged_cells_range = MergedCellsRange(start_column, end_column, start_line, end_line)
    controller.merge_cells_on_all_tabs(merged_cells_range)

 
@app.command()
def convertminutes(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                   colread : Annotated[str, typer.Option(prompt = column_read_prompt)],
                   colwrite : Annotated[str, typer.Option(prompt = column_store_prompt)], 
                   tabs : Annotated[Optional[List[str]], typer.Option()] = None,
                   line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous avez une colonne (colread) contenant des temps de la forme xx jours 5 heures 10 min 5 s.
     Vous souhaitez convertir dans une colonne (colwrite) les temps en minutes.

    Commande : 

        Version guidée : python xlspython.py convertminutes
        Version complète : python xlspython.py convertminutes --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --line linenumber
    
    """
    file_object = File(file)
    tab_options = TabOptions(column_to_read=colread, column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options, first_line=line) 
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)
    
    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('convert_time_in_minutes_in_columns') 

@app.command()
def groupofanswers(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                   colread : Annotated[str, typer.Option(prompt = column_read_prompt)],
                   colwrite : Annotated[str, typer.Option(prompt = column_store_prompt)], 
                   tabs : Annotated[Optional[List[str]], typer.Option()] = None,
                   line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous avez une colonne (colread) contenant des réponses. Chacune de ses réponses appartient à un groupe.
    Vous souhaitez afficher dans une colonne (colwrite) le groupe d'appartenance de la réponse. On vous demandera d'entrer des noms de groupes et dans la foulée, les réponses qui appartiennent au groupe.

    Commande : 

        Version guidée : python xlspython.py groupofanswers
        Version complète : python xlspython.py groupofanswers --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --line linenumber
    """
    file_object = File(file)
    tab_options = TabOptions(column_to_read=colread, column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options, first_line=line)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    # Creation of the groups of answers dictionary 
    map_store = MapStore("Enter one answer associated with this group:", 
                         "Enter the name of one group of answers of type answer1,answer2,answer3")
    map_store.create_mapping() 
    
    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('insert_group_associated_with_answer', map_store.mapping)     

@app.command()
def colorcasescolumn(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                     column : Annotated[str, typer.Option(prompt = column_read_prompt)],
                     tabs : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir une colonne (column) et colorer certaines chaînes de caractères dans cette colonne. Ces chaînes et les couleurs associées vous seront demandées
    durant l'exécution de la commande.

    Commande : 

        Version guidée : python xlspython.py colorcasescolumn 
        Version complète : python xlspython.py colorcasescolumn --file name.xlsx --sheet nametab --column columnletter
    """
    file_object = File(file)
    tab_options = TabOptions(column_to_read=column)
    tab_controller = ColorTabController(file_object, tab_options=tab_options) 
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    map_store = MapStore("Please enter a string which will be colored", color_prompt)
    map_store.create_mapping() 

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('color_cases_in_column', map_store.mapping)  

def _choose_tabs_to_modify(file_object, tabs, prompt):
    store = InputStore(tabs, prompt)
    tabs = store.ask_argument_until_none()
    if len(tabs) == 0:
        tabs = file_object.sheets_name
    return FileOptions(names_of_tabs_to_modify=tabs)     

@app.command()
def colorcasestab(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                  tabs : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir un onglet (sheet) et colorer certaines chaînes de caractères. Ces chaînes et les couleurs associées vous seront demandées
    durant l'exécution de la commande.

    Commande : 

        Version guidée : python xlspython.py colorcasestab 

        Version complète : python xlspython.py colorcasestab --file name.xlsx --sheet nametab
    """
    file_object = File(file)
    tab_options = TabOptions()
    tab_controller = ColorTabController(file_object, tab_options=tab_options) 
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    map_store = MapStore("Please enter a string which will be colored", color_prompt)
    map_store.create_mapping() 

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('color_cases_in_tab', map_store.mapping) 

@app.command()
def addcolumn(file_from : Annotated[str, typer.Option(prompt = 'Enter the xlsx file in which you want to write ')],
            tab_from : Annotated[str, typer.Option(prompt = 'Enter the corresponding sheet name')],
            colread_from : Annotated[str, typer.Option(prompt = 'Enter the column of this sheet containing the identifiers')],
            colwrite : Annotated[str, typer.Option(prompt = 'Enter the column from which you want to write')],
            file_to : Annotated[str, typer.Option(prompt = 'Enter the xlsx file from which you import data ')],
            tab_to : Annotated[str, typer.Option(prompt = 'Enter the corresponding sheet name')],
            colread_to : Annotated[str, typer.Option(prompt = 'Enter the column of this sheet containing the identifiers')],
            col_to_import : Annotated[Optional[List[str]], typer.Option()] = None):                    
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
      Vous souhaitez importer des colonnes (colimport) d'un fichier (file2, sheet2) dans un autre fichier (file,sheet).
    La mise en correspondance entre les deux onglets s'effectue via une colonne d'identifiants dans chaque fichier (colread, colread2). L'important s'effectue 
    dans file à partir de la colonne colwrite.    

    Commande : 

        Version guidée : python xlspython.py addcolumn 

        Version complète : python xlspython.py addcolumn --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --file2 name.xlsx --sheet2 nametab --colread2 columnletter --colimport col1 --colimport col2
    
    """ 
    file_object_from = File(file_from)
    file_object_to = File(file_to)
    controller = TwoFilesController(file_object_from, file_object_to, tab_from, tab_to, colread_from, colread_to)  

    store = InputStore(col_to_import, ask_argument_prompt('column to import'))
    col_to_import = store.ask_argument_until_none()

    controller.copy_columns_in_a_tab_differently_sorted(col_to_import, colwrite) 

@app.command()
def colorlines(file : Annotated[str, typer.Option(prompt = file_prompt)], 
               color : Annotated[str, typer.Option(prompt = 'Enter the color in a hexadecimal format')],
               tabs : Annotated[Optional[List[str]], typer.Option()] = None,
               strings : Annotated[Optional[List[str]], typer.Option()] = None):                    

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir un onglet (sheet) et colorer (color) les lignes contenant certaines chaînes de caractères (strings).

    Commande : 

        Version guidée : python xlspython.py colorlines 

        Version complète : python xlspython.py colorlines --file name.xlsx --sheet nametab --color colorinhexadecimal --strings chaine1 --strings chaine2
    """
    file_object = File(file)
    tab_options = TabOptions()
    tab_controller = ColorTabController(file_object, tab_options=tab_options, color=color) 
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    strings_store = InputStore(strings, ask_argument_prompt('string'))
    strings = strings_store.ask_argument_until_none()

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('color_lines_containing_strings', *strings) 

@app.command()
def cutstring(file : Annotated[str, typer.Option(prompt = file_prompt)], 
              colcut : Annotated[str, typer.Option(prompt = 'Enter the column containing strings to cut')],
              colwrite : Annotated[str, typer.Option(prompt = column_store_prompt)], 
              tabs : Annotated[Optional[List[str]], typer.Option()] = None,
              separator : Annotated[Optional[str], typer.Option(prompt = separator_prompt)] = ','):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. 
    Une colonne (colcut) contient des chaînes de caractères séparées par un symbole (separator). Vous souhaitez les couper en morceaux 
    et créer des colonnes (à partir de colwrite) pour chacun de ces morceaux.

    Commande : 

        Version guidée : python xlspython.py cutstring 

        Version complète : python xlspython.py cutstring --file name.xlsx --sheet nametab --colcut columnletter --colwrite columnletter --separator symbol
    """
    file_object = File(file)
    tab_options = TabOptions(column_to_read=colcut, column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('insert_splitted_strings_of', separator) 

@app.command()
def deletecols(file : Annotated[str, typer.Option(prompt = file_prompt)], 
               columns : Annotated[str, typer.Option(prompt = group_column_prompt)],
               tabs : Annotated[Optional[List[str]], typer.Option()] = None):

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Prend une séquence de colonnes et les supprime.

    Commande : 

        Version guidée : python xlspython.py deletecols  

        Version complète : python xlspython.py deletecols --file name.xlsx --sheet nametab --columns A --columns D 
    
    """ 
    file_object = File(file)
    tab_controller = DeleteController(file_object)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('delete_columns', columns) 

@app.command()
def keepcols(file : Annotated[str, typer.Option(prompt = file_prompt)], 
             columns : Annotated[str, typer.Option(prompt = group_column_prompt)],
             tabs : Annotated[Optional[List[str]], typer.Option()] = None):

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous entrez les colonnes que vous souhaitez garder. Toutes les autres seront supprimées.

    Commande : 

        Version guidée : python xlspython.py deletecols  

        Version complète : python xlspython.py deletecols --file name.xlsx --sheet nametab --columns A --columns D 
    
    """  
    file_object = File(file)
    tab_controller = DeleteController(file_object)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('delete_other_columns', columns) 


@app.command()
def deletelinesstr(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                colread : Annotated[str, typer.Option(prompt = column_read_prompt)],
                tabs : Annotated[Optional[List[str]], typer.Option()] = None,
                strings : Annotated[Optional[List[str]], typer.Option()] = None):                    
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir une colonne (colread) et si une chaîne (strings) apparaît dans cette colonne, supprimer la ligne associée.

    Commande : 

        Version guidée : python xlspython.py deletelines 

        Version complète : python xlspython.py deletelines --file name.xlsx --sheet nametab --colread columnletter --strings chaine1 --strings chaine2
    
    """ 
    file_object = File(file)
    tab_controller = DeleteController(file_object)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    strings_store = InputStore(strings, ask_argument_prompt('string'))
    strings = strings_store.ask_argument_until_none()    

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('delete_lines_containing_strings_in_given_column', colread, *strings) 


@app.command()
def deletetwins(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                colread : Annotated[str, typer.Option(prompt = column_read_prompt)],
                tabs : Annotated[Optional[List[str]], typer.Option()] = None,
                line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
      Certains participants à un questionnaire répondent plusieurs fois. Vous souhaitez parcourir une colonne (colread)
     qui les identifie et ne garder que leur dernière réponse à ce questionnaire.

    Commande : 

        Version guidée : python xlspython.py deletetwins 

        Version complète : python xlspython.py deletetwins --file name.xlsx --sheet nametab --colread columnletter --line linenumber
    """
    file_object = File(file)
    tab_controller = DeleteController(file_object)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('delete_twins_lines_and_color_last_twin', colread) 


@app.command()
def columnbyqcmanswer(file : Annotated[str, typer.Option(prompt = file_prompt)],
                     colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                     colwrite : Annotated[str, typer.Option(prompt = 'Enter the column from which you want to write')], 
                     tabs : Annotated[Optional[List[str]], typer.Option()] = None,
                     answers : Annotated[Optional[List[str]], typer.Option()] = None,
                     line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
                   

    """
    Fonction agissant sur un ou plusieurs onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. 
    Une colonne (colread) contient toutes les réponses d'un participant à une question de QCM. Vous souhaitez créer autant de colonnes que
    de réponses (answers) à la question et mettre dans chaque colonne (à partir de colwrite) si les participants l'ont coché ou non (list).

    Commande : 

        Version guidée : python xlspython.py columnbyqcmanswer 

        Version complète : python xlspython.py columnbyqcmanswer --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --answers chaine1 --answers chaine2 --list oui non
    
    """
    file_object = File(file)
    tab_options = TabOptions(column_to_read=colread, column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options, first_line=line)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    answers_store = InputStore(answers, ask_argument_prompt('QCM answer'))
    answers = answers_store.ask_argument_until_none()

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('fill_one_column_by_QCM_answer', *answers) 

@app.command()
def gathermultianswers(file : Annotated[str, typer.Option(prompt = file_prompt)],
                       tab : Annotated[str, typer.Option(prompt = sheet_prompt)],
                       colread : Annotated[str, typer.Option(prompt = column_read_prompt)],
                       colstore : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the data to store')],
                       line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
      Certains participants à un questionnaire répondent plusieurs fois. Vous souhaitez parcourir une colonne (colread)
     qui les identifie et créer, dans un autre onglet, une ligne par participant ayant répondu plusieurs fois. Cette ligne contient les différentes
     réponses de ce participant contenues dans une colonne (colstore) donnée.

    Commande : 

        Version guidée : python xlspython.py gathermultianswers 

        Version complète : python xlspython.py gathermultianswers --file name.xlsx --sheet nametab --colread columnletter --colstore columnletter --line linenumber
    
    """
    file_object = File(file)
    file_options = FileOptions(name_of_tab_to_read=tab)
    controller = OneTabCreatedController(file_object, file_options=file_options, first_line=line)
    controller.gather_multiple_answers(colread, colstore)  

@app.command()
def maxnames(file : Annotated[str, typer.Option(prompt = file_prompt)], 
             colstore : Annotated[str, typer.Option(prompt = column_store_prompt)],
             tabs : Annotated[Optional[List[str]], typer.Option()] = None,
             columnlist : Annotated[Optional[List[str]], typer.Option()] = None,
             line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
    
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez une liste de colonnes avec des chiffres, chaque colonne a un nom dans sa première cellule.
      Cette fonction crée une colonne dans laquelle on entre pour chaque ligne le nom de la colonne ou des colonnes qui contient le max.

    Commande : 

        Version guidée : python xlspython.py maxnames 

        Version complète : python xlspython.py maxnames --file name.xlsx --sheet nametab --colstore columnletter --columnlist A --columnlist C
    
    """
    file_object = File(file)
    column_store = InputStore(columnlist, "Enter the letter of a column you want to read")
    columnlist = column_store.ask_argument_until_none()
    tab_options = TabOptions(columns_to_read= columnlist, column_to_write=colstore)
    tab_controller = InsertController(file_object, tab_options=tab_options, first_line=line)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('insert_tags_of_maximum_of_column_list' ) 


@app.command()
def colcongruent(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                firstcol : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the word (example: prime, probe)')],
                secondcol : Annotated[str, typer.Option(prompt = 'Enter the column letter corresponding to prime')],
                thirdcol : Annotated[str, typer.Option(prompt = 'Enter the column letter corresponding to probe')],
                colwrite : Annotated[str, typer.Option(prompt = 'Enter the column letter in which you want to write')],
                tabs : Annotated[Optional[List[str]], typer.Option()] = None):
    
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    On a deux colonnes, une contenant des mots, l'autre contenant une chaine de la forme motnb_o.jpg (congruent1_o.jpg), on veut créer
    une colonne qui contient le mot de la première colonne + mot (congruent) si le mot de la première colonne fait partie d'une liste de mots prédéfinie.

    Commande : 

        Version guidée : python xlspython.py colcongruent
 
    """
    file_object = File(file)
    tab_options = TabOptions(columns_to_read=[firstcol, secondcol, thirdcol], column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('insert_column_for_prime_probe_congruence') 

@app.command()
def colgetpart(file : Annotated[str, typer.Option(prompt = file_prompt)], 
                colread : Annotated[str, typer.Option(prompt = 'Enter the column letter to read ')],
                colwrite : Annotated[str, typer.Option(prompt = 'Enter the column letter where to write')],
                separator : Annotated[str, typer.Option(prompt = 'Enter the separator')],
                piece : Annotated[int, typer.Option(prompt = 'Enter the number of the part you want to get. For example, tap 1 to get the begin of the word')],
                tabs : Annotated[Optional[List[str]], typer.Option()] = None,
                line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez une colonne qui contient une chaîne qui contient un séparateur dont vous voulez prendre une partie. 
    Ce mot est inséré dans une nouvelle colonne. 

    Commande : 

        Version guidée : python xlspython.py colgetbegin
    """
    file_object = File(file)
    tab_options = TabOptions(column_to_read=colread, column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options, first_line=line)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('write_piece_of_string_in_column', separator, piece - 1) 

@app.command()
def maptwocols(file : Annotated[str, typer.Option(prompt = file_prompt)], 
               colread : Annotated[str, typer.Option(prompt = 'Enter the column letters to read separated by a comma (C,E)')],
               colwrite : Annotated[str, typer.Option(prompt = column_store_prompt)],
               tabs : Annotated[Optional[List[str]], typer.Option()] = None,
               line : Annotated[Optional[int], typer.Option(prompt = line_prompt)] = '2'):
    """
    Fonction agissant sur un onglet d'un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous avez deux colonnes de lecture, suivant ce qui est écrit sur une ligne, vous voulez ou non insérer quelque chose 
    dans une nouvelle colonne.

    Commande : 

        Version guidée : python xlspython.py maptwocols
    """
    file_object = File(file)
    tab_options = TabOptions(columns_to_read=colread.split(","), column_to_write=colwrite)
    tab_controller = InsertController(file_object, tab_options=tab_options, first_line=line)
    file_options = _choose_tabs_to_modify(file_object, tabs, multiple_tabs_prompt)

    map_store = MapStore("Enter a value you want to put in the new column", 
                         "Enter the two strings which should lead to this new value. You must enter it with the same order as the order you entered the columns to read")
    map_store.create_mapping() 
    
    controller = MultipleSameTabController(file_object, tab_controller, file_options)
    controller.apply_method_on_some_tabs('map_two_columns_to_a_third_column', map_store.mapping) 


if __name__ == "__main__":
    app()
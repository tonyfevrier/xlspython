import typer
from typing import Optional, List, Tuple
from module_pour_excel import File, Sheet, Str
from typing_extensions import Annotated
from utils import UtilsForcommands as Ufc

app = typer.Typer()


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
def multipletabs(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                 sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name ')],
                 colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing strings ')],
                 line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un fichier. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.

    Vous avez un fichier xlsx dont une colonne (colread) contient des participants qui ont pu répondre plusieurs fois à un questionnaire. 
    Vous souhaitez créer un onglet par participant avec toutes les lignes qui correspondent.

    Commande : 

        Version guidée : python xlspython.py multipletabs

        Version complète : python xlspython.py multipletabs --file name.xlsx --sheet nametab --colread columnletter --line linenumber
    
    """
    fileobject = File(file) 
    fileobject.create_one_onglet_by_participant(sheet,colread,first_line=line)
    

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
    fileobject = File(file,dataonly=True) 
    fileobject.extract_column_from_all_sheets(colread)

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

    if sendmail ==  'n':
        fileobject.one_file_by_tab_sendmail()
    else:
        objet = typer.prompt('Please enter the object of your email',default="")
        message = typer.prompt('Please enter the message of your email',default="") 
        jsonfile = typer.prompt('Please enter the name of the json file containing mail adresses. If you want to send to the mail paris-saclay, just press enter',default="") 
        fileobject.one_file_by_tab_sendmail(send = True, adressjson = jsonfile, objet = objet, message = message)


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
    fileobject.gather_columns_in_one(sheet,*column_lists)

@app.command()
def stringinbinary(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   answers : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.

    Cette fonction lit les cellules d'une colonne (colread) et crée une nouvelle colonne (colwrite) contenant 1 si la valeur de la cellule est dans les bonnes réponses (answers)
    0 sinon.

    Commande : 

        Version guidée : python xlspython.py stringinbinary

        Version complète : python xlspython.py stringinbinary --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --answers chaine1 --answers chaine2
    
    """ 
    answers = Ufc.askArgumentUntilNone(answers,"Enter one good answer and then press enter. Press directly enter if you have entered all the good answers")
    
    sheetobject = Sheet(file,sheet)
 
    bool = Ufc.insertOrOverwrite(colwrite)
    sheetobject.column_transform_string_in_binary(colread,colwrite,*answers, insert = bool) 

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
    fileobject.apply_column_formula_on_all_sheets(*columns) 

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
    fileobject.apply_cells_formula_on_all_sheets(*cells) 
    
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
    fileobject.merge_cells_on_all_tabs(start_column,end_column,start_row,end_row)


 
@app.command()
def convertminutes(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous avez une colonne (colread) contenant des temps de la forme xx jours 5 heures 10 min 5 s.
     Vous souhaitez convertir dans une colonne (colwrite) les temps en minutes.

    Commande : 

        Version guidée : python xlspython.py convertminutes

        Version complète : python xlspython.py convertminutes --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --line linenumber
    
    """
    sheetobject = Sheet(file,sheet)
    bool = Ufc.insertOrOverwrite(colwrite)
    sheetobject.column_convert_in_minutes(colread,colwrite,line_beginning=line, insert=bool)

@app.command()
def groupofanswers(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')],  
                   line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous avez une colonne (colread) contenant des réponses. Chacune de ses réponses appartient à un groupe.
    Vous souhaitez afficher dans une colonne (colwrite) le groupe d'appartenance de la réponse.

    Commande : 

        Version guidée : python xlspython.py groupofanswers

        Version complète : python xlspython.py groupofanswers --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --line linenumber
    """
    #Creation of the groups of answers dictionary 
    groups_of_responses = Ufc.createDictListValueByCmd("Enter the name of one group of answers")

    sheetobject = Sheet(file,sheet)
    bool = Ufc.insertOrOverwrite(colwrite)
    sheetobject.column_set_answer_in_group(colread,colwrite,groups_of_responses, line_beginning = line, insert=bool)

@app.command()
def colorcasescolumn(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                     sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                     column : Annotated[str, typer.Option(prompt = 'Enter the column')]):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir une colonne (column) et colorer certaines chaînes de caractères dans cette colonne. Ces chaînes et les couleurs associées vous seront demandées
    durant l'exécution de la commande.

    Commande : 

        Version guidée : python xlspython.py colorcasescolumn 

        Version complète : python xlspython.py colorcasescolumn --file name.xlsx --sheet nametab --column columnletter
    """
    #Creation of the dictionary with the strings to color and their color
    color = Ufc.createDictByCmd("Please enter a string which will be colored", "Please enter the color in hexadecimal type")

    sheetobject = Sheet(file,sheet)
    sheetobject.color_special_cases_in_column(column,color)

@app.command()
def colorcasestab(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                  sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')]):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir un onglet (sheet) et colorer certaines chaînes de caractères. Ces chaînes et les couleurs associées vous seront demandées
    durant l'exécution de la commande.

    Commande : 

        Version guidée : python xlspython.py colorcasestab 

        Version complète : python xlspython.py colorcasestab --file name.xlsx --sheet nametab
    """
    #Creation of the dictionary with the strings to color and their color
    color = Ufc.createDictByCmd("Please enter a string which will be colored", "Please enter the color in hexadecimal type")

    sheetobject = Sheet(file,sheet)
    sheetobject.color_special_cases_in_sheet(color)

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
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous souhaitez importer des colonnes (colimport) d'un fichier (file2, sheet2) dans un autre fichier (file,sheet).
    La mise en correspondance entre les deux onglets s'effectue via une colonne d'identifiants dans chaque fichier (colread, colread2). L'important s'effectue 
    dans file à partir de la colonne colwrite.    

    Commande : 

        Version guidée : python xlspython.py addcolumn 

        Version complète : python xlspython.py addcolumn --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --file2 name.xlsx --sheet2 nametab --colread2 columnletter --colimport col1 --colimport col2
    
    """ 
    colimport = Ufc.askArgumentUntilNone(colimport,"Enter one column to import and then press enter. Press directly enter if you have entered all the columns to import.")
    sheetobject = Sheet(file,sheet)
    sheetobject.add_column_in_sheet_differently_sorted(colread,colwrite,[file2,sheet2,colread2,colimport])

@app.command()
def colorlines(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
               sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
               color : Annotated[str, typer.Option(prompt = 'Enter the color in a hexadecimal format')],
               strings : Annotated[Optional[List[str]], typer.Option()] = None):                    

    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls.
    Vous souhaitez parcourir un onglet (sheet) et colorer (color) les lignes contenant certaines chaînes de caractères (strings).

    Commande : 

        Version guidée : python xlspython.py colorlines 

        Version complète : python xlspython.py colorlines --file name.xlsx --sheet nametab --color colorinhexadecimal --strings chaine1 --strings chaine2
    """
    strings = Ufc.askArgumentUntilNone(strings,"Enter one string leading to the coloration of the line and then press enter. Press directly enter if you have entered all the strings.")
    
    sheetobject = Sheet(file,sheet)
    sheetobject.color_lines_containing_chaines(color,*strings)

@app.command()
def cutstring(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
              sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
              colcut : Annotated[str, typer.Option(prompt = 'Enter the column containing strings to cut')],
              colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
              separator : Annotated[Optional[str], typer.Option(prompt = '(Optional) Enter the separator or press enter')] = ','):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Une colonne (colcut) contient des chaînes de caractères séparées par un symbole (separator). Vous souhaitez les couper en morceaux 
    et créer des colonnes (à partir de colwrite) pour chacun de ces morceaux.

    Commande : 

        Version guidée : python xlspython.py cutstring 

        Version complète : python xlspython.py cutstring --file name.xlsx --sheet nametab --colcut columnletter --colwrite columnletter --separator symbol
    """
    sheetobject = Sheet(file,sheet) 
    sheetobject.column_cut_string_in_parts(colcut,colwrite,separator,insert = bool)

@app.command()
def deletelines(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                strings : Annotated[Optional[List[str]], typer.Option()] = None):                    
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Vous souhaitez parcourir une colonne (colread) et si une chaîne (strings) apparaît dans cette colonne, supprimer la ligne associée.

    Commande : 

        Version guidée : python xlspython.py deletelines 

        Version complète : python xlspython.py deletelines --file name.xlsx --sheet nametab --colread columnletter --strings chaine1 --strings chaine2
    
    """
    strings = Ufc.askArgumentUntilNone(strings,"Enter one string leading to the delation of the line and then press enter. Press directly enter if you have entered all the strings.")

    sheetobject = Sheet(file,sheet)
    sheetobject.delete_lines(colread,*strings)

@app.command()
def deletetwins(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Certains participants à un questionnaire répondent plusieurs fois. Vous souhaitez parcourir une colonne (colread)
     qui les identifie et ne garder que leur dernière réponse à ce questionnaire.

    Commande : 

        Version guidée : python xlspython.py deletetwins 

        Version complète : python xlspython.py deletetwins --file name.xlsx --sheet nametab --colread columnletter --line linenumber
    """
    sheetobject = Sheet(file,sheet)
    sheetobject.delete_doublons(colread,line_beginning=line)

@app.command()
def columnbyqcmanswer(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column from which you want to write')], 
                   answers : Annotated[Optional[List[str]], typer.Option()] = None,
                   list : Annotated[Tuple[str, str], typer.Option(prompt = 'Enter what you want to write in the cells or press enter')] = ('oui', 'non')):

    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Une colonne (colread) contient toutes les réponses d'un participant à une question de QCM. Vous souhaitez créer autant de colonnes que
    de réponses (answers) à la question et mettre dans chaque colonne (à partir de colwrite) si les participants l'ont coché ou non (list).

    Commande : 

        Version guidée : python xlspython.py columnbyqcmanswer 

        Version complète : python xlspython.py columnbyqcmanswer --file name.xlsx --sheet nametab --colread columnletter --colwrite columnletter --answers chaine1 --answers chaine2 --list oui non
    
    """
    answers = Ufc.askArgumentUntilNone(answers,"Enter one QCM answer and then press enter. Press directly enter if you have entered all the answers.")

    sheetobject = Sheet(file,sheet)
    sheetobject.create_one_column_by_QCM_answer(colread,colwrite,list,*answers)

@app.command()
def gathermultianswers(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                       sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                       colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                       colstore : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the data to store ')],
                       line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet. Pensez à mettre le fichier sur lequel vous appliquez la commande dans un dossier nommé fichiers_xls. Certains participants à un questionnaire répondent plusieurs fois. Vous souhaitez parcourir une colonne (colread)
     qui les identifie et créer, dans un autre onglet, une ligne par participant ayant répondu plusieurs fois. Cette ligne contient les différentes
     réponses de ce participant contenues dans une colonne (colstore) donnée.

    Commande : 

        Version guidée : python xlspython.py gathermultianswers 

        Version complète : python xlspython.py gathermultianswers --file name.xlsx --sheet nametab --colread columnletter --colstore columnletter --line linenumber
    
    """
    sheetobject = Sheet(file,sheet)
    sheetobject.gather_multiple_answers(colread,colstore,line_beggining=line)

@app.command()
def maxnames(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
             sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
             colstore : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the data to store ')],
             columnlist : Annotated[Optional[List[str]], typer.Option()] = None):
    """
    Vous avez une liste de colonnes avec des chiffres, chaque colonne a un nom dans sa première cellule. Cette fonction crée une colonne dans laquelle on entre pour chaque 
    ligne le nom de la colonne ou des colonnes qui contient le max.

    Commande : 

        Version guidée : python xlspython.py maxnames 

        Version complète : python xlspython.py maxnames --file name.xlsx --sheet nametab --colstore columnletter --columnlist A --columnlist C
    
    """
    columnlist = Ufc.askArgumentUntilNone(columnlist, "Enter the letter of a column you want to read")

    sheetobject = Sheet(file,sheet)
    sheetobject.give_names_of_maximum(colstore, *columnlist)


if __name__ == "__main__":
    app()
import typer
from typing import Optional, List, Tuple
from module_pour_excel import *
from typing_extensions import Annotated
from utils import UtilsForcommands as Ufc

"""  
tester toutes les commandes : reste formulaonsheets, 
faire formulaonsheets 
Ajouter qqch disant à l'utilisateur mettant help qu'il doit mettre les fichiers dans un dossier fichiers_xls.
Ajouter dans chaque docstring une écriture complète de la commande.
Ajouter les arguments optionnels de mod pour excel que je n'ai pas mis encore.
Nettoyer les insert = True security etc qui ne servent plus.
Ajouter une fonction qui prend n colonnes et qui crée deux grandes colonne à partir d'elles : une avec les valeurs et une avec les noms de la colonne correspondante en face. 
"""


def docstring_and_execute(command_function):
    print(command_function.__doc__)
    command_function()


app = typer.Typer()


@app.command()
def filesave(file : Annotated[str, typer.Option(prompt = 'Enter the file you want to save ')]):
    """
    Fonction agissant sur un fichier.
    
    Fonction qui crée une sauvegarde du fichier entré et qui l'appelle name_file_time où time est le moment d'enregistrement.
    """
    fileobject = File(file)
    fileobject.sauvegarde()



@app.command()
def multipletabs(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                 tab : Annotated[str, typer.Option(prompt = 'Enter the sheet name ')],
                 colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing strings ')],
                 line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un fichier.

    Vous avez un fichier xlsx dont une colonne contient des participants qui ont pu répondre plusieurs fois à un questionnaire. 
    Vous souhaitez créer un onglet par participant avec toutes les lignes qui correspondent.
    """
    fileobject = File(file) 
    fileobject.create_one_onglet_by_participant(tab,colread,first_line=line)
    

@app.command()
def extractcolsheets(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                     colread : Annotated[str, typer.Option(prompt = 'Enter the column letter ')]
                    ):
    """
    Fonction agissant sur un fichier.

    Fonction qui récupère une même colonne dans chaque onglet pour former une nouvelle feuille contenant toutes les colonnes.
    La première cellule de chaque colonne correspond alors au nom de l'onglet.
    """
    fileobject = File(file) 
    fileobject.extract_column_from_all_sheets(colread)

@app.command()
def stringinbinary(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   answers : Annotated[Optional[List[str]], typer.Option()] = None, 
                   ):
    """
    Fonction agissant sur un onglet.

    Cette fonction lit les cellules d'une colonne et crée une nouvelle colonne contenant 1 si la valeur de la cellule est dans les bonnes réponses (answers)
    0 sinon.

    L'option answers doit être écrite au format reponse1,reponse2 (la virgule sépare les réponses, ne pas mettre d'espace superflu).
    """
    answers = Ufc.askArgumentUntilNone(answers,"Enter one good answer and then press enter. Press directly enter if you have entered all the good answers")
    
    sheetobject = Sheet(file,sheet)
    sheetobject.column_transform_string_in_binary(colread,colwrite,*answers) 

# Créer un fichier test pour tester cette fonction.
# @app.command()
# def formulaonsheets(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
#                     *column_list : Annotated[str, typer.Option(prompt = 'Enter the column letter ')]
#                     ):
#     """
#     Fonction agissant sur un fichier.

#     Fonction qui reproduit les formules d'une ou plusieurs colonnes du premier onglet sur toutes les colonnes situées à la même position dans les 
#           autres onglets.
#     """
#     fileobject = File(file, dataonly = False)
#     fileobject.apply_column_formula_on_all_sheets(*column_list) 


 
@app.command()
def convertinminutes(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un onglet.
    """
    sheetobject = Sheet(file,sheet)
    sheetobject.column_convert_in_minutes(colread,colwrite,line_beginning=line)

@app.command()
def groupofanswers(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')],  
                   line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'
                   ):
    """
    Fonction agissant sur un onglet.
    """
    #Creation of the groups of answers dictionary 
    groups_of_responses = Ufc.createDictListValueByCmd("Enter the name of one group of answers")

    sheetobject = Sheet(file,sheet)
    sheetobject.column_set_answer_in_group(colread,colwrite,groups_of_responses, line_beginning = line)

@app.command()
def colorcasescolumn(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   column : Annotated[str, typer.Option(prompt = 'Enter the column')], 
                   ):
    """
    Fonction agissant sur un onglet.
    """
    #Creation of the dictionary with the strings to color and their color
    color = Ufc.createDictByCmd("Please enter a string which will be colored", "Please enter the color in hexadecimal type")

    sheetobject = Sheet(file,sheet)
    sheetobject.color_special_cases_in_column(column,color)

@app.command()
def colorcasestab(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')], 
                   ):
    """
    Fonction agissant sur un onglet.
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
    Fonction agissant sur un onglet.
    """
    colimport = Ufc.askArgumentUntilNone(colimport,"Enter one column to import and then press enter. Press directly enter if you have entered all the columns to import.")
    sheetobject = Sheet(file,sheet)
    sheetobject.add_column_in_sheet_differently_sorted(colread,colwrite,[file2,sheet2,colread2,colimport])

@app.command()
def colorlines(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   color : Annotated[str, typer.Option(prompt = 'Enter the color ')],
                   strings : Annotated[Optional[List[str]], typer.Option()] = None):                    

    """
    Fonction agissant sur un onglet.
    """
    strings = Ufc.askArgumentUntilNone(strings,"Enter one string leading to the coloration of the line and then press enter. Press directly enter if you have entered all the strings.")
    
    sheetobject = Sheet(file,sheet)
    sheetobject.color_lines_containing_chaines(color,*strings)

@app.command()
def cutstring(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colcut : Annotated[str, typer.Option(prompt = 'Enter the column containing strings to cut')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column where you want to write')], 
                   separator : Annotated[Optional[str], typer.Option(prompt = '(Optional) Enter the separator or press enter')] = ',' 
                   ):
    """
    Fonction agissant sur un onglet.
    """
    sheetobject = Sheet(file,sheet)
    sheetobject.column_cut_string_in_parts(colcut,colwrite,separator)

@app.command()
def deletelines(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                   strings : Annotated[Optional[List[str]], typer.Option()] = None):                    
    """
    Fonction agissant sur un onglet.
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
    Fonction agissant sur un onglet.
    """
    sheetobject = Sheet(file,sheet)
    sheetobject.delete_doublons(colread,line_beginning=line)

@app.command()
def columnbyqcmanswer(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column containing the answers')],
                   colwrite : Annotated[str, typer.Option(prompt = 'Enter the column from which you want to write')], 
                   answers : Annotated[Optional[List[str]], typer.Option()] = None,
                   list : Annotated[Tuple[str, str], typer.Option(prompt = 'Enter what you want to write in the cells or press enter')] = ('oui', 'non'),
                   ):

    """
    Fonction agissant sur un onglet.
    """
    answers = Ufc.askArgumentUntilNone(answers,"Enter one QCM answer and then press enter. Press directly enter if you have entered all the answers.")

    sheetobject = Sheet(file,sheet)
    sheetobject.create_one_column_by_QCM_answer(colread,colwrite,list,*answers)

@app.command()
def gathermultianswers(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')],
                   sheet : Annotated[str, typer.Option(prompt = 'Enter the sheet name')],
                   colread : Annotated[str, typer.Option(prompt = 'Enter the column letter containing identifiers ')],
                   colstore : Annotated[str, typer.Option(prompt = 'Enter the column letter containing the data to store ')],
                   line : Annotated[Optional[int], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'
                   ):
    """
    Fonction agissant sur un onglet.

    
    """
    sheetobject = Sheet(file,sheet)
    sheetobject.gather_multiple_answers(colread,colstore,line_beggining=line)



if __name__ == "__main__":
    app()
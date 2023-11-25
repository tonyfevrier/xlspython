import typer
from typing import Optional
from module_pour_excel import *
from typing_extensions import Annotated

""" 
Enlever to les lineend qui ne servent à rien en argument avec maxrow+1
Ajouter un décorateur qui donne la docstring de la fonction avant qu'on ne demande les arguments à remplir.

def docwrapper(func):
    def wrapper(*args, **kwargs): 
        print(func.__doc__)
        return func(*args,**kwargs)
    return wrapper
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
                 column_read : Annotated[str, typer.Option(prompt = 'Enter the column letter containing strings ')],
                 line : Annotated[Optional[str], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un fichier.

    Vous avez un fichier xlsx dont une colonne contient des participants qui ont pu répondre plusieurs fois à un questionnaire. 
    Vous souhaitez créer un onglet par participant avec toutes les lignes qui correspondent.
    """
    fileobject = File(file) 
    fileobject.create_one_onglet_by_participant(tab,column_read,int(line))
    

@app.command()
def extractcolsheets(file : Annotated[str, typer.Option(prompt = 'Enter the xlsx file ')], 
                     column_read : Annotated[str, typer.Option(prompt = 'Enter the column letter ')]
                    ):
    """
    Fonction agissant sur un fichier.

    Fonction qui récupère une même colonne dans chaque onglet pour former une nouvelle feuille contenant toutes les colonnes.
    La première cellule de chaque colonne correspond alors au nom de l'onglet.
    """
    fileobject = File(file) 
    fileobject.extract_column_from_all_sheets(column_read)
    
@app.command()
def formulaonsheets():
    """
    Fonction agissant sur un fichier.

    Fonction qui reproduit les formules d'une colonne ou plusieurs colonnes
          du premier onglet sur toutes les colonnes situées à la même position dans les 
          autres onglets.

        Input : 
            -column_list : int. les positions des colonnes où récupérer et coller.

        Exemples d'utilisation : 

            Bien veiller à mettre dataonly = False sinon il ne copiera pas les formules mais
            les valeurs des cellules. On peut aussi copier les valeurs des cellules : pour cela,
            enlever dataonly = False

            Sur une colonne
                file = File('dataset.xlsx', dataonly = False)
                file.apply_column_formula_on_all_sheets(2) 
 
    """
    pass

@app.command()
def stringinbinary():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def convertinminutes():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def stringinbinary():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def groupofanswers():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def colorcasescolumn():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def colorcasestab():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def addcolumn():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def colorlines():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def cutstring():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def deletelines():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def deletetwins():
    """
    Fonction agissant sur un onglet.
    """
    pass

@app.command()
def columnbyqcmanswer():
    """
    Fonction agissant sur un onglet.
    """
    pass


if __name__ == "__main__":
    app()
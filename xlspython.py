import typer
from typing import Optional
from module_pour_excel import *
from typing_extensions import Annotated

""" 
Enlever to les lineend qui ne servent à rien en argument avec maxrow+1
"""

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
                 column_read : Annotated[str, typer.Option(prompt = 'Enter the column containing strings ')],
                 line : Annotated[Optional[str], typer.Option(prompt = '(Optional) Enter the number of the line or press enter')] = '2'):
    """
    Fonction agissant sur un fichier.

    Vous avez un fichier xlsx dont une colonne contient des participants qui ont pu répondre plusieurs fois à un questionnaire. 
    Vous souhaitez créer un onglet par participant avec toutes les lignes qui correspondent.
    """
    fileobject = File(file) 
    fileobject.create_one_onglet_by_participant(tab,column_read,int(line))
    

@app.command()
def extractcolsheets():
    """
    Fonction agissant sur un fichier.
    """
    pass
    
@app.command()
def formulaonsheets():
    """
    Fonction agissant sur un fichier.
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
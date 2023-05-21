from unittest import TestCase,main
from module_pour_excel import *


class TestFile(TestCase):
    
    def test_open_and_copy(self):
        file = File('test.xlsx')
        self.assertNotEqual(file.writebook,None) 

        print(datetime.now().strftime("%Y-%m-%d %Hh%M"))

    def test_files_identical(self):
        """On prend deux fichiers excel, on vérifie qu'ils ont les mêmes onglets et que dans chaque onglet on a les mêmes cellules."""

        file1 = File('test_date_2023-05-20.xlsx')
        file1.sauvegarde()
        
        file2 = File('test_copie.xlsx')

        self.assertEqual(file1.sheets_name,file2.sheets_name)

        for onglet in file1.sheets_name: 
            sheet1 = file1.writebook[onglet]
            sheet2 = file2.writebook[onglet]
            for i in range(1,sheet1.max_row+1):
                for j in range(1,sheet1.max_column+1):
                    self.assertEqual(sheet1.cell(i,j).value,sheet2.cell(i,j).value)

        

class TestSheet(TestCase):
    def test_sheet_correctly_opened(self):
        """Ici je teste que l'attribut sheet de la classe sheet contient bien la bonne page correspondant à l'onglet.
        Pour cela, je génère la feuille via mes classes et par la préocédure habituelle et je regarde si la première colonne des deux fichiers se correspondent.""" 
        feuille = Sheet('test.xlsx','sheet1')

        readbook = openpyxl.load_workbook('fichiers_xls/test.xlsx', data_only=True)
        feuille2 = readbook.worksheets[0] 
        for i in range(1,feuille2.max_row):
            self.assertEqual(feuille.sheet.cell(i,1).value,feuille2.cell(i,1).value)
         
    def column_identical(self,name_file1, name_file2, index_onglet, column1,column2):
        """
        Méthode qui prend deux fichiers et regarde si à une colonne donnée les valeurs sont les mêmes
        """
        file1 = File(name_file1) 
        file2 = File(name_file2)  
        sheet1 = file1.writebook.worksheets[index_onglet] 
        sheet2 = file2.writebook.worksheets[index_onglet] 
        self.assertEqual(sheet1.max_row,sheet2.max_row) 
        
        for i in range(2,sheet1.max_row+1 ): 
            self.assertEqual(sheet1.cell(i,column1).value,sheet2.cell(i,column2).value) 

    def test_column_transform_string_in_binary(self):
        sheet = Sheet('test.xlsx','sheet1')

        sheet.column_transform_string_in_binary(12,13,'partie 1 : Vrai',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0, 13, 13) 
        sheet.column_transform_string_in_binary(14,15,'partie 2 : Vrai',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0, 15, 15) 
        sheet.column_transform_string_in_binary(16,17,'partie 3 : Vrai',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0, 17, 17)
        sheet.column_transform_string_in_binary(41,42,'Laser Interferometer Gravitational-Wave Observatory(LIGO)','virgo','Virgo',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0, 42, 42)

        self.assertEqual(type(sheet.column_transform_string_in_binary(12,13,'partie 1 : Vrai',line_end= 15,insert=False)),str)
        self.assertEqual(type(sheet.column_transform_string_in_binary(12,13,'partie 1 : Vrai',line_end= 15,security=False,insert=False)),type(None))

        sheet2 = Sheet('test.xlsx', 'Feuille2')
        #Tester l'insertion de colonne : regarder si le nb de colonnes a augmenté (pour éviter de tout modifier test.xlsx j'insère à la fin du fichier ou alors je supprime après coup ma colonne.)
        sheet2.column_transform_string_in_binary(6,7,'partie 12 : Faux',1,line_end= 15)
        self.column_identical('test.xlsx','test.xlsx', 1, 7,8)
        #A COMPLETER

    
    def test_column_security(self):
        sheet = Sheet('test.xlsx','sheet1')
        
        self.assertEqual(sheet.column_security(1), False)
        self.assertEqual(sheet.column_security(123), True)
    


class TestStr(TestCase):
    def test_transform_string_in_binary(self):
        chaine = Str('prout') 
        
        self.assertEqual(chaine.transform_string_in_binary('prout','rr'),1)
        self.assertEqual(chaine.transform_string_in_binary('rr'),0)
        self.assertEqual(chaine.transform_string_in_binary(''),0)

    def test_clean_string(self):
        chaine1 = Str('prout').clean_string()
        chaine2 = Str(' prout').clean_string()
        chaine3 = Str('prout ').clean_string()
        chaine4 = Str(' prout ').clean_string()
        chaine5 = Str('prout  ').clean_string()
        chaine6 = Str('  prout').clean_string()
        self.assertEqual(chaine1.chaine,'prout')
        self.assertEqual(chaine2.chaine,'prout') 
        self.assertEqual(chaine3.chaine,'prout') 
        self.assertEqual(chaine4.chaine,'prout') 
        self.assertEqual(chaine5.chaine,'prout') 
        self.assertEqual(chaine6.chaine,'prout') 

        

if __name__== "__main__":
    main()
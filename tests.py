from unittest import TestCase,main
from module_pour_excel import *
import xlrd

class TestFile(TestCase):
    
    def test_open_and_copy(self):
        file = File('test.xls')
        self.assertNotEqual(file.readbook,None)
        self.assertNotEqual(file.writebook,None)
        

class TestSheet(TestCase):
    def test_sheet_correctly_opened(self):
        """Ici je teste que l'attribut sheet de la classe sheet contient bien la bonne page correspondant à l'onglet.
        Pour cela, je génère la feuille via mes classes et par la préocédure habituelle et je regarde si la première colonne des deux fichiers se correspondent.""" 
        feuille = Sheet('test.xls','sheet1')

        readbook = xlrd.open_workbook('fichiers_xls/test.xls') 
        feuille2 = readbook.sheet_by_index(0) 
        for i in range(feuille2.nrows):
            self.assertEqual(feuille.sheet_read.cell_value(i,0),feuille2.cell_value(i,0)) 
         
    def column_identical(self,name_file1, name_file2, column):
        """
        Méthode qui prend deux fichiers et regarde si à une colonne donnée les valeurs sont les mêmes"""
        file1 = File(name_file1) 
        file2 = File(name_file2)  
        sheet1 = file1.readbook.sheet_by_index(0)
        sheet2 = file2.readbook.sheet_by_index(0)
        self.assertEqual(sheet1.nrows,sheet2.nrows) 
        for i in range(sheet1.nrows): 
            self.assertEqual(sheet1.cell_value(i,column),sheet2.cell_value(i,column)) 

    def test_column_transform_string_in_binary(self):
        sheet = Sheet('test.xls','sheet1')
        sheet.column_transform_string_in_binary(11,12,'partie 1 : Vrai',line_end= 14)
        self.column_identical('test.xls','test_generated.xls', 12) 
        sheet.column_transform_string_in_binary(11,12,'partie 2 : Vrai',line_end= 14)
        self.column_identical('test.xls','test_generated.xls', 14) 
        sheet.column_transform_string_in_binary(11,12,'partie 3 : Vrai',line_end= 14)
        self.column_identical('test.xls','test_generated.xls', 16)
        sheet.column_transform_string_in_binary(40,41,'Laser Interferometer Gravitational-Wave Observatory(LIGO)','virgo','Virgo',line_end= 14)
        self.column_identical('test.xls','test_generated.xls', 41)

    """
    def test_column_security(self):
        sheet = Sheet('test.xls','sheet1')
        
        self.assertEqual(sheet.column_security(1),False)
        self.assertEqual(sheet.column_security(100),True)
    """


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
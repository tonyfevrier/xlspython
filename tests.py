from unittest import TestCase, main
from module_pour_excel import *

"""IMPORTANT : pour faire refonctionner les tests il faut que je retrouve un récent test.xlsx qui 
n'est pas corrompu."""

class TestFile(TestCase):
    
    def test_open_and_copy(self):
        file = File('test_copie.xlsx')
        self.assertNotEqual(file.writebook,None) 

        print(datetime.now().strftime("%Y-%m-%d %Hh%M"))

    def test_files_identical(self):
        """On prend deux fichiers excel, on vérifie qu'ils ont les mêmes onglets et que dans chaque onglet on a les mêmes cellules."""
        self.verify_files_identical(File('test_date_2023-05-20.xlsx'),File('test_copie.xlsx'))
        """
        file1 = File('test_date_2023-05-20.xlsx')
        #file1.sauvegarde()
        
        file2 = File('test_copie.xlsx')

        self.assertEqual(file1.sheets_name,file2.sheets_name)

        for onglet in file1.sheets_name: 
            sheet1 = file1.writebook[onglet]
            sheet2 = file2.writebook[onglet]
            for i in range(1,sheet1.max_row+1):
                for j in range(1,sheet1.max_column+1):
                    self.assertEqual(sheet1.cell(i,j).value,sheet2.cell(i,j).value)
        """

    def verify_files_identical(self, file1, file2):
        self.assertEqual(file1.sheets_name,file2.sheets_name)

        for onglet in file1.sheets_name: 
            sheet1 = file1.writebook[onglet]
            sheet2 = file2.writebook[onglet]
            for i in range(1,sheet1.max_row+1):
                for j in range(1,sheet1.max_column+1):
                    self.assertEqual(sheet1.cell(i,j).value,sheet2.cell(i,j).value)

    def test_create_one_onglet_by_participant(self): 

        file = File('test_create_one_onglet_by_participant.xlsx')
        file.create_one_onglet_by_participant('Stroops_test (7)', 1)
 
        file = File('test_create_one_onglet_by_participant.xlsx')
        self.verify_files_identical(File('test_create_one_onglet_by_participant_before.xlsx'),file)
         
        for i in range(1,len(file.sheets_name)): 
            del file.writebook[file.sheets_name[i]]
        file.writebook.save(file.path + 'test_create_one_onglet_by_participant.xlsx')

    def test_extract_column_from_all_sheets(self):
        file = File('test_extract_column.xlsx')
        file.extract_column_from_all_sheets(2)

        file = File('test_extract_column.xlsx')
        self.verify_files_identical(File('test_extract_column_ref.xlsx'),file)

        del file.writebook[file.sheets_name[-1]]
        file.writebook.save(file.path + 'test_extract_column.xlsx')



class TestSheet(TestCase):
    def test_sheet_correctly_opened(self):
        """Ici je teste que l'attribut sheet de la classe sheet contient bien la bonne page correspondant à l'onglet.
        Pour cela, je génère la feuille via mes classes et par la préocédure habituelle et je regarde si la première colonne des deux fichiers se correspondent.""" 
        feuille = Sheet('test.xlsx','sheet1')

        readbook = openpyxl.load_workbook('fichiers_xls/test.xlsx', data_only=True)
        feuille2 = readbook.worksheets[0] 
        for i in range(1,feuille2.max_row):
            self.assertEqual(feuille.sheet.cell(i,1).value,feuille2.cell(i,1).value)
         
    def column_identical(self,name_file1, name_file2, index_onglet1, index_onglet2, column1,column2):
        """
        Méthode qui prend deux fichiers et regarde si à une colonne donnée les valeurs sont les mêmes
        """
        file1 = File(name_file1) 
        file2 = File(name_file2)  
        sheet1 = file1.writebook.worksheets[index_onglet1] 
        sheet2 = file2.writebook.worksheets[index_onglet2] 
        self.assertEqual(sheet1.max_row,sheet2.max_row) 
        
        for i in range(2,sheet1.max_row+1 ): 
            self.assertEqual(sheet1.cell(i,column1).value,sheet2.cell(i,column2).value) 

    def test_column_transform_string_in_binary(self):
        sheet = Sheet('test.xlsx','sheet1')  

        sheet.column_transform_string_in_binary(12,13,'partie 1 : Vrai',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0,0, 13, 13) 
        sheet.column_transform_string_in_binary(14,15,'partie 2 : Vrai',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0,0, 15, 15) 
        sheet.column_transform_string_in_binary(16,17,'partie 3 : Vrai',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0,0, 17, 17)
        sheet.column_transform_string_in_binary(41,42,'Laser Interferometer Gravitational-Wave Observatory(LIGO)','virgo','Virgo',line_end= 15,security=False,insert=False)
        self.column_identical('test.xlsx','test_generated.xlsx',0,0, 42, 42)

        self.assertEqual(type(sheet.column_transform_string_in_binary(12,13,'partie 1 : Vrai',line_end= 15,insert=False)),str)
        self.assertEqual(type(sheet.column_transform_string_in_binary(12,13,'partie 1 : Vrai',line_end= 15,security=False,insert=False)),type(None))

        sheet2 = Sheet('test.xlsx', 'Feuille2')
         
        sheet2.column_transform_string_in_binary(6,7,'partie 12 : Faux',1,line_end= 15)
        self.column_identical('test.xlsx','test.xlsx', 1, 1, 7,8)
        sheet2.sheet.delete_cols(7) #sinon à chaque lancement de test.py il insère une colonne en plus.
        sheet2.writebook.save(sheet2.path + 'test.xlsx') 

    def test_column_set_answer_in_group(self):
        sheet = Sheet('test_column_set_answer.xlsx','sheet1') 
        #groups_of_response = {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']}
        groups_of_response = {}
        for elt in ['2','5','6']:
            groups_of_response[elt] = "group1"
        
        for elt in ['7','8','9']:
            groups_of_response[elt] = "group2"

        for elt in ['1','3','4']:
            groups_of_response[elt] = "group3"
        groups_of_response['10'] = "group4"

        sheet.column_set_answer_in_group(2,3,groups_of_response,line_end= 16)

        self.column_identical('test_column_set_answer.xlsx','test_column_set_answer.xlsx',0,1,3,3)



    def test_column_security(self):
        sheet = Sheet('test.xlsx','sheet1')
        
        self.assertEqual(sheet.column_security(1), False)
        self.assertEqual(sheet.column_security(123), True)
    
    """
    def test_color_special_cases_in_sheet(self):
        sheet = Sheet('test.xlsx','Feuille3')
        sheet2 = Sheet('test.xlsx','Feuille4')

        sheet.color_special_cases_in_sheet({"partie 6 : Faux":'0000a933',0:'00ffff00',"partie 7 : Vrai":'00ff0000','Les électrons sont plus petits que les atomes':'002a6099','accuracy_Q6':'00bf0041'})
        for i in range(1,sheet.sheet.max_row+1):
            for j in range(1,sheet2.sheet.max_column+1): 
                print(i,j)
                self.assertEqual(sheet.sheet.cell(i,j).fill.fgColor.rgb,sheet2.sheet.cell(i,j).fill.fgColor.rgb)
    """
    
    def test_add_column_in_sheet_differently_sorted(self):
        sheet1 = Sheet('test.xlsx','Feuille5') 
        sheet1.add_column_in_sheet_differently_sorted(3,5,['test.xlsx','sheet1',3,[2,6]]) 
        self.column_identical('test.xlsx','test.xlsx',4,5,5,5)
        self.column_identical('test.xlsx','test.xlsx',4,5,6,6)
        sheet1.sheet.delete_cols(5,2)
        sheet1.writebook.save(sheet1.path + 'test.xlsx')
        
    def test_color_line_containing_chaines(self):
        sheet = Sheet('test.xlsx','color_line')
        sheet.color_lines_containing_chaines('0000a933','-','+')
        
    def test_column_cut_string_in_parts(self):
        sheet = Sheet('test.xlsx','cutinparts')
        sheet.column_cut_string_in_parts(2,3,';') 
        self.column_identical('test.xlsx','test.xlsx',7,8, 3, 3)
        self.column_identical('test.xlsx','test.xlsx',7,8, 4, 4)
        self.column_identical('test.xlsx','test.xlsx',7,8, 5, 5)
        self.column_identical('test.xlsx','test.xlsx',7,8, 6, 6)
        sheet.sheet.delete_cols(3,3)
        sheet.writebook.save(sheet.path + 'test.xlsx') 

    def test_delete_lines(self):
        sheet = Sheet('test.xlsx','delete_lines') 
        sheet.delete_lines(4,0)
        self.column_identical('test.xlsx','test.xlsx',9,10, 1, 1)
        self.column_identical('test.xlsx','test.xlsx',9,10, 2, 2)
        self.column_identical('test.xlsx','test.xlsx',9,10, 3, 3)
        self.column_identical('test.xlsx','test.xlsx',9,10, 4, 4)
        self.column_identical('test.xlsx','test.xlsx',9,10, 5, 5)
        self.column_identical('test.xlsx','test.xlsx',9,10, 6, 6)


class TestStr(TestCase):
    def test_transform_string_in_binary(self):
        chaine = Str('prout') 
        
        self.assertEqual(chaine.transform_string_in_binary('prout','rr'),1)
        self.assertEqual(chaine.transform_string_in_binary('rr'),0)
        self.assertEqual(chaine.transform_string_in_binary(''),0)

    def test_set_answer_in_group(self):
        chaine = Str(1)
        chaine2 = Str(9)
        """
        groups_of_response = {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']}

        self.assertEqual(chaine.set_answer_in_group(groups_of_response), "group3")
        self.assertEqual(chaine2.set_answer_in_group(groups_of_response), "group2")
        """

        groups_of_response = {}
        for elt in ['2','5','6']:
            groups_of_response[elt] = "group1"
        
        for elt in ['7','8','9']:
            groups_of_response[elt] = "group2"

        for elt in ['1','3','4']:
            groups_of_response[elt] = "group3"
        groups_of_response['10'] = "group4"
        
        self.assertEqual(chaine.set_answer_in_group(groups_of_response),"group3")
        self.assertEqual(chaine2.set_answer_in_group(groups_of_response), "group2")


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

    def test_cut_string_in_parts(self):
        chaine = Str("partie 1 : Vrai; partie 2 : Faux; partie 3 : Vrai; partie 4 : Vrai; partie 5 : Vrai")
        tuple_of_str = chaine.cut_string_in_parts(";")
        
        self.assertEqual(tuple_of_str,("partie 1 : Vrai"," partie 2 : Faux"," partie 3 : Vrai"," partie 4 : Vrai"," partie 5 : Vrai"))

    def test_convert_time_in_minutes(self):
        duration1 = Str("2 jours 2 heures")
        duration2 = Str("1 heure 25 min")
        duration3 = Str("16 min 35 s")
        
        self.assertEqual(duration1.convert_time_in_minutes(), '3000.0')
        self.assertEqual(duration2.convert_time_in_minutes(), '85.0')
        self.assertEqual(duration3.convert_time_in_minutes(), '16.58')


        

if __name__== "__main__":
    main()
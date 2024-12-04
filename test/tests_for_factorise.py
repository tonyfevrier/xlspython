from unittest import TestCase, main 
from model.model_factorise import File, OptionalNamesOfFile, OptionalNamesOfTab, MergedCellsRange 
from controller.one_file_multiple_tabs import OneFileMultipleTabsController, MultipleSameTabController
from controller.two_files import OneFileCreatedController, TwoFilesController
from utils.utils import Other, Str, ColumnDelete, ColumnInsert, LineDelete, LineInsert, TabUpdate
 
import os
 

# class TestPath(TestCase):

#     """ def test_delete_other_columns(self):
#         Pb de ce test : ne fonctionne qu'une fois car une fois les colonnes supprimées il en supprimera d'autres et le test sera faux.
#         path = Path('fichiers_xls/gathertests/')
#         path.delete_other_columns('A-F,H-J', 'test_keep_only_columns.xlsx', 'sheet1')
#         directories = [f for f in os.listdir(path.pathname) if os.path.isdir(os.path.join(path.pathname, f))]

#         for directory in directories:
#             verify_sheets_identical(Sheet('test_keep_only_columns.xlsx', 'sheet1', path.pathname + '/').sheet,
#                                     Sheet('test_keep_only_columns.xlsx', 'Feuille2', path.pathname + '/').sheet) """
    
#     def test_create_one_onglet_by_participant(self):
#         path = Path('fichiers_xls/gathertests/')
#         controler = PathControler(path)
#         #controler.apply_method_on_homononymous_files('test_cmd_ongletbyparticipant.xlsx', 'create_one_onglet_by_participant', 'test', 'A', 'divided_test_cmd_ongletbyparticipant.xlsx','fichiers_xls/gathertests/')
#         controler.create_one_onglet_by_participant('test_cmd_ongletbyparticipant.xlsx', 'test', 'A')
#         file = File('divided_test_cmd_ongletbyparticipant_before.xlsx', path.pathname)
#         file2 = File('divided_test_cmd_ongletbyparticipant.xlsx', path.pathname)
        
#         verify_files_identical(file, file2)
#         os.remove(path.pathname + 'divided_test_cmd_ongletbyparticipant.xlsx') 
        

class TestFile(TestCase):
    
    # def test_open_and_copy(self): 
    #     file = File('test_copie.xlsx') 
    #     self.assertNotEqual(file.writebook,None)   
    #     del file

    # def test_files_identical(self):
    #     """On prend deux fichiers excel, on vérifie qu'ils ont les mêmes onglets et que dans chaque onglet on a les mêmes cellules."""
    #     file1 = File('test_date_2023-05-20.xlsx')
    #     file2 = File('test_copie.xlsx')
    #     verify_files_identical(file1, file2)
    #     del file1, file2

    def test_split_one_tab_in_multiple_tabs(self): 
        file = File('test_create_one_onglet_by_participant.xlsx') 
        controler = OneFileCreatedController(file, OptionalNamesOfFile(name_of_tab_to_read='Stroops_test (7)', column_to_read='A'))
        controler.make_horodated_copy_of_a_file()
        controler.split_one_tab_in_multiple_tabs()  
        file2 = File('divided_test_create_one_onglet_by_participant.xlsx')
        verify_files_identical(File('test_create_one_onglet_by_participant_before.xlsx'),
                               file2)
        
        os.remove("fichiers_xls/divided_test_create_one_onglet_by_participant.xlsx")

    def test_extract_a_column_from_all_tabs(self):
        file = File('test_extract_column.xlsx')
        controler = OneFileMultipleTabsController(file, OptionalNamesOfFile(column_to_read='B'))
        controler.extract_a_column_from_all_tabs() 
        verify_files_identical(File('test_extract_column_ref.xlsx'),file)

        del file.writebook[file.sheets_name[-1]]
        file.writebook.save(file.path + 'test_extract_column.xlsx') 

    def test_apply_column_formula_on_all_tabs(self):
        file = File('dataset.xlsx', dataonly = False)
        controler = OneFileMultipleTabsController(file, OptionalNamesOfFile(columns_to_read=['B','C']))
        controler.apply_columns_formula_on_all_tabs()
    

    def test_gather_groups_of_multiple_columns_in_tabs_of_two_columns_containing_tags_and_values(self):
        file = File("test_gather_columns_in_one.xlsx")
        controler = OneFileMultipleTabsController(file, OptionalNamesOfFile(name_of_tab_to_read='test'))
        controler.gather_groups_of_multiple_columns_in_tabs_of_two_columns_containing_tags_and_values(['C','D','E'], ['G','H','I'])

        file = File("test_gather_columns_in_one.xlsx")
        verify_files_identical(File("test_gather_columns_in_one - ref.xlsx"), file)
 
        del file.writebook['tab_column_gathered_CDE']
        del file.writebook['tab_column_gathered_GHI']
        file.writebook.save(file.path + 'test_gather_columns_in_one.xlsx')
        del file


    def test_one_file_by_tab(self):
        file = File("test_onefile_sendmail.xlsx")
        controler = OneFileCreatedController(file)
        controler.create_one_file_by_tab()
 
        sheet1 = File("tony fevrier.xlsx", "multifiles/").writebook["Sheet"]
        sheet2 = File("Marine Moyon.xlsx", "multifiles/").writebook["Sheet"] 

        sheet1o = File("test_onefile_sendmail.xlsx").writebook["tony fevrier"] 
        sheet2o = File("test_onefile_sendmail.xlsx").writebook["Marine Moyon"]

        verify_sheets_identical(sheet1, sheet1o)
        verify_sheets_identical(sheet2, sheet2o) 

    def test_merge_cells_on_all_tabs(self): 
        file = File("test_merging.xlsx")
        controler = OneFileMultipleTabsController(file)
        controler.merge_cells_on_all_tabs(MergedCellsRange('C', 'D', 12, 15))

        #voir comment tester le fait qu'une cellule est mergée : comprendre l'objet mergedcells
        """ for tab in file1.sheets_name:
            sheet = file1.writebook[tab]
            mergedcells = sheet.merged_cells
            print(mergedcells.ranges, type(mergedcells))
            self.assertEqual('C5' in mergedcells.ranges,True)
            self.assertIn(sheet['C6'],mergedcells)
            self.assertIn(sheet['C7'],mergedcells)
            self.assertIn(sheet['D5'],mergedcells)
            self.assertIn(sheet['D6'],mergedcells) """
        
    def test_apply_cell_formula_on_all_sheets(self):
        file = File("test_merging.xlsx")
        controler = OneFileMultipleTabsController(file)
        controler.apply_cells_formula_on_all_tabs('A10','B10','C10')

        for tab in file.sheets_name[1:]:
            sheet = file.writebook[tab]
            self.assertEqual(sheet['A10'].value, file.writebook[file.sheets_name[0]]['A10'].value)
            self.assertEqual(sheet['B10'].value, file.writebook[file.sheets_name[0]]['B10'].value)
            self.assertEqual(sheet['C10'].value, file.writebook[file.sheets_name[0]]['C10'].value)
    
    def test_check_linenumber_of_tabs(self):
        file = File('test.xlsx')
        controler = OneFileMultipleTabsController(file)
        tabs = controler.list_tabs_with_different_number_of_lines(14)
        self.assertListEqual(tabs, ['cutinparts', 'cutinpartsbis', 'delete_lines', 'delete_lines_bis'])

    def test_extract_cells_from_all_tabs(self):
        file = File('test_extract_cells_from_all_sheets.xlsx')
        controler = OneFileCreatedController(file)
        controler.extract_cells_from_all_tabs('C7','D7','C8','D8') 
        file2 = File('gathered_data_test_extract_cells_from_all_sheets.xlsx')
        verify_sheets_identical(file2.writebook['Sheet'], File('test_extract_cells_from_all_sheets - after.xlsx').writebook['gathered_data']) 

#     def test_apply_same_method_on_all_sheets(self):
#         file = File("test_method_on_all_sheets.xlsx")
#         controler = FileControler(file)
#         controler.apply_method_on_some_sheets(file.sheets_name, 'column_transform_string_in_binary', 'C', 'D', "partie 7 : Faux")
#         ref_sheet = File('test_method_on_all_sheets - ref.xlsx').writebook['Feuille5']
#         for sheet in file.sheets_name:
#             actual_sheet = file.writebook[sheet]
#             verify_sheets_identical(actual_sheet, ref_sheet)
#             actual_sheet.delete_cols(4)
#             file.writebook.save(file.path + 'test_method_on_all_sheets.xlsx') 


# class TestSheet(TestCase, Other):
#     def test_sheet_correctly_opened(self):
#         """Ici je teste que l'attribut sheet de la classe sheet contient bien la bonne page correspondant à l'onglet.
#         Pour cela, je génère la feuille via mes classes et par la préocédure habituelle et je regarde si la première colonne des deux fichiers se correspondent.""" 
#         feuille = File('test.xlsx').writebook['sheet1']

#         readbook = openpyxl.load_workbook('fichiers_xls/test.xlsx', data_only=True)
#         feuille2 = readbook.worksheets[0] 
#         for i in range(1,feuille2.max_row):
#             self.assertEqual(feuille.cell(i,1).value,feuille2.cell(i,1).value)
         
    def column_identical(self,name_file1, name_file2, index_onglet1, index_onglet2, column1,column2):
        """
        Méthode qui prend deux fichiers et regarde si à une colonne donnée les valeurs sont les mêmes
        """
        file1 = File(name_file1) 
        file2 = File(name_file2)  
        sheet1 = file1.writebook.worksheets[index_onglet1] 
        sheet2 = file2.writebook.worksheets[index_onglet2] 
        self.assertEqual(sheet1.max_row,sheet2.max_row) 
        
        for i in range(2,sheet1.max_row + 1): 
            self.assertEqual(sheet1.cell(i,column1).value,sheet2.cell(i,column2).value)

#     def test_column_transform_string_in_binary(self): 
#         file = File('test.xlsx')
#         controler = FileControler(file)
#         sheet2 = file.writebook['Feuille2']
        
#         controler.apply_method_on_some_sheets(['Feuille2'], 'column_transform_string_in_binary', 'F','G','partie 12 : Faux',1) 
#         #controler.column_transform_string_in_binary('Feuille2','F','G','partie 12 : Faux',1)
#         self.column_identical('test.xlsx','test.xlsx', 1, 1, 7,8)
#         sheet2.delete_cols(7) #sinon à chaque lancement de test.py il insère une colonne en plus.
#         file.writebook.save(file.path + 'test.xlsx') 

#     def test_column_set_answer_in_group(self):
#         file = File('test_column_set_answer.xlsx')
#         sheet = file.writebook['sheet1']
#         controler = FileControler(file)  
        
#         groups_of_response = {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']}  

#         controler.apply_method_on_some_sheets(['sheet1'], 'column_set_answer_in_group', 'B','C', groups_of_response) 
#         #controler.column_set_answer_in_group('sheet1','B','C',groups_of_response)
 
#         self.column_identical('test_column_set_answer.xlsx','test_column_set_answer.xlsx',0,1,3,3)
#         self.column_identical('test_column_set_answer.xlsx','test_column_set_answer.xlsx',0,1,4,4)

#         sheet.delete_cols(3)
#         file.updateCellFormulas(sheet,False,'column',['C'])
#         file.writebook.save(file.path + 'test_column_set_answer.xlsx') 
        

#     def test_column_security(self):
#         file = File('test.xlsx')
#         controler = FileControler(file)
#         sheet = file.writebook['sheet1'] 
        
#         self.assertEqual(controler.column_security(sheet, 1), False)
#         self.assertEqual(controler.column_security(sheet, 123), True)
    
#     """
#     def test_color_special_cases_in_sheet(self):
#         sheet = Sheet('test.xlsx','Feuille3')
#         sheet2 = Sheet('test.xlsx','Feuille4')

#         sheet.color_special_cases_in_sheet({"partie 6 : Faux":'0000a933',0:'00ffff00',"partie 7 : Vrai":'00ff0000','Les électrons sont plus petits que les atomes':'002a6099','accuracy_Q6':'00bf0041'})
#         for i in range(1,sheet.sheet.max_row+1):
#             for j in range(1,sheet2.sheet.max_column+1): 
#                 print(i,j)
#                 self.assertEqual(sheet.sheet.cell(i,j).fill.fgColor.rgb,sheet2.sheet.cell(i,j).fill.fgColor.rgb)
#     """
    
    def test_add_column_in_sheet_differently_sorted(self):
        file1 = File('test.xlsx', dataonly=True)
        file2 = File('test.xlsx')
        sheet2 = file2.writebook['Feuille5']

        controler = TwoFilesController(file1, file2, 'sheet1', 'Feuille5', 'C', 'C') 
        controler.copy_columns_in_a_tab_differently_sorted(['B','F'], 'E') 

        self.column_identical('test.xlsx','test.xlsx',4,5,5,5)
        self.column_identical('test.xlsx','test.xlsx',4,5,6,6)
        self.column_identical('test.xlsx','test.xlsx',4,5,8,8)
        
        sheet2.delete_cols(5,2)
        TabUpdate(sheet2, ColumnDelete(['E','F'])).update_cells_formulas()
        file2.writebook.save(file2.path + 'test.xlsx')

    def test_color_column(self):
        file = File('test.xlsx')
        controler = MultipleSameTabController(file, 
                                              optional_names_of_file=OptionalNamesOfFile(names_of_tabs_to_modify=['cutinpartsbis']))  
        controler.tab_controller.optional_names_of_tab.column_to_read = 'D'#OptionalNamesOfTab(column_to_read='D')
        controler.apply_method_on_some_tabs('color_cases_in_column', 
                                            {' partie 2 : Vrai':'0000a933'}) 
         
    def test_color_cases_in_sheet(self):
        file = File('test.xlsx')
        controler = MultipleSameTabController(file, 
                                              optional_names_of_file=OptionalNamesOfFile(names_of_tabs_to_modify=['cutinpartsbis']))   
        controler.apply_method_on_some_tabs('color_cases_in_sheet', 
                                            {'partie 1 : Vrai':'0000a933', 'Abbas':'0000a933'}) 
        
#     def test_color_line_containing_chaines(self):
#         file = File('test.xlsx')
#         controler = FileControler(file) 
#         controler.apply_method_on_some_sheets(['color_line'], 'color_lines_containing_chaines', '0000a933', '-', '+')
#         #controler.color_lines_containing_chaines('color_line','0000a933','-','+')
        
#     def test_column_cut_string_in_parts(self):
#         file = File('test.xlsx')
#         controler = FileControler(file)
#         sheet = file.writebook['cutinparts']  
#         controler.apply_method_on_some_sheets(['cutinparts'], 'column_cut_string_in_parts', 'B','C',';')
#         #controler.column_cut_string_in_parts('cutinparts','B','C',';') 
#         self.column_identical('test.xlsx','test.xlsx',7,8, 3, 3)
#         self.column_identical('test.xlsx','test.xlsx',7,8, 4, 4)
#         self.column_identical('test.xlsx','test.xlsx',7,8, 5, 5)
#         self.column_identical('test.xlsx','test.xlsx',7,8, 6, 6)
#         sheet.delete_cols(3,3)
#         file.writebook.save(file.path + 'test.xlsx') 

#     def test_delete_lines(self):
#         file = File('test.xlsx')
#         controler = FileControler(file)  
#         controler.apply_method_on_some_sheets(['delete_lines'], 'delete_lines_containing_str', 'D', '0')
#         controler.apply_method_on_some_sheets(['delete_lines'], 'delete_lines_containing_str', 'D', 'p a')
#         #controler.delete_lines_containing_str('delete_lines', 'D', '0')
#         #controler.delete_lines_containing_str('delete_lines', 'D','p a')
#         self.column_identical('test.xlsx','test.xlsx',9,10, 1, 1)
#         self.column_identical('test.xlsx','test.xlsx',9,10, 2, 2)
#         self.column_identical('test.xlsx','test.xlsx',9,10, 3, 3)
#         self.column_identical('test.xlsx','test.xlsx',9,10, 4, 4)
#         self.column_identical('test.xlsx','test.xlsx',9,10, 5, 5)
#         self.column_identical('test.xlsx','test.xlsx',9,10, 6, 6)

#     def test_delete_lines_with_formulas(self):
#         file = File('listing_par_etape - Copie.xlsx')
#         controler = FileControler(file) 
#         controler.apply_method_on_some_sheets(['Feuil1'], 'delete_lines_containing_str', 'B', 'pas consenti')
#         #controler.delete_lines_containing_str('Feuil1', 'B', 'pas consenti') 
#         self.column_identical('listing_par_etape - Copie.xlsx','listing_par_etape - Copie.xlsx',0, 1, 2, 2)
#         self.column_identical('listing_par_etape - Copie.xlsx','listing_par_etape - Copie.xlsx',0, 1, 10, 10) 

#     def test_delete_doublons(self): 
#         file = File('test_doublons.xlsx')
#         controler = FileControler(file)
#         sheet1 = file.writebook['sheet1']  
#         sheet2 = file.writebook['Feuille2']  
#         controler.apply_method_on_some_sheets(['sheet1'], 'delete_doublons', 'C', color = True)
#         #controler.delete_doublons('sheet1', 'C', color = True)
#         verify_sheets_identical(sheet1,sheet2)

#     def test_create_one_column_by_QCM_answer(self):
#         file = File('test_create_one_column.xlsx')
#         controler = FileControler(file)
#         sheet = file.writebook['sheet1']  

#         controler.apply_method_on_some_sheets(['sheet1'], 'create_one_column_by_QCM_answer', 'D','E',['OUI', 'NON'], 'Alain', 'Henri', 'Tony', 'Dulcinée')
#         #controler.create_one_column_by_QCM_answer('sheet1','D','E',['OUI', 'NON'], 'Alain', 'Henri', 'Tony', 'Dulcinée') 
        
#         self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 5, 5)
#         self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 6, 6) 
#         self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 7, 7) 
#         self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 8, 8) 

#         sheet.delete_cols(4,4)

#     def test_gather_multiple_answers(self):
#         file = File('testongletbyparticipant.xlsx')
#         controler = FileControler(file)
#         #sheet = file.writebook['test']    
#         controler.gather_multiple_answers('test','A','B')
#         #del sheet

#         file2 = File('testongletbyparticipant-result.xlsx')
#         sheet1, sheet2 = file.writebook['severalAnswers'], file2.writebook['Feuille2'] 
#         verify_sheets_identical(sheet1, sheet2)
#         del sheet1, sheet2
        
#         #del file.writebook[file.sheets_name[-1]]
#         file.writebook.save(file.path + 'testongletbyparticipant.xlsx')
#         del file

#     def test_give_names_of_maximum(self):
#         file = File('test_give_names.xlsx')
#         controler = FileControler(file)
#         sheet = file.writebook['sheet1']  

#         controler.apply_method_on_some_sheets(['sheet1'], 'give_names_of_maximum',  ['A', 'B', 'C'], 'D')
#         #controler.give_names_of_maximum('sheet1', ['A', 'B', 'C'], 'D') 

#         verify_sheets_identical(sheet, file.writebook['Feuille2'])

#         sheet.delete_cols(4)
#         file.writebook.save(file.path + 'test_give_names.xlsx') 
        
#     """ def test_delete_other_columns(self):
#         # Fonctionnel une fois
#         sheet = Sheet('test_keep_only_columns.xlsx','sheet1')
#         sheet.delete_other_columns('B-H,L,M,AI-AJ')

#         verify_sheets_identical(sheet, Sheet('test_keep_only_columns.xlsx','Feuille2')) """

#     def test_column_get_part_of_str(self):
#         file = File('test_colgetpartofstr.xlsx')
#         controler = FileControler(file)
#         sheet = file.writebook['Feuille2'] 
#         controler.apply_method_on_some_sheets(['Feuille2'], 'column_get_part_of_str', 'C','D','_',0)
#         controler.apply_method_on_some_sheets(['Feuille2'], 'column_get_part_of_str', 'F','G',';',1)
        
#         #controler.column_get_part_of_str('Feuille2','C','D','_',0)
#         #controler.column_get_part_of_str('Feuille2','F','G',';',1)
#         verify_sheets_identical(sheet, file.writebook['expected'])
#         sheet.delete_cols(7)
#         sheet.delete_cols(4)
#         file.writebook.save(file.path + 'test_colgetpartofstr.xlsx') 

#     def test_map_two_columns_to_a_third_column(self):
#         file = File('test_maptwocolumns.xlsx')
#         controler = FileControler(file)
#         sheet = file.writebook['Feuille2']  
#         controler.apply_method_on_some_sheets(['Feuille2'], 'map_two_columns_to_a_third_column',  ['B', 'C'], 'D', {'cat1':['prime','1'], 'cat2':['probe','2']})

#         #controler.map_two_columns_to_a_third_column('Feuille2', ['B', 'C'], 'D', {'cat1':['prime','1'], 'cat2':['probe','2']})
#         verify_sheets_identical(sheet, file.writebook['expected'])
#         sheet.delete_cols(4)
#         file.writebook.save(file.path + 'test_maptwocolumns.xlsx')        


# class TestStr(TestCase, Other):
#     def test_transform_string_in_binary(self):
#         chaine = Str('prout') 
        
#         self.assertEqual(chaine.transform_string_in_binary('prout','rr'),1)
#         self.assertEqual(chaine.transform_string_in_binary('rr'),0)
#         self.assertEqual(chaine.transform_string_in_binary(''),0)

#     def test_set_answer_in_group(self):
#         chaine = Str(1)
#         chaine2 = Str(9)
        
#         groups_of_response = {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']}
#         reversed_group = self.reverse_dico_for_set_answer_in_group(groups_of_response)
 
#         """ 
#         groups_of_response = {}
#         for elt in ['2','5','6']:
#             groups_of_response[elt] = "group1"
        
#         for elt in ['7','8','9']:
#             groups_of_response[elt] = "group2"

#         for elt in ['1','3','4']:
#             groups_of_response[elt] = "group3"
#         groups_of_response['10'] = "group4" """
        
#         self.assertEqual(chaine.set_answer_in_group(reversed_group),"group3")
#         self.assertEqual(chaine2.set_answer_in_group(reversed_group), "group2")


#     def test_clean_string(self):
#         chaine1 = Str('prout').clean_string()
#         chaine2 = Str(' prout').clean_string()
#         chaine3 = Str('prout ').clean_string()
#         chaine4 = Str(' prout ').clean_string()
#         chaine5 = Str('prout  ').clean_string()
#         chaine6 = Str('  prout').clean_string()
#         self.assertEqual(chaine1.chaine,'prout')
#         self.assertEqual(chaine2.chaine,'prout') 
#         self.assertEqual(chaine3.chaine,'prout') 
#         self.assertEqual(chaine4.chaine,'prout') 
#         self.assertEqual(chaine5.chaine,'prout') 
#         self.assertEqual(chaine6.chaine,'prout') 

#     def test_cut_string_in_parts(self):
#         chaine = Str("partie 1 : Vrai; partie 2 : Faux; partie 3 : Vrai; partie 4 : Vrai; partie 5 : Vrai")
#         tuple_of_str = chaine.cut_string_in_parts(";")
        
#         self.assertEqual(tuple_of_str,("partie 1 : Vrai"," partie 2 : Faux"," partie 3 : Vrai"," partie 4 : Vrai"," partie 5 : Vrai"))

#     def test_convert_time_in_minutes(self):
#         duration1 = Str("2 jours 2 heures")
#         duration2 = Str("1 heure 25 min")
#         duration3 = Str("16 min 35 s")
        
#         self.assertEqual(duration1.convert_time_in_minutes(), '3000,0')
#         self.assertEqual(duration2.convert_time_in_minutes(), '85,0')
#         self.assertEqual(duration3.convert_time_in_minutes(), '16,58')

#     def test_columns_from_string(self):
#         self.assertListEqual(Str.columns_from_strings("C-E,H,J-L"), ['C','D','E','H','J','K','L'])

#     def test_listFromColumnsStrings(self):
#         self.assertListEqual(Str.listFromColumnsStrings("C-E,H,J-L", "D,G","H-K"),[['C','D','E','H','J','K','L'],['D','G'],['H','I','J','K']])

#     def test_range_Letter(self):
#         self.assertListEqual(Str.rangeLetter('D-H'), ['D','E','F','G','H'])

    def testUpdateOneFormulaForOneInsertion(self):
        formula = LineInsert(['2'])._update_a_cell("SI(J10+K$1+L$3)")
        self.assertEqual(formula, "SI(J11+K$1+L$4)")

        formula = LineDelete(['11'])._update_a_cell("SI(J12+K$1+L$3)")
        self.assertEqual(formula, "SI(J11+K$1+L$3)")

        formula = LineDelete(['13'])._update_a_cell("SI(J12+K$1+L$3)")
        self.assertEqual(formula, "SI(J12+K$1+L$3)")

        formula = ColumnInsert(['C'])._update_a_cell("SI(J10+K$1+L$3)")
        self.assertEqual(formula, "SI(K10+L$1+M$3)")

        formula = ColumnDelete(['C'])._update_a_cell("SI(J10+K$1+L$3)")
        self.assertEqual(formula, "SI(I10+J$1+K$3)")

    def testUpdateOneFormula(self):
        formula = LineInsert(['2', '5'])._update_a_cell("SI(J10+K$1+L$3)") 
        self.assertEqual(formula, "SI(J12+K$1+L$4)")
 
        formula = ColumnInsert(['C', 'D','E', 'F'])._update_a_cell("SI(D10+E$1)") 
        self.assertEqual(formula, "SI(H10+I$1)")
        
def verify_files_identical(file1, file2):
    testcase = TestCase()
    testcase.assertEqual(file1.sheets_name,file2.sheets_name)

    for onglet in file1.sheets_name: 
        sheet1 = file1.writebook[onglet]
        sheet2 = file2.writebook[onglet]
        for i in range(1,sheet1.max_row+1):
            for j in range(1,sheet1.max_column+1):
                testcase.assertEqual(sheet1.cell(i,j).value,sheet2.cell(i,j).value)

def verify_sheets_identical(sheet1, sheet2):  
    testcase = TestCase()
    testcase.assertEqual(sheet1.max_row,sheet2.max_row)
    testcase.assertEqual(sheet1.max_column,sheet2.max_column)

    for i in range(1,sheet1.max_row+1):
        for j in range(1,sheet1.max_column+1):
            testcase.assertEqual(sheet1.cell(i,j).value,sheet2.cell(i,j).value)

if __name__== "__main__":
    main()
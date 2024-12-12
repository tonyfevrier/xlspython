from unittest import TestCase, main 
from model.model_factorise import File, FileOptions, TabOptions, MergedCellsRange
from controller.one_file_one_tab import ColorTabController, DeleteController, InsertController 
from controller.one_file_multiple_tabs import OneTabCreatedController, MultipleSameTabController, EvenTabsController
from controller.two_files import OneFileCreatedController, TwoFilesController
from utils.utils import Other, Str, ColumnDelete, ColumnInsert, LineDelete, LineInsert, TabUpdateFormula
 
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
        controler = OneFileCreatedController(file, FileOptions(name_of_tab_to_read='Stroops_test (7)', column_to_read='A'))
        #controler.make_horodated_copy_of_a_file()
        controler.split_one_tab_in_multiple_tabs()  
        file2 = File('divided_test_create_one_onglet_by_participant.xlsx')
        verify_files_identical(File('test_create_one_onglet_by_participant_before.xlsx'),
                               file2)
        
        os.remove("fichiers_xls/divided_test_create_one_onglet_by_participant.xlsx")

    def test_extract_a_column_from_all_tabs(self):
        file = File('test_extract_column.xlsx')
        controler = OneTabCreatedController(file, FileOptions(column_to_read='B'))
        controler.extract_a_column_from_all_tabs() 
        verify_files_identical(File('test_extract_column_ref.xlsx'),file)

        del file.writebook[file.sheets_name[-1]]
        file.writebook.save(file.path + 'test_extract_column.xlsx') 

    def test_apply_column_formula_on_all_tabs(self):
        file = File('dataset.xlsx', dataonly = False)
        controler = EvenTabsController(file, FileOptions(columns_to_read=['B','C']))
        controler.apply_columns_formula_on_all_tabs()
    

    def test_gather_groups_of_multiple_columns_in_tabs_of_two_columns_containing_tags_and_values(self):
        file = File("test_gather_columns_in_one.xlsx")
        controler = OneTabCreatedController(file, FileOptions(name_of_tab_to_read='test'))
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
        controler = EvenTabsController(file)
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
        controler = EvenTabsController(file)
        controler.apply_cells_formula_on_all_tabs('A10','B10','C10')

        for tab in file.sheets_name[1:]:
            sheet = file.writebook[tab]
            self.assertEqual(sheet['A10'].value, file.writebook[file.sheets_name[0]]['A10'].value)
            self.assertEqual(sheet['B10'].value, file.writebook[file.sheets_name[0]]['B10'].value)
            self.assertEqual(sheet['C10'].value, file.writebook[file.sheets_name[0]]['C10'].value)
    
    def test_check_linenumber_of_tabs(self):
        file = File('test.xlsx')
        controler = EvenTabsController(file)
        tabs = controler.list_tabs_with_different_number_of_lines(14)
        self.assertListEqual(tabs, ['cutinparts', 'cutinpartsbis', 'delete_lines', 'delete_lines_bis', 'time_min', 'time_min_expected'])

    def test_extract_cells_from_all_tabs(self):
        file = File('test_extract_cells_from_all_sheets.xlsx')
        controler = OneFileCreatedController(file)
        controler.extract_cells_from_all_tabs('C7','D7','C8','D8') 
        file2 = File('gathered_data_test_extract_cells_from_all_sheets.xlsx')
        verify_sheets_identical(file2.writebook['Sheet'], File('test_extract_cells_from_all_sheets - after.xlsx').writebook['gathered_data']) 

#     def test_apply_same_method_on_all_sheets(self):
#         file = File("test_method_on_all_sheets.xlsx")
#         controler = FileControler(file)
#         controler.apply_method_on_some_tabs(file.sheets_name, 'column_transform_string_in_binary', 'C', 'D', "partie 7 : Faux")
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

    def test_column_transform_string_in_binary(self): 
        file = File('test.xlsx')
        controler = MultipleSameTabController(file,
                                              tab_controller=InsertController(file, tab_options=TabOptions(column_to_read='F', column_to_write='G')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['Feuille2']))
        
        controler.apply_method_on_some_tabs('transform_string_in_binary_in_column', 'partie 12 : Faux', 1) 
        self.column_identical('test.xlsx','test.xlsx', 1, 1, 7,8)

        sheet2 = file.writebook['Feuille2']
        sheet2.delete_cols(7) #sinon à chaque lancement de test.py il insère une colonne en plus.
        file.writebook.save(file.path + 'test.xlsx') 

    def test_convert_time_in_minutes_in_columns(self): 
        file = File('test.xlsx')
        controler = MultipleSameTabController(file,
                                              tab_controller=InsertController(file, tab_options=TabOptions(column_to_read='E', column_to_write='F')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['time_min']))
        
        controler.apply_method_on_some_tabs('convert_time_in_minutes_in_columns') 
        self.column_identical('test.xlsx','test.xlsx', 11, 12, 6, 6)

        sheet2 = file.writebook['time_min']
        sheet2.delete_cols(6)
        file.writebook.save(file.path + 'test.xlsx') 

    def test_column_set_answer_in_group(self):
        file = File('test_column_set_answer.xlsx')
        controler = MultipleSameTabController(file,
                                              tab_controller=InsertController(file, tab_options=TabOptions(column_to_read='B', column_to_write='C')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['sheet1']))  
        
        map_groups_to_answers = {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']}  

        controler.apply_method_on_some_tabs('insert_group_associated_with_answer', map_groups_to_answers)  
 
        self.column_identical('test_column_set_answer.xlsx','test_column_set_answer.xlsx',0,1,3,3)
        self.column_identical('test_column_set_answer.xlsx','test_column_set_answer.xlsx',0,1,4,4)

        sheet = file.writebook['sheet1']
        sheet.delete_cols(3)

        modification_object = ColumnDelete(['C'])
        controler.tab_controller._update_cell_formulas(modification_object) 
        file.writebook.save(file.path + 'test_column_set_answer.xlsx') 
        

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
        TabUpdateFormula(ColumnDelete(['E','F'])).update_cells_formulas(sheet2)
        file2.writebook.save(file2.path + 'test.xlsx')

    def test_color_column(self):
        file = File('test.xlsx')
        controler = MultipleSameTabController(file,
                                              ColorTabController(file, tab_options=TabOptions(column_to_read='D')), 
                                              file_options=FileOptions(names_of_tabs_to_modify=['cutinpartsbis']))  
        #controler.tab_controller.tab_options.column_to_read = 'D'#TabOptions(column_to_read='D')
        controler.apply_method_on_some_tabs('color_cases_in_column', 
                                            {' partie 2 : Vrai':'0000a933'}) 
         
    def test_color_cases_in_sheet(self):
        file = File('test.xlsx')
        controler = MultipleSameTabController(file,
                                              ColorTabController(file, tab_options=TabOptions()), 
                                              file_options=FileOptions(names_of_tabs_to_modify=['cutinpartsbis']))   
        controler.apply_method_on_some_tabs('color_cases_in_sheet', 
                                            {'partie 1 : Vrai':'0000a933', 'Abbas':'0000a933'}) 
        
    def test_color_line_containing_chaines(self):
        file = File('test.xlsx')
        controler = MultipleSameTabController(file, 
                                              ColorTabController(file, color='0000a933'),
                                              file_options=FileOptions(names_of_tabs_to_modify=['color_line']))    
        controler.apply_method_on_some_tabs('color_lines_containing_strings', '-', '+') 
        
    def test_column_cut_string_in_parts(self):
        file = File('test.xlsx')
        controler = MultipleSameTabController(file,
                                              InsertController(file, tab_options=TabOptions(column_to_write='C')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['cutinparts']))

        controler.apply_method_on_some_tabs('insert_splitted_strings_of', 'B',';')
        sheet = file.writebook['cutinparts'] 
        self.column_identical('test.xlsx','test.xlsx',7,8, 3, 3)
        self.column_identical('test.xlsx','test.xlsx',7,8, 4, 4)
        self.column_identical('test.xlsx','test.xlsx',7,8, 5, 5)
        self.column_identical('test.xlsx','test.xlsx',7,8, 6, 6)
        sheet.delete_cols(3,3)
        file.writebook.save(file.path + 'test.xlsx') 

    def test_delete_lines(self):
        file = File('test.xlsx')
        controler = MultipleSameTabController(file, tab_controller=DeleteController(file),
                                              file_options=FileOptions(names_of_tabs_to_modify=['delete_lines']))  
        controler.apply_method_on_some_tabs('delete_lines_containing_strings_in_given_column', 'D', '0')
        controler.apply_method_on_some_tabs('delete_lines_containing_strings_in_given_column', 'D', 'p a') 
        self.column_identical('test.xlsx','test.xlsx',9,10, 1, 1)
        self.column_identical('test.xlsx','test.xlsx',9,10, 2, 2)
        self.column_identical('test.xlsx','test.xlsx',9,10, 3, 3)
        self.column_identical('test.xlsx','test.xlsx',9,10, 4, 4)
        self.column_identical('test.xlsx','test.xlsx',9,10, 5, 5)
        self.column_identical('test.xlsx','test.xlsx',9,10, 6, 6)

    def test_delete_lines_with_formulas(self):
        file = File('listing_par_etape - Copie.xlsx')
        controler = MultipleSameTabController(file, tab_controller=DeleteController(file),
                                              file_options=FileOptions(names_of_tabs_to_modify=['Feuil1']))   
        controler.apply_method_on_some_tabs('delete_lines_containing_strings_in_given_column', 'B', 'pas consenti') 
        self.column_identical('listing_par_etape - Copie.xlsx','listing_par_etape - Copie.xlsx',0, 1, 2, 2)
        self.column_identical('listing_par_etape - Copie.xlsx','listing_par_etape - Copie.xlsx',0, 1, 10, 10) 

    def test_delete_doublons(self): 
        file = File('test_doublons.xlsx')
        controler = MultipleSameTabController(file, 
                                              tab_controller=DeleteController(file), 
                                              file_options=FileOptions(names_of_tabs_to_modify=['sheet2', 'sheet1', 'sheet3']))
        sheet_result = file.writebook['result']  
        controler.apply_method_on_some_tabs('delete_twins_lines_and_color_last_twin', 'C', color = 'FFFFFF00') 
        sheet1 = file.writebook['sheet1']  
        sheet2 = file.writebook['sheet2']  
        sheet3 = file.writebook['sheet3']  
        verify_sheets_identical(sheet1, sheet_result)
        verify_sheets_identical(sheet2, sheet_result)
        verify_sheets_identical(sheet3, sheet_result)

    def test_create_one_column_by_QCM_answer(self):
        file = File('test_create_one_column.xlsx')
        controler = MultipleSameTabController(file, 
                                              InsertController(file, tab_options=TabOptions(column_to_read='D', column_to_write='E')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['sheet1'])) 

        controler.apply_method_on_some_tabs('fill_one_column_by_QCM_answer', 'Alain', 'Henri', 'Tony', 'Dulcinée') 
        
        self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 5, 5)
        self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 6, 6) 
        self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 7, 7) 
        self.column_identical('test_create_one_column.xlsx','test_create_one_column.xlsx',0, 1, 8, 8) 

        sheet = file.writebook['sheet1'] 
        sheet.delete_cols(4,4)

    def test_gather_multiple_answers(self):
        file = File('testongletbyparticipant.xlsx')
        controler = OneTabCreatedController(file, file_options=FileOptions(name_of_tab_to_read='test'))  
        controler.gather_multiple_answers('A','B') 

        file2 = File('testongletbyparticipant-result.xlsx')
        sheet1, sheet2 = file.writebook['severalAnswers'], file2.writebook['Feuille2'] 
        verify_sheets_identical(sheet1, sheet2)
        del sheet1, sheet2

        file.writebook.save(file.path + 'testongletbyparticipant.xlsx')
        del file

    def test_give_names_of_maximum(self):
        file = File('test_give_names.xlsx')
        controler = MultipleSameTabController(file,
                                              tab_controller=InsertController(file, tab_options=TabOptions(columns_to_read=['A', 'B', 'C'], column_to_write='D')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['sheet1'])) 

        controler.apply_method_on_some_tabs('insert_tags_of_maximum_of_column_list') 

        sheet = file.writebook['sheet1']  
        verify_sheets_identical(sheet, file.writebook['Feuille2'])
        sheet.delete_cols(4)
        file.writebook.save(file.path + 'test_give_names.xlsx') 
        
    """ def test_delete_other_columns(self):
        # Fonctionnel une fois
        file = File('test_keep_only_columns.xlsx')
        controller = DeleteController(file, 'sheet1')
        controller.delete_other_columns('A-C,D-K')

        verify_sheets_identical(file.get_tab_by_name('sheet1'), File('test_keep_only_columns.xlsx').get_tab_by_name('Feuille2')) """

    def test_delete_columns(self):
        # Fonctionnel une fois
        file = File('test_keep_only_columns.xlsx')
        controller = MultipleSameTabController(file, DeleteController(file, 'sheet1'), 
                                  file_options=FileOptions(names_of_tabs_to_modify=['sheet1']))  
                                   
        controller.apply_method_on_some_tabs('delete_columns', 'L,M,N-V')

        verify_sheets_identical(file.get_tab_by_name('sheet1'), File('test_keep_only_columns.xlsx').get_tab_by_name('Feuille2'))

    def test_column_get_part_of_str(self):
        file = File('test_colgetpartofstr.xlsx')
        controler = MultipleSameTabController(file, tab_controller=InsertController(file, tab_options=TabOptions(column_to_read='C', column_to_write='D')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['Feuille2']))

        controler.apply_method_on_some_tabs('write_piece_of_string_in_column', '_', 0)

        controler.tab_controller.tab_options = TabOptions(column_to_read='F', column_to_write='G')
        controler.apply_method_on_some_tabs('write_piece_of_string_in_column', ';', 1)
        
        sheet = file.writebook['Feuille2'] 
        verify_sheets_identical(sheet, file.writebook['expected'])
        sheet.delete_cols(7)
        sheet.delete_cols(4)
        file.writebook.save(file.path + 'test_colgetpartofstr.xlsx') 

    def test_map_two_columns_to_a_third_column(self):
        file = File('test_maptwocolumns.xlsx')
        controler = MultipleSameTabController(file,
                                              tab_controller=InsertController(file, tab_options=TabOptions(columns_to_read=['B', 'C'], column_to_write='D')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['Feuille2']))
         
        controler.apply_method_on_some_tabs('map_two_columns_to_a_third_column', {'cat1':['prime','1'], 'cat2':['probe','2']})
        
        sheet = file.writebook['Feuille2']  
        verify_sheets_identical(sheet, file.writebook['expected'])
        sheet.delete_cols(4)
        file.writebook.save(file.path + 'test_maptwocolumns.xlsx')  

    def test_insert_column_for_prime_probe_congruence(self):      
        file = File('test_prime_probe.xlsx')
        controler = MultipleSameTabController(file,
                                              tab_controller=InsertController(file, tab_options=TabOptions(columns_to_read=['B', 'C', 'D'], column_to_write='E')),
                                              file_options=FileOptions(names_of_tabs_to_modify=['Feuille2']))
         
        controler.apply_method_on_some_tabs('insert_column_for_prime_probe_congruence')
        
        sheet = file.writebook['Feuille2']  
        verify_sheets_identical(sheet, file.writebook['expected'])
        sheet.delete_cols(5)
        file.writebook.save(file.path + 'test_prime_probe.xlsx')  

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
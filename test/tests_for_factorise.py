from unittest import TestCase, main 
from model.model_factorise import File, FileOptions, TabOptions, MergedCellsRange, Path
from controller.one_file_one_tab import ColorTabController, DeleteController, InsertController 
from controller.one_file_multiple_tabs import OneTabCreatedController, MultipleSameTabController, EvenTabsController
from controller.two_files import OneFileCreatedController, TwoFilesController
from controller.path import PathController
from utils.utils_factorise import ColumnDelete, ColumnInsert, LineDelete, LineInsert, TabUpdateFormula, MapIndexLetter, String, Dictionary
from openpyxl.utils import column_index_from_string

 
import os
 

# class TestPath(TestCase):
    
#     def test_create_one_onglet_by_participant(self):
#         path = Path('fichiers_xls/gathertests/')
#         controller = PathController(path, 'test_cmd_ongletbyparticipant.xlsx', 
#                                     OneFileCreatedController(file_options=FileOptions(name_of_tab_to_read='test', column_to_read='A'), new_path=path.pathname))
#         controller.apply_method_on_homononymous_files('split_one_tab_in_multiple_tabs') 
#         file = File('divided_test_cmd_ongletbyparticipant_before.xlsx', path.pathname)
#         file2 = File('divided_test_cmd_ongletbyparticipant.xlsx', path.pathname)
        
#         verify_files_identical(file, file2)
#         os.remove(path.pathname + 'divided_test_cmd_ongletbyparticipant.xlsx') 

#     def test_transform_string_in_binary_in_column(self):
#         path = Path('fichiers_xls/gathertests/')

#         controller = PathController(path, 'test_string_in_binary.xlsx', 
#                                     InsertController(tab_name='sheet1' , tab_options = TabOptions(column_to_read='G', column_to_write='H'), save=True))
#         controller.apply_method_on_homononymous_tabs('transform_string_in_binary_in_column', 'partie 1 : Vrai') 

#         file_object = File('test_string_in_binary.xlsx', path.pathname + 'Nouveau dossier/') 
#         file_object2 = File('test_string_in_binary.xlsx', path.pathname + 'Nouveau dossier - Copie/') 
#         sheeta = file_object.writebook['sheet1']
#         sheetb = file_object.writebook['expected']
#         sheet2a = file_object2.writebook['sheet1']
#         sheet2b = file_object2.writebook['expected']
#         verify_sheets_identical(sheeta, sheetb) 
#         verify_sheets_identical(sheet2a, sheet2b) 
#         sheet2a.delete_cols(8)
#         sheeta.delete_cols(8)
#         file_object.writebook.save(path.pathname + 'Nouveau dossier/test_string_in_binary.xlsx')
#         file_object2.writebook.save(path.pathname + 'Nouveau dossier - Copie/test_string_in_binary.xlsx')

#     def test_copy_a_tab_in_new_workbook(self):
#         path = Path('gatherfiles/')
#         controller = PathController(path, 'test.xlsx', 
#                                     OneFileCreatedController(new_path=path.pathname))
#         controller.apply_method_on_homononymous_files('copy_a_tab_in_new_workbook', 'Sheet') 
#         file = File('gathered_test.xlsx', path.pathname)
#         file2 = File('gathered_test_ref.xlsx', path.pathname)
        
#         verify_files_identical(file, file2)
#         os.remove(path.pathname + 'gathered_test.xlsx') 

    

        

class TestOneTabCreatedController(TestCase):

    def setUp(self):
        self.file_object = None
        self.controller = None
        self.file_data_compare_1 = None
        self.file_data_compare_2 = None
        self.file_options = None

    def test_extract_a_column_from_all_tabs(self):
        self._build_files_to_compare('test_extract_column.xlsx', 'test_extract_column_ref.xlsx')
        self.file_options = FileOptions(column_to_read='B')
        
        self.controller = OneTabCreatedController(self.file_data_compare_1.file_object, self.file_options)
        self.controller.extract_a_column_from_all_tabs()
 
        self._compare_files() 
        tab_to_delete = [self.file_data_compare_1.file_object.sheets_name[-1]]
        self._delete_created_tabs(tab_to_delete)
    
    def _build_files_to_compare(self, file_name, file_name_reference): 
        self.file_data_compare_1 = FileData(file_name)
        self.file_data_compare_1.create_file_object()
        self.file_data_compare_2 = FileData(file_name_reference)
        self.file_data_compare_2.create_file_object() 

    def _compare_files(self):
        self.assert_object = AssertIdentical(self.file_data_compare_1, self.file_data_compare_2) 
        self.assert_object.verify_files_identical()

    def _compare_tabs(self):
        self.assert_object = AssertIdentical(self.file_data_compare_1, self.file_data_compare_2)  
        self.assert_object.verify_tabs_identical()

    def _delete_created_tabs(self, tabs_to_delete):
        self.file_object = self.file_data_compare_1.file_object
        for tab in tabs_to_delete: 
            del self.file_object.writebook[tab]
        self.file_object.writebook.save(self.file_object.path + self.file_object.name_file)

    def test_gather_groups_of_multiple_columns_in_tabs_of_two_columns_containing_tags_and_values(self):
        self._build_files_to_compare('test_gather_columns_in_one.xlsx', 'test_gather_columns_in_one - ref.xlsx')
        self.file_options = FileOptions(name_of_tab_to_read='test')

        self.controller = OneTabCreatedController(self.file_data_compare_1.file_object, self.file_options)
        self.controller.gather_groups_of_multiple_columns_in_tabs_of_two_columns_containing_tags_and_values(['C','D','E'], ['G','H','I'])
 
        self._compare_files()
        tabs_to_delete = ['tab_column_gathered_CDE', 'tab_column_gathered_GHI']
        self._delete_created_tabs(tabs_to_delete)

    def test_gather_multiple_answers(self):
        self._build_gather_multiple_answers()

        self.controller = OneTabCreatedController(self.file_data_compare_1.file_object, self.file_options)
        self.controller.gather_multiple_answers('A', 'B')

        self._compare_tabs()
        tab_to_delete = ['multiple_answers']
        self._delete_created_tabs(tab_to_delete)

    def _build_gather_multiple_answers(self): 
        self.file_data_compare_1 = FileData('testongletbyparticipant.xlsx', 'multiple_answers')
        self.file_data_compare_1.create_file_object()
        self.file_data_compare_2 = FileData('testongletbyparticipant-result.xlsx', 'Feuille2')
        self.file_data_compare_2.create_file_object() 
        self.file_options = FileOptions(name_of_tab_to_read='test')


class TestEvenTabsController(TestCase):
    pass

#     def test_apply_column_formula_on_all_tabs(self):
#         file = File('dataset.xlsx', dataonly = False)
#         controler = EvenTabsController(file, FileOptions(columns_to_read=['B','C']))
#         controler.apply_columns_formula_on_all_tabs()

#     def test_merge_cells_on_all_tabs(self): 
#         file = File("test_merging.xlsx")
#         controler = EvenTabsController(file)
#         controler.merge_cells_on_all_tabs(MergedCellsRange('C', 'D', 12, 15))

#         #voir comment tester le fait qu'une cellule est mergée : comprendre l'objet mergedcells
#         """ for tab in file1.sheets_name:
#             sheet = file1.writebook[tab]
#             mergedcells = sheet.merged_cells
#             print(mergedcells.ranges, type(mergedcells))
#             self.assertEqual('C5' in mergedcells.ranges,True)
#             self.assertIn(sheet['C6'],mergedcells)
#             self.assertIn(sheet['C7'],mergedcells)
#             self.assertIn(sheet['D5'],mergedcells)
#             self.assertIn(sheet['D6'],mergedcells) """
        
#     def test_apply_cell_formula_on_all_sheets(self):
#         file = File("test_merging.xlsx")
#         controler = EvenTabsController(file)
#         controler.apply_cells_formula_on_all_tabs('A10','B10','C10')

#         for tab in file.sheets_name[1:]:
#             sheet = file.writebook[tab]
#             self.assertEqual(sheet['A10'].value, file.writebook[file.sheets_name[0]]['A10'].value)
#             self.assertEqual(sheet['B10'].value, file.writebook[file.sheets_name[0]]['B10'].value)
#             self.assertEqual(sheet['C10'].value, file.writebook[file.sheets_name[0]]['C10'].value)
    
#     def test_check_linenumber_of_tabs(self):
#         file = File('test.xlsx')
#         controler = EvenTabsController(file)
#         tabs = controler.list_tabs_with_different_number_of_lines(14)
#         self.assertListEqual(tabs, ['cutinparts', 'cutinpartsbis', 'delete_lines', 'delete_lines_bis', 'time_min', 'time_min_expected'])


class TestOneFileCreatedController(TestCase):

    def setUp(self): 
        self.file_object = None
        self.controller = None
        self.file_data_compare_1 = None
        self.file_data_compare_2 = None
        self.file_options = None

    def _compare_tabs(self):
        self.file_data_compare_2.create_file_object()
        self.assert_object = AssertIdentical(self.file_data_compare_1, self.file_data_compare_2) 
        self.assert_object.verify_tabs_identical()

    def _compare_files(self):
        self.file_data_compare_2.create_file_object()
        self.assert_object = AssertIdentical(self.file_data_compare_1, self.file_data_compare_2) 
        self.assert_object.verify_files_identical()

    def _delete_created_file(self):
        os.remove(self.file_data_compare_2.file_object.path + self.file_data_compare_2.name_file)

    def test_extract_cells_from_all_tabs(self): 
        self._build_extract_cells_from_all_tabs_data()

        self.controller = OneFileCreatedController(self.file_object)
        self.controller.extract_cells_from_all_tabs('C7','D7','C8','D8') 

        self._compare_tabs()
        self._delete_created_file() 
    
    def _build_extract_cells_from_all_tabs_data(self):
        self.file_object = File('test_extract_cells_from_all_sheets.xlsx')  
        self.file_data_compare_1 = FileData('test_extract_cells_from_all_sheets - after.xlsx', 'gathered_data') 
        self.file_data_compare_2 = FileData('gathered_data_test_extract_cells_from_all_sheets.xlsx', 'Sheet') 

    def test_split_one_tab_in_multiple_tabs(self): 
        self._build_split_one_tab_in_multiple_tabs_data() 

        self.controller = OneFileCreatedController(self.file_object, file_options=self.file_options)
        self.controller.split_one_tab_in_multiple_tabs() 

        self._compare_files()
        self._delete_created_file()
    
    def _build_split_one_tab_in_multiple_tabs_data(self):
        self.file_object = File('test_create_one_onglet_by_participant.xlsx')  
        self.file_data_compare_1 = FileData('test_create_one_onglet_by_participant_before.xlsx')
        self.file_options = FileOptions(name_of_tab_to_read='Stroops_test (7)', column_to_read='A')
        self.file_data_compare_2 = FileData('divided_test_create_one_onglet_by_participant.xlsx')

    def test_one_file_by_tab(self):
        self._build_one_file_by_tab_for_first_tab()

        self.controller = OneFileCreatedController(self.file_object)
        self.controller.create_one_file_by_tab()

        self._compare_tabs()
        self._delete_created_file()

        self._build_one_file_by_tab_for_second_tab()

        self._compare_tabs()
        self._delete_created_file()
    
    def _build_one_file_by_tab_for_first_tab(self):
        self.file_object = File('test_onefile_sendmail.xlsx')  
        self.file_data_compare_1 = FileData('test_onefile_sendmail.xlsx', 'tony fevrier') 
        self.file_data_compare_2 = FileData('tony fevrier.xlsx', 'Sheet', path="multifiles/")

    def _build_one_file_by_tab_for_second_tab(self):
        self.file_data_compare_1 = FileData('test_onefile_sendmail.xlsx', 'Marine Moyon') 
        self.file_data_compare_2 = FileData('Marine Moyon.xlsx', 'Sheet', path="multifiles/")


class TestColumnInsertion(TestCase):

    def setUp(self): 
        self.controller = None
        self.file_data = None
        self.file_data_compare = None
        self.method_data = None 
        self.tab_options = None
        self.columns_to_compare = []
        self.columns_to_delete = []

    def apply_compare_columns_restore(fonction):
        def wrapper(self):
            fonction(self) 
            self._apply_the_method_to_test() 
            self._compare_new_columns()
            self._restore_file_state_before_modification()
        return wrapper 
    
    def apply_compare_tabs_restore(fonction):
        def wrapper(self):
            fonction(self) 
            self._apply_the_method_to_test() 
            self._compare_tabs()
            self._restore_file_state_before_modification()
        return wrapper 
    
    def _apply_the_method_to_test(self):
        tab_controller = InsertController(self.file_data.file_object, tab_options=self.tab_options)  
        file_options = FileOptions(names_of_tabs_to_modify=[self.file_data.tab_name])
        self.controller = MultipleSameTabController(self.file_data.file_object, tab_controller, file_options)
        self.controller.apply_method_on_some_tabs(self.method_data.method_name, *self.method_data.args, **self.method_data.kwargs) 

    def _compare_new_columns(self):
        self.assert_object = AssertIdentical(self.file_data, self.file_data_compare) 
        columns_to_compare = MapIndexLetter.get_list_of_columns_indexes(self.columns_to_compare)
        for column in columns_to_compare:
            self.assert_object.verify_columns_identical(column, column)

    def _compare_tabs(self):
        self.assert_object = AssertIdentical(self.file_data, self.file_data_compare) 
        self.assert_object.verify_tabs_identical()
    
    def _restore_file_state_before_modification(self):
        tab = self.file_data.file_object.writebook[self.file_data.tab_name]  
        tab.delete_cols(column_index_from_string(self.columns_to_delete[0]), len(self.columns_to_delete))
        modification_object = ColumnDelete(self.columns_to_delete)
        self.controller.tab_controller._update_cell_formulas(modification_object) 
        self.file_data.file_object.writebook.save(self.file_data.file_object.path + self.file_data.name_file) 

    @apply_compare_columns_restore
    def test_transform_string_in_binary_in_column(self): 
        self.file_data = FileData('test.xlsx', 'Feuille2')
        self.file_data_compare = FileData('test.xlsx', 'Feuille2')
        self.tab_options = TabOptions(column_to_read='F', column_to_write='G')
        self.method_data = MethodData('transform_string_in_binary_in_column', 'partie 12 : Faux', 1) 
        self.columns_to_compare = ['G']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_columns_restore
    def test_convert_time_in_minutes_in_columns(self): 
        self.file_data = FileData('test.xlsx', 'time_min')
        self.file_data_compare = FileData('test.xlsx', 'time_min_expected')
        self.tab_options = TabOptions(column_to_read='E', column_to_write='F')
        self.method_data = MethodData('convert_time_in_minutes_in_columns') 
        self.columns_to_compare = ['F']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_columns_restore
    def test_insert_group_associated_with_answer(self):
        self.file_data = FileData('test_column_set_answer.xlsx', 'sheet1')
        self.file_data_compare = FileData('test_column_set_answer.xlsx', 'Feuille2')
        self.tab_options = TabOptions(column_to_read='B', column_to_write='C')
        map_groups_to_answers = {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']}  
        self.method_data = MethodData('insert_group_associated_with_answer', map_groups_to_answers)
        self.columns_to_compare = ['C']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_columns_restore
    def test_insert_splitted_strings_of(self):
        self.file_data = FileData('test.xlsx', 'cutinparts')
        self.file_data_compare = FileData('test.xlsx', 'cutinpartsbis') 
        self.tab_options = TabOptions(column_to_read='B', column_to_write='C')
        self.method_data = MethodData('insert_splitted_strings_of', ';')
        self.columns_to_compare = ['C', 'D', 'E']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_columns_restore
    def test_fill_one_column_by_QCM_answer(self):
        self.file_data = FileData('test_create_one_column.xlsx', 'sheet1')
        self.file_data_compare = FileData('test_create_one_column.xlsx', 'Feuille2') 
        self.tab_options = TabOptions(column_to_read='D', column_to_write='E')
        self.method_data = MethodData('fill_one_column_by_QCM_answer', 'Alain', 'Henri', 'Tony', 'Dulcinée')
        self.columns_to_compare = ['E', 'F', 'G', 'H']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_columns_restore
    def test_write_piece_of_string_in_column_1(self):
        self.file_data = FileData('test_colgetpartofstr.xlsx', 'Feuille2')
        self.file_data_compare = FileData('test_colgetpartofstr.xlsx', 'expected') 
        self.tab_options = TabOptions(column_to_read='C', column_to_write='D')
        self.method_data = MethodData('write_piece_of_string_in_column', '_', 0)
        self.columns_to_compare = ['D']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_columns_restore
    def test_write_piece_of_string_in_column_2(self):
        self.file_data = FileData('test_colgetpartofstr.xlsx', 'Feuille2')
        self.file_data_compare = FileData('test_colgetpartofstr.xlsx', 'expected2') 
        self.tab_options = TabOptions(column_to_read='E', column_to_write='F')
        self.method_data = MethodData('write_piece_of_string_in_column', ';', 1)
        self.columns_to_compare = ['F']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_columns_restore
    def test_map_two_columns_to_a_third_column(self):
        self.file_data = FileData('test_maptwocolumns.xlsx', 'Feuille2')
        self.file_data_compare = FileData('test_maptwocolumns.xlsx', 'expected') 
        self.tab_options = TabOptions(columns_to_read=['B', 'C'], column_to_write='D')
        self.method_data = MethodData('map_two_columns_to_a_third_column', {'cat1':['prime','1'], 'cat2':['probe','2']})
        self.columns_to_compare = ['D']
        self.columns_to_delete = self.columns_to_compare

    @apply_compare_tabs_restore
    def test_verify_tabs_when_map_two_columns_to_a_third_column(self):
        self.file_data = FileData('test_maptwocolumns.xlsx', 'Feuille2')
        self.file_data_compare = FileData('test_maptwocolumns.xlsx', 'expected') 
        self.tab_options = TabOptions(columns_to_read=['B', 'C'], column_to_write='D')
        self.method_data = MethodData('map_two_columns_to_a_third_column', {'cat1':['prime','1'], 'cat2':['probe','2']})
        self.columns_to_delete = ['D']

    @apply_compare_tabs_restore
    def test_insert_column_for_prime_probe_congruence(self):   
        self.file_data = FileData('test_prime_probe.xlsx', 'Feuille2')
        self.file_data_compare = FileData('test_prime_probe.xlsx', 'expected') 
        self.tab_options = TabOptions(columns_to_read=['B', 'C', 'D'], column_to_write='E')
        self.method_data = MethodData('insert_column_for_prime_probe_congruence')
        self.columns_to_delete = ['E']

    @apply_compare_tabs_restore
    def test_insert_tags_of_maximum_of_column_list(self):
        self.file_data = FileData('test_give_names.xlsx', 'sheet1')
        self.file_data_compare = FileData('test_give_names.xlsx', 'Feuille2') 
        self.tab_options = TabOptions(columns_to_read=['A', 'B', 'C'], column_to_write='D')
        self.method_data = MethodData('insert_tags_of_maximum_of_column_list')
        self.columns_to_delete = ['D']


class TestDeleteItems(TestCase):
    def setUp(self):
        pass

    #     def test_delete_lines(self):
#         file = File('test.xlsx')
#         controler = MultipleSameTabController(file, tab_controller=DeleteController(file),
#                                               file_options=FileOptions(names_of_tabs_to_modify=['delete_lines']))  
#         controler.apply_method_on_some_tabs('delete_lines_containing_strings_in_given_column', 'D', '0')
#         controler.apply_method_on_some_tabs('delete_lines_containing_strings_in_given_column', 'D', 'p a') 
#         column_identical('test.xlsx','test.xlsx',9,10, 1, 1)
#         column_identical('test.xlsx','test.xlsx',9,10, 2, 2)
#         column_identical('test.xlsx','test.xlsx',9,10, 3, 3)
#         column_identical('test.xlsx','test.xlsx',9,10, 4, 4)
#         column_identical('test.xlsx','test.xlsx',9,10, 5, 5)
#         column_identical('test.xlsx','test.xlsx',9,10, 6, 6)

#     def test_delete_lines_with_formulas(self):
#         file = File('listing_par_etape - Copie.xlsx')
#         controler = MultipleSameTabController(file, tab_controller=DeleteController(file),
#                                               file_options=FileOptions(names_of_tabs_to_modify=['Feuil1']))   
#         controler.apply_method_on_some_tabs('delete_lines_containing_strings_in_given_column', 'B', 'pas consenti') 
#         column_identical('listing_par_etape - Copie.xlsx','listing_par_etape - Copie.xlsx',0, 1, 2, 2)
#         column_identical('listing_par_etape - Copie.xlsx','listing_par_etape - Copie.xlsx',0, 1, 10, 10) 

#     def test_delete_doublons(self): 
#         file = File('test_doublons.xlsx')
#         controler = MultipleSameTabController(file, 
#                                               tab_controller=DeleteController(file), 
#                                               file_options=FileOptions(names_of_tabs_to_modify=['sheet2', 'sheet1', 'sheet3']))
#         sheet_result = file.writebook['result']  
#         controler.apply_method_on_some_tabs('delete_twins_lines_and_color_last_twin', 'C', color = 'FFFFFF00') 
#         sheet1 = file.writebook['sheet1']  
#         sheet2 = file.writebook['sheet2']  
#         sheet3 = file.writebook['sheet3']  
#         verify_sheets_identical(sheet1, sheet_result)
#         verify_sheets_identical(sheet2, sheet_result)
#         verify_sheets_identical(sheet3, sheet_result)
        
#     """ def test_delete_other_columns(self):
#         # Fonctionnel une fois
#         file = File('test_keep_only_columns.xlsx')
#         controller = DeleteController(file, 'sheet1')
#         controller.delete_other_columns('A-C,D-K')

#         verify_sheets_identical(file.get_tab_by_name('sheet1'), File('test_keep_only_columns.xlsx').get_tab_by_name('Feuille2')) """

#     def test_delete_columns(self):
#         # Fonctionnel une fois
#         file = File('test_keep_only_columns.xlsx')
#         controller = MultipleSameTabController(file, DeleteController(file, 'sheet1'), 
#                                   file_options=FileOptions(names_of_tabs_to_modify=['sheet1']))  
                                   
#         controller.apply_method_on_some_tabs('delete_columns', 'L,M,N-V')

#         verify_sheets_identical(file.get_tab_by_name('sheet1'), File('test_keep_only_columns.xlsx').get_tab_by_name('Feuille2'))



class TestColorItems(TestCase):
    def setUp(self):
        pass
    
#     def test_color_column(self):
#         file = File('test.xlsx')
#         controler = MultipleSameTabController(file,
#                                               ColorTabController(file, tab_options=TabOptions(column_to_read='D')), 
#                                               file_options=FileOptions(names_of_tabs_to_modify=['cutinpartsbis']))  
#         #controler.tab_controller.tab_options.column_to_read = 'D'#TabOptions(column_to_read='D')
#         controler.apply_method_on_some_tabs('color_cases_in_column', 
#                                             {' partie 2 : Vrai':'0000a933'}) 
         
#     def test_color_cases_in_sheet(self):
#         file = File('test.xlsx')
#         controler = MultipleSameTabController(file,
#                                               ColorTabController(file, tab_options=TabOptions()), 
#                                               file_options=FileOptions(names_of_tabs_to_modify=['cutinpartsbis']))   
#         controler.apply_method_on_some_tabs('color_cases_in_sheet', 
#                                             {'partie 1 : Vrai':'0000a933', 'Abbas':'0000a933'}) 
        
#     def test_color_line_containing_chaines(self):
#         file = File('test.xlsx')
#         controler = MultipleSameTabController(file, 
#                                               ColorTabController(file, color='0000a933'),
#                                               file_options=FileOptions(names_of_tabs_to_modify=['color_line']))    
#         controler.apply_method_on_some_tabs('color_lines_containing_strings', '-', '+') 


class TestTwoFilesController(TestCase):
    def setUp(self): 
        self.file_object = None
        self.controller = None
        self.file_data_compare_1 = None
        self.file_data_compare_2 = None
        self.column_to_read_1 = None
        self.column_to_read_2 = None
        self.columns_to_compare = []
        self.columns_to_copy = []
        self.column_insertion = None

    def test_copy_columns_in_a_tab_differently_sorted(self):
        self._build_test()
        self._apply_method_to_test()
        self._compare_new_columns()
        self._restore_file_state_and_save() 

    def _build_test(self):
        self.file_object_from = File('test.xlsx', dataonly=True)
        self.tab_name_from = 'sheet1'
        self.file_data_compare_1 = FileData('test.xlsx', 'Feuille5')
        self.file_data_compare_2 = FileData('test.xlsx', 'Feuille5bis')
        self.file_data_compare_1.create_file_object() 
        self.column_to_read_from = 'C'
        self.column_to_read_to = 'C'
        self.columns_to_compare = ['E', 'F', 'H']
        self.columns_to_copy = ['B','F']
        self.column_insertion = 'E'

    def _apply_method_to_test(self):
        self.file1 = self.file_data_compare_1.file_object
        self.file2 = self.file_data_compare_2.file_object
        self.controller = TwoFilesController(self.file_object_from, self.file1,
                                             self.tab_name_from, self.file_data_compare_1.tab_name,
                                             self.column_to_read_from, self.column_to_read_to)
        self.controller.copy_columns_in_a_tab_differently_sorted(self.columns_to_copy, self.column_insertion)

    def _compare_new_columns(self):
        self.assert_object = AssertIdentical(self.file_data_compare_1, self.file_data_compare_2) 
        columns_to_compare_indexes = MapIndexLetter.get_list_of_columns_indexes(self.columns_to_compare)
        for column_index in columns_to_compare_indexes:
            self.assert_object.verify_columns_identical(column_index, column_index)

    def _restore_file_state_and_save(self):
        tab_modified = self.file1.writebook[self.file_data_compare_1.tab_name]        
        tab_modified.delete_cols(column_index_from_string(self.column_insertion), len(self.columns_to_copy))
        TabUpdateFormula(ColumnDelete(['E','F'])).update_cells_formulas(tab_modified)
        self.file1.writebook.save(self.file1.path + self.file1.name_file)


class TestString(TestCase, String):
    
    def test_transform_string_in_binary(self):
        self.assertEqual(self.transform_string_in_binary('rrr','rr', 'rrr'), 1)
        self.assertEqual(self.transform_string_in_binary('rr','rrr'), 0)
        self.assertEqual(self.transform_string_in_binary('','rrr'), 0)

    def test_set_answer_in_group(self): 
        map_groups_to_answers = {"group1":['2','5','6'], "group2":['7','8','9'], "group3":['1','3','4'], "group4":['10']}
        map_answers_to_groups = Dictionary.reverse_dictionary(map_groups_to_answers) 
        
        self.assertEqual(self.set_answer_in_group('1', map_answers_to_groups), "group3")
        self.assertEqual(self.set_answer_in_group('9', map_answers_to_groups), "group2")

    def test_clean_string_from_spaces(self):
        strings = ['tony', 'tony ', ' tony', ' tony ', 'tony  ', '  tony']
        for string in strings:
            cleaned_string = self.clean_string_from_spaces(string)
            self.assertEqual(cleaned_string, 'tony')
            
    def test_convert_time_in_minutes(self):
        map_durations_to_minutes = {"2 jours 2 heures": '3000,0', "1 heure 25 min": '85,0', "16 min 35 s": '16,58'}
        for duration in map_durations_to_minutes.keys():
            duration_in_min = self.convert_time_in_minutes(duration)
            duration_in_min_expected = map_durations_to_minutes[duration]
            self.assertEqual(duration_in_min, duration_in_min_expected)

    def test_get_columns_from(self):
        string = "C-E,H,J-L"
        expected_list = ['C','D','E','H','J','K','L']
        self.assertListEqual(self.get_columns_from(string), expected_list)

    def test_get_range_letter(self):
        string = 'D-H'
        expected_list = ['D','E','F','G','H']
        self.assertListEqual(self.get_range_letter(string), expected_list)


class TestTabUpdate(TestCase, TabUpdateFormula):

    def test_update_formula_when_insert_one_line(self):
        modification_object = LineInsert(['2'])
        updated_formula = modification_object._update_a_cell("SI(J10+K$1+L$3)")
        self.assertEqual(updated_formula, "SI(J11+K$1+L$4)")

    def test_update_formula_when_delete_one_line(self):
        modification_object = LineDelete(['11'])
        updated_formula = modification_object._update_a_cell("SI(J12+K$1+L$3)")
        self.assertEqual(updated_formula, "SI(J11+K$1+L$3)")

    def test_update_formula_when_delete_one_line_greater_than_all_line_numbers(self):
        modification_object = LineDelete(['13'])
        updated_formula = modification_object._update_a_cell("SI(J12+K$1+L$3)")
        self.assertEqual(updated_formula, "SI(J12+K$1+L$3)")

    def test_update_formula_when_insert_one_column(self):
        modification_object = ColumnInsert(['C'])
        updated_formula = modification_object._update_a_cell("SI(J10+K$1+L$3)")
        self.assertEqual(updated_formula, "SI(K10+L$1+M$3)")

    def test_update_formula_when_delete_one_column(self):
        modification_object = ColumnDelete(['C'])
        updated_formula = modification_object._update_a_cell("SI(J10+K$1+L$3)")
        self.assertEqual(updated_formula, "SI(I10+J$1+K$3)")

    def test_update_formula_when_insert_multiple_lines(self):
        modification_object = LineInsert(['2', '5'])
        updated_formula = modification_object._update_a_cell("SI(J10+K$1+L$3)") 
        self.assertEqual(updated_formula, "SI(J12+K$1+L$4)")
 
    def test_update_formula_when_insert_multiple_columns(self):
        modification_object = ColumnInsert(['C', 'D','E', 'F'])
        updated_formula = modification_object._update_a_cell("SI(D10+E$1)") 
        self.assertEqual(updated_formula, "SI(H10+I$1)")


class AssertIdentical(TestCase):
    """Check if two files are the same, or have the same tabs or columns"""

    def __init__(self, file_data1, file_data2, *args, **kwargs):
        super().__init__(*args, **kwargs) 
        self.file_data1 = file_data1
        self.file_data2 = file_data2
        self.file_object1 = self.file_data1.file_object
        self.file_object2 = self.file_data2.file_object
        self._initialize_tabs_attributes()

    def _initialize_tabs_attributes(self): 
        try:  
            self.tab1 = self.file_object1.writebook[self.file_data1.tab_name] 
            self.tab2 = self.file_object2.writebook[self.file_data2.tab_name]  
        except KeyError: 
            self.tab1 = None
            self.tab2 = None

    def verify_files_identical(self):  
        self.assertEqual(self.file_object1.sheets_name, self.file_object2.sheets_name)

        for tab_name in self.file_object1.sheets_name: 
            self.tab1 = self.file_object1.writebook[tab_name]
            self.tab2 = self.file_object2.writebook[tab_name]
            self.verify_tabs_identical()

    def verify_tabs_identical(self):   
        self.assertEqual(self.tab1.max_row, self.tab2.max_row)
        self.assertEqual(self.tab1.max_column, self.tab2.max_column)

        for i in range(1,self.tab1.max_row+1):
            for j in range(1,self.tab1.max_column+1):
                self.assertEqual(self.tab1.cell(i,j).value,self.tab2.cell(i,j).value)

    def verify_columns_identical(self, column1, column2):  
        self.assertEqual(self.tab1.max_row, self.tab2.max_row) 
            
        for i in range(2, self.tab1.max_row + 1): 
            self.assertEqual(self.tab1.cell(i, column1).value,self.tab2.cell(i, column2).value)


class FileData():
    """Data objects containing file informations for test"""
    def __init__(self, name_file, tab_name=None, path='fichiers_xls/'):
        self.name_file = name_file
        self.tab_name = tab_name
        self.path = path
        self.initialize_file_object()

    def initialize_file_object(self):
        try:
            self.create_file_object()
        except FileNotFoundError:
            self.file_object = None            

    def create_file_object(self):
        self.file_object = File(self.name_file, path=self.path)


class MethodData():
    """Data object containing method arguments for test"""
    def __init__(self, method_name, *args, **kwargs):
        self.method_name = method_name 
        self.args = args
        self.kwargs = kwargs
    

if __name__== "__main__":
    main()
import module_pour_excel as mpe

#file = mpe.File('dataset_après_moulinette_ter(2).xlsx',path = 'dataset_dulcinee/')
#file.sauvegarde()

sheet = mpe.Sheet('dataset_après_moulinette_ter(2).xlsx','Feuille22',path = 'dataset_dulcinee/')
#sheet.add_column_in_sheet_differently_sorted(3, 4,['nom_mails_cohorte_consentants_et_non_consentants.xlsx','nom_num_cohorte_all',3,[5]])
sheet.add_column_in_sheet_differently_sorted(3, 7,['questionnaire_de_fin.xlsx','2023_SPOCS&amp;S Questionnaire ',3,[i for i in range(4,16)]])

#sheet2 = mpe.Sheet('pour_analyse_cohorte_bis.xlsx','que terminé et tps<=20min',path = 'dataset_dulcinee/')
#sheet2.add_column_in_sheet_differently_sorted(3, 4,['nom_mails_cohorte_consentants_et_non_consentants.xlsx','nom_num_cohorte_all',3,[5]])
#sheet2.add_column_in_sheet_differently_sorted(3, 8,['dataset_après_moulinette_4(2).xlsx','consent vs nn consent  cohort',3,[74]])

#sheet.column_cut_str_in_parts(13,14,';')



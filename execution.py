import module_pour_excel as mpe

file = mpe.File('dataset_après_moulinette_ter.xlsx',path = 'dataset_dulcinee/')
file.sauvegarde()
sheet = mpe.Sheet('dataset_après_moulinette_ter.xlsx','2023_SPOCS&amp;S Questionna',path = 'dataset_dulcinee/')
sheet.column_cut_str_in_parts(13,14,';')



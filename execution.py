import module_pour_excel as mpe
import json
#file = mpe.File('temps_SPOC.xlsx',path = 'dataset_dulcinee/')
#file.sauvegarde()

sheet = mpe.Sheet('dataset_final_25092023.xlsx','COPIE ',path = 'dataset_dulcinee/')
#sheet.add_column_in_sheet_differently_sorted(3, 258,['questionnaire_de_fin.xlsx','2023_SPOCS&amp;S Questionnaire ',3,[i for i in range(8,16)]])

def createDicoGroup(dicoOrigin):
    dico = {}
    for key in dicoOrigin.keys():
        for value in dicoOrigin[key]:
            dico[value]=key
    return dico

sheet.column_set_answer_in_group(258,259,createDicoGroup({"Absolutisme":['1','4','5','6'],"Evaluatisme":['3'],"Multiplisme":['2']}),line_end=1337)
"""
sheet.column_set_answer_in_group(257,258,createDicoGroup({"Absolutisme":['4','6','7'],"Evaluatisme":['2','3','5'],"Multiplisme":[]}),line_end=866)
sheet.column_set_answer_in_group(259,260,createDicoGroup({"Absolutisme":[],"Evaluatisme":[],"Multiplisme":[]}),line_end=866)
sheet.column_set_answer_in_group(261,262,createDicoGroup({"Absolutisme":['1','4','5'],"Evaluatisme":['2','3'],"Multiplisme":['6']}),line_end=866)
sheet.column_set_answer_in_group(263,264,createDicoGroup({"Absolutisme":['1','2','3','5','8','9'],"Evaluatisme":['4','6','7'],"Multiplisme":['10']}),line_end=866)
sheet.column_set_answer_in_group(265,266,createDicoGroup({"Absolutisme":['1','2','3','6'],"Evaluatisme":['4','5','7'],"Multiplisme":['8']}),line_end=866)
sheet.column_set_answer_in_group(267,268,createDicoGroup({"Absolutisme":['4','5'],"Evaluatisme":['1','2'],"Multiplisme":['3']}),line_end=866)
sheet.column_set_answer_in_group(269,270,createDicoGroup({"Absolutisme":['1','3'],"Evaluatisme":['4','5'],"Multiplisme":['2']}),line_end=866) 
"""



#print(createDicoGroup({"Absolutisme":['1','4','5','6'],"Evaluatisme":['3'],"Multiplisme":['2']}))

"""
sheet = mpe.Sheet('dataset_final_01092023.xlsx','test_fin_style épis',path = 'dataset_dulcinee/')

sheet.column_set_answer_in_group(255,256,createDicoGroup({"Absolutisme":['1','4','5','6'],"Evaluatisme":['3'],"Multiplisme":['2']}),line_end=866)
sheet.column_set_answer_in_group(257,258,createDicoGroup({"Absolutisme":['4','6','7'],"Evaluatisme":['2','3','5'],"Multiplisme":[]}),line_end=866)
sheet.column_set_answer_in_group(259,260,createDicoGroup({"Absolutisme":[],"Evaluatisme":[],"Multiplisme":[]}),line_end=866)
sheet.column_set_answer_in_group(261,262,createDicoGroup({"Absolutisme":['1','4','5'],"Evaluatisme":['2','3'],"Multiplisme":['6']}),line_end=866)
sheet.column_set_answer_in_group(263,264,createDicoGroup({"Absolutisme":['1','2','3','5','8','9'],"Evaluatisme":['4','6','7'],"Multiplisme":['10']}),line_end=866)
sheet.column_set_answer_in_group(265,266,createDicoGroup({"Absolutisme":['1','2','3','6'],"Evaluatisme":['4','5','7'],"Multiplisme":['8']}),line_end=866)
sheet.column_set_answer_in_group(267,268,createDicoGroup({"Absolutisme":['4','5'],"Evaluatisme":['1','2'],"Multiplisme":['3']}),line_end=866)
sheet.column_set_answer_in_group(269,270,createDicoGroup({"Absolutisme":['1','3'],"Evaluatisme":['4','5'],"Multiplisme":['2']}),line_end=866)
"""
"""
sheet = mpe.Sheet('currency.xlsx','Active',path = 'fichiers_xls/')

dico = sheet.create_dico_from_columns(3,2,5,276)
print(dico)

dico_js = json.dumps(dico)


nom_fichier = 'currency.json'

with open(nom_fichier,'w') as f:
    f.write(dico_js)

f.close()
"""

#sheet.add_column_in_sheet_differently_sorted(3, 4,['nom_mails_cohorte_consentants_et_non_consentants.xlsx','nom_num_cohorte_all',3,[5]])
#sheet.add_column_in_sheet_differently_sorted(3, 7,['questionnaire_de_fin.xlsx','2023_SPOCS&amp;S Questionnaire ',3,[i for i in range(4,16)]])

#sheet2 = mpe.Sheet('pour_analyse_cohorte_bis.xlsx','que terminé et tps<=20min',path = 'dataset_dulcinee/')
#sheet2.add_column_in_sheet_differently_sorted(3, 4,['nom_mails_cohorte_consentants_et_non_consentants.xlsx','nom_num_cohorte_all',3,[5]])
#sheet2.add_column_in_sheet_differently_sorted(3, 8,['dataset_après_moulinette_4(2).xlsx','consent vs nn consent  cohort',3,[74]])

#sheet.column_cut_str_in_parts(13,14,';')



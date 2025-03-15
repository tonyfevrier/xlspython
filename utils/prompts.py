directory_prompt = 'Enter the name of the directory containing all directories'
multiple_tabs_prompt = "If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter"
multiple_file_prompt = 'Enter the name of the file to work on. You must write its extension (.xlsx, .xlst). Each directory must contain this file name.'
file_prompt = 'Enter the name of the file to work on. You must write its extension (.xlsx, .xlst)'  
sheet_prompt = 'Enter the name of the sheet to work on'
column_read_prompt = 'Enter the column letter to read'
column_store_prompt = 'Enter the column letter where to store the data'
group_column_prompt = "Enter a group of column of the form A-D,E,G,H-J,Z. There must be no space."
line_prompt = '(Optional) Enter the number of the line or press enter'
cell_prompt = "Enter cells you want to extract with groups of cells separated by a comma. For example : C12,C14-16,D-F13,H-I:21-22 allows to extract the following cells C12 C14 C15 C16 D13 E13 F13 H21 H22 I21 I22"  #"Entrez une cellule que vous souhaitez copier : "
color_prompt = "Please enter the color in hexadecimal type"
separator_prompt = '(Optional) Enter the separator'

def ask_argument_prompt(name):
    return f'"Enter one {name} and then press enter. Press directly enter if you have entered all the good {name}s'
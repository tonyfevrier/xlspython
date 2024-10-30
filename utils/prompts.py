directory_prompt = 'Enter the name of the directory containing all directories'
multiple_tabs_prompt = "If you want to execute the program on all sheets, press immediately enter. Otherwise write sheet names one by one and press enter each time. When you write all sheets, press enter"
file_prompt = 'Enter the name of the file to work on. You must write its extension (.xlsx, .xlst)'  
sheet_prompt = 'Enter the name of the sheet to work on'
column_read_prompt = 'Enter the column letter to read'
column_store_prompt = 'Enter the column letter where to store the data'
group_column_prompt = "Enter a group of column of the form A-D,E,G,H-J,Z. There must be no space."
line_prompt = '(Optional) Enter the number of the line or press enter'
cell_prompt = "Entrez une cellule que vous souhaitez copier : "
color_prompt = "Please enter the color in hexadecimal type"

def ask_argument_prompt(name):
    return f'"Enter one {name} and then press enter. Press directly enter if you have entered all the good {name}s'
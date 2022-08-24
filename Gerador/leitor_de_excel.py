import tkinter as tk
from tkinter import filedialog
import pandas as pd
import PIL

import_file_path = r"X:\SERVIÇOS\TM-S\Relatório Padrão TMS - Não Apagar.xlsx"
df = pd.read_excel(import_file_path)
print(df["TEXTO"][0])



#import pandas as pd

#X:\SERVIÇOS\TM-S\Relatório Padrão TMS - Não Apagar.xlsx

#file = 'X:\SERVIÇOS\TM-S\Relatório Padrão TMS - Não Apagar.xlsx'

#x1 = pd.ExcelFile(file)

#print(x1.sheet_names)

#df1 = x1.parse('Planilha1')

#data = pd.read_excel (r'X:\SERVIÇOS\TM-S\Relatório Padrão TMS - Não Apagar.xlsx')
#df = pd.DataFrame(data, columns = ['TÍTULO'])
#print(df)

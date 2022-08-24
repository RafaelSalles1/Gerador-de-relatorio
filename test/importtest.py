import os
import sys
from _tkinter import *
from docx import Document
#pasta_que_eu_quero = sys.path[1] + r'\Relatório Padrão TMS - Não Apagar.xlsx'
#print(pasta_que_eu_quero)

#print(sys.path[0])

#if r'\test' in sys.path[0]:
 ##  teste = teste.replace(r'\test', '')
   # print(teste)


document = Document()
document.save('test.docx')


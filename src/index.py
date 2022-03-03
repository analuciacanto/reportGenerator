import json
import codecs
import openpyxl as xl
from openpyxl import Workbook
from firstTab import createFirstTab
from secondTab import createSecondTab

import time
ini = time.time()
iniJson = time.time()

# Importando Json
dados = json.load(codecs.open('jsonV3.json', 'r', 'utf-8-sig'))
data = dados['data']
users = data['users']

fimJson = time.time()


iniRelatorio = time.time()
# Gerando relatório
wb = Workbook()
dest_filename = 'Declaracao progresso.xlsx'
fimRelatorio = time.time()

iniFirstTab = time.time()
# First Tab
createFirstTab(wb, users)

fimFirstTab = time.time()
# Others Tabs

iniSecondTab = time.time()

createSecondTab(wb, users)
fimSecondTab = time.time()

# Formatando

# Utils

wb.save(filename=dest_filename)
fim = time.time()
print("Tempo total de convers?o Json: ", fimJson-iniJson)
print("Tempo total de gera??o do relat?rio: ", fimRelatorio-iniRelatorio)
print("Tempo total de gera??o da primeira aba: ", fimFirstTab-iniFirstTab)
print("Tempo total de gera??o da segunda aba: ", fimSecondTab-iniSecondTab)
print("Tempo total ", fim-ini)

import json
import codecs
import openpyxl as xl
from openpyxl import Workbook
from firstTab import createFirstTab
from otherTabs import createOtherTabs


# Importando Json
dados = json.load(codecs.open('jsonV3.json', 'r', 'utf-8-sig'))
data = dados['data']
users = data['users']

# Gerando relatório
wb = Workbook()
dest_filename = 'Declaracao progresso.xlsx'


# First Tab
createFirstTab(wb, users)

# Others Tabs
createOtherTabs(wb, users)
# Formatando


# Utils


wb.save(filename=dest_filename)

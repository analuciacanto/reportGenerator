import json
import codecs
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, alignment

# Importando Json
dados = json.load(codecs.open('jsonV3.json', 'r', 'utf-8-sig'))
data = dados['data']
users = data['users']

# Gerando relatório
wb = Workbook()
dest_filename = 'Declaracao progresso.xlsx'

# First Tab
ws1 = wb.active
ws1.title = "Vigência"

header = ["UID",	"Nome",	"E-mail",	"CPF",	"Empresa",	"Departamento",	"Grupos",
          "Perfil",	"Status",	"Dt Cadastro do usuário", "Formulário Preenchido?"]
ws1.append(header)

for row in range(len(users)):
    user = users[row]
    ws1.append([user['id'], user
               ['first_name'], user['email'], user['number_id'], user["company"]["name"], user["department"]["name"], "--", "--", "Ativo" if user['active'] else "Inativo", user['created'], "--" if user['forms'] == [] else "Preenchido"])


# Formatando

# Header
header_cells = ['A1', 'B1', 'C1', 'D1',
                'E1', 'F1', "G1", "H1", "I1", "J1", "K1"]
for cell in header_cells:

    ws1[cell].fill = PatternFill(start_color='00003366',
                                 end_color='00003366',
                                 fill_type='solid')

    ws1[cell].font = Font(bold=True, color="00FFFFFF",
                          size="12", name='Calibri',)

    ws1[cell].border = Border(left=Side(border_style='thin', color='00C0C0C0'),
                              right=Side(border_style='thin',
                                         color='00C0C0C0'),
                              top=Side(border_style='thin',
                                       color='00C0C0C0'),
                              bottom=Side(border_style='thin',
                                          color='00C0C0C0'),)

# Other Lines

cells = ['A', 'B', 'C', 'D',
         'E', 'F', "G", "H", "I", "J", "K"]

for i in range(2, len(users) + 2):
    for cell in cells:
        ws1[cell + str(i)].border = Border(left=Side(border_style='thin', color='00C0C0C0'),
                                           right=Side(border_style='thin',
                                                      color='00C0C0C0'),
                                           top=Side(border_style='thin',
                                                    color='00C0C0C0'),
                                           bottom=Side(border_style='thin',
                                                       color='00C0C0C0'),)
        ws1[cell + str(i)].alignment = Alignment(horizontal='left')
        ws1[cell + str(i)].font = Font(bold=False,
                                       size="12", name='Calibri')
        if cell == "K":
            if ws1[cell + str(i)].value == "Preenchido":
                ws1[cell + str(i)].fill = PatternFill(start_color='00008000',
                                                      end_color='00008000',
                                                      fill_type='solid')
            else:
                ws1[cell + str(i)].fill = PatternFill(start_color='00FF0000',
                                                      end_color='00FF0000',
                                                      fill_type='solid')

        else:
            ws1[cell + str(i)].fill = PatternFill(start_color='00F5F5F5',
                                                  end_color='00F5F5F5',
                                                  fill_type='solid')


# Utils


def setCellWidth(ws):
    for col in ws.columns:
        max_lenght = 8
        col_name = re.findall('\w\d', str(col[0]))
        col_name = col_name[0]
        col_name = re.findall('\w', str(col_name))[0]
        for cell in col:
            try:
                if len(str(cell.value)) > max_lenght:
                    max_lenght = len(cell.value)
            except:
                pass
        adjusted_width = (max_lenght + 5)
        ws.column_dimensions[col_name].width = adjusted_width


setCellWidth(ws1)
wb.save(filename=dest_filename)

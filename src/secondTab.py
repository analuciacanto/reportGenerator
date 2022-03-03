from math import ceil
from sqlite3 import Row
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from utils import setCellWidth


# First Tab
import re


def createSecondTab(wb, users):
    ws = wb.create_sheet(
        title="Form-id")
    addHeader(ws, users)
    setCellWidth(ws, 1, 1, ws.max_column)


def addHeader(ws, users):
    header = ["Dec ID",	"UID",	"Nome do usuário",	"E-mail",	"CPF",	 "Empresa", "Departamento",	"Grupos a que pertence",
              "Gestor do usuário",	"Perfil do usuário",	'Status do usuário', "Data de cadastro do usuário",
              "Etapa atual da análise",	"Data da avaliação do Gestor",	"Preenchido Por",
              "Data do Preenchimento",	"Conclusão da análise",	"Última atualização feita por",	"Data da última atualização",
                        "Número de anexos na análise",	"Anotações internas",	"Parecer",  "Você está preenchendo para outro colaborador? ",
              "Qual colaborador?", "Algum outro colaborador/Administrador da CCR estava presente?", "Adicionar Colaborador/Administrador", "Nome", "Cargo", "Divisão", "Unidade de Negócios",
              "Justifique:", "Dados dos Colaboradores/Administradores presentes na interação",   "Nome completo", "Cargo", "Divisão", "Unidade de Negócios:", "Dados dos Agentes Públicos presentes na interação",
              "Nome do Agente Público", "Órgão", "Área de atuação do Agente Público", "Data da interação:", "Local da interação:", "Horário de Início:",
                        "Horário de Término:", "Quem motivou a interação?", "Assunto discutido:", "Justifique:", "Qual?", "Qual o tema? ",
              "Houve formalização da interação em ata ou trocas de ofícios?", "Anexar documentos:",
              "A interação incluiu o custeio de algum tipo de hospitalidade (refeição, deslocamento, etc?",
              "Comentários?", "Anexar documentos (opcional): ",
              "Houve algum pedido, sinalização, indicação, requerimento ou conduta imprópria por qualquer dos presentes?", "Comentários:",
              "Anexe aqui (Opcional): ", "Houve algum pedido, sinalização, indicação ou requerimento para doação ou patrocínio de algum projeto ou evento?",
              "Comentários:", "Anexe aqui ( Opcional):", "Deseja informar algo não listado neste formulário?", "Comentários:", "Anexe aqui (Opcional):",
              "Declaro que as informações acima são verídicas completas e corretas, me responsabilizando pelo conteúdo nela contido", ]

    headerStyles(ws, header)
    ws.append(header)
    addContent(ws, users)
    addStyles(ws, header)


def addContent(ws, users):
    line = 2
    for user in range(len(users)):
        line = line + 1
        for form in range(len(users[user]["forms"])):
            questions = users[user]["forms"][form]['questions']
            author = users[user]["forms"][form]["author"]

            infos = ["--", users[user]["id"], users[user]["first_name"], users[user]["email"],  users[user]["number_id"],	 users[user]["company"]["name"], users[user]["department"]["name"],	"--",
                     "--",	"--",	users[user]["active"], users[user]["created"],

                     users[user]["forms"][form]
                     ["status"]["title"],	      users[user]["forms"][form]
                     ["approved_date"],	author["first_name"],	     users[user]["forms"][form]
                     ["created"],	     users[user]["forms"][form]
                     ["approved_date"],	"--",	"--",	"--",	"--",	"--"]

            for i in range(32):
                try:
                    if (i == 3 or i == 5 or i == 6):
                        infos.append("--")
                        if (questions[str(i)]['value'] == [] and i != 6):
                            for i in range(4):
                                infos.append("--")
                        elif (questions[str(i)]['value'] == [] and i == 6):
                            for i in range(3):
                                infos.append("--")
                        elif (questions[str(i)]['value'] != None):
                            if i == 6:
                                name = ""
                                part = ""
                                area = ""
                                for value in questions[str(i)]['value']:
                                    name = name + value[0]['value'] + "\n"
                                    part = part + value[1]['value'] + "\n"
                                    area = area + value[2]['value'] + "\n"
                                infos.append(name)
                                infos.append(part)
                                infos.append(area)
                            else:
                                name = ""
                                office = ""
                                division = ""
                                unity = ""

                                for value in questions[str(i)]['value']:
                                    name = name + value[0]['value'] + "\n"
                                    office = office + value[1]['value'] + "\n"
                                    division = division + \
                                        value[2]['value'] + "\n"
                                    unity = unity + value[3]['value'] + "\n"
                                infos.append(name)
                                infos.append(office)
                                infos.append(division)
                                infos.append(unity)

                    else:
                        infos.append(questions[str(i)]['value'])

                except KeyError:
                    continue
            addCells(ws, infos)


def addCells(ws, infos):
    ws.append(infos)


def mergeCells(ws, startRow, endRow, startColumn, endColumn, initialCell, value, backgroundColor, fontColor, bold, fontSize):
    ws.merge_cells(start_row=startRow, start_column=startColumn,
                   end_row=endRow, end_column=endColumn)

    cell = ws[initialCell]
    cell.value = value
    cell.fill = PatternFill("solid", fgColor=backgroundColor)
    cell.alignment = Alignment(
        horizontal="center", vertical="center")
    cell.font = Font(bold=bold, color=fontColor,
                     size=fontSize, name='Calibri',)


def addHeaderColor(ws, row, firstColumn, lastColumn, color):
    for column in range(firstColumn, lastColumn):

        ws.cell(row, column).fill = PatternFill(start_color=color,
                                                end_color=color,
                                                fill_type='solid')
        ws.cell(row, column).font = Font(bold=False, color="00FFFFFF",
                                         size="12", name='Calibri',)


def addStyles(ws, header):
    addHeaderColor(ws, 2, 1, 13, '002f5496')
    addHeaderColor(ws, 2, 13, 23, '001f3864')
    addHeaderColor(ws, 2, 23, len(header) + 1, '004472c4')


def headerStyles(ws, header):
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 20

    mergeCells(ws, 1, 1, 1, 12, "A1",  "Dados do usuário",
               "002f5496", "00FFFFFF", True, "15")

    mergeCells(ws, 1, 1, 13, 22,
               "M1", "Dados da análise",
               "001f3864", "00FFFFFF", True, "15")

    mergeCells(ws, 1, 1, 23,  len(header), "W1",  "Formulário",
               "004472c4", "00FFFFFF", True, "15")

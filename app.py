import openpyxl
from PIL import Image, ImageDraw, ImageFont
import datetime

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_participantes = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horario = linha[5].value
    data_emissao = linha[6].value

    if isinstance (data_inicio, datetime.datetime):
        data_inicio = data_inicio.strftime('%d/%m/%Y')
    if isinstance(data_final, datetime.datetime):
        data_final = data_final.strftime('%d/%m/%Y')
    if isinstance(data_emissao, datetime.datetime):
        data_emissao = data_emissao.strftime('%d/%m/%Y')

    fonte_nome = ImageFont.truetype('./tahoma.ttf', 43)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 43)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 43)

    image = Image.open('./certificado_teste.jpg')

    desenhar = ImageDraw.Draw(image)

    desenhar.text((532, 432), nome_participantes, fill='#002D69', font=fonte_nome)
    desenhar.text((532, 495), nome_curso, fill='#002D69', font=fonte_geral)
    desenhar.text((680, 559), tipo_participacao, fill='#002D69', font=fonte_geral)
    desenhar.text((710, 624), str(carga_horario), fill='#002D69', font=fonte_geral)
    desenhar.text((1600, 432), str(data_inicio), fill='#002D69', font=fonte_data)
    desenhar.text((1658, 503), str(data_final),fill='#002D69', font=fonte_data)
    desenhar.text((570, 908), str(data_emissao),fill='#002D69', font=fonte_data)

    image.save(f'./{indice}_{nome_participantes}_certificado.png')
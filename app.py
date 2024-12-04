import openpyxl
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import os


# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

# Carregar as fontes
fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)  # Fonte para o nome do participante
fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)   # Fonte para detalhes gerais
fonte_data = ImageFont.truetype('./tahoma.ttf', 55)    # Fonte para datas

# Processar cada linha da planilha
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):  # A partir da segunda linha
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    carga_horaria = linha[5].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    data_emissao = linha[6].value

    # Converter datas para string se forem objetos de data
    if isinstance(data_inicio, datetime):
        data_inicio = data_inicio.strftime('%d/%m/%Y')
    if isinstance(data_final, datetime):
        data_final = data_final.strftime('%d/%m/%Y')
    if isinstance(data_emissao, datetime):
        data_emissao = data_emissao.strftime('%d/%m/%Y')

    # Abrir o modelo do certificado
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # Inserir os dados no certificado
    desenhar.text((1020, 827), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), tipo_participacao, fill='black', font=fonte_geral)

    desenhar.text((1480, 1182), str(carga_horaria), fill='black', font=fonte_geral)  # Convertido para string

    desenhar.text((750, 1770), data_inicio, fill='blue', font=fonte_data)
    desenhar.text((750, 1930), data_final, fill='blue', font=fonte_data)
    desenhar.text((2220, 1930), data_emissao, fill='blue', font=fonte_data)

    # Salvar o certificado gerado

    # Obtendo o caminho da pasta de downloads do usuário
    download_path = os.path.join(os.getcwd(), "download")

    # Garante que a pasta exista
    os.makedirs(download_path, exist_ok=True)

    # Nome do arquivo
    certificado_nome = f'certificado_{indice + 1}_{nome_participante}.png'

    # Caminho completo do arquivo
    file_path = os.path.join(download_path, certificado_nome)

    # Salvando o arquivo
    image.save(file_path)

    print(f'Certificado gerado: {file_path}')

    # certificado_nome = f'certificado_{indice + 1}_{nome_participante}.png'
    # image.save(certificado_nome)
    # print(f'Certificado gerado: {certificado_nome}')

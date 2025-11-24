import os, requests
from dotenv import load_dotenv
from datetime import datetime
import csv
from send_email import send_email
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from fpdf import FPDF

def ensure_csv_exists(filepath, fieldnames):
    if not os.path.isfile(filepath):
        with open(filepath, 'w', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()

def append_csv_row(filepath, data_dict):
    with open(filepath, 'a', encoding='utf-8') as file:
        fieldnames = list(data_dict.keys())
        writer = csv.DictWriter(file, fieldnames)
        writer.writerow(data_dict)

# Carrega as variáveis de ambiente
load_dotenv()

# Define filepath e fieldnames
filepath = 'weather_log.csv'
fieldnames = ['datetime', 'city', 'temperature', 'humidity', 'description']

# Verifica se há ou não o arquivo especificado
ensure_csv_exists(filepath=filepath, fieldnames=fieldnames)

# Monta a URL
city = input('Digite a cidade: ')
KEY = os.getenv('OPENWEATHER_API_KEY')
URL = f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid={KEY}&units=metric&lang=pt_br'

# Fazer o request e armazena o resultado em result_dict
try:
    response = requests.get(URL)
    data = response.json()
    if not response.ok:
        result_dict ={
            'error': data.get('message', 'Unknown error'),
            'status_code': response.status_code
        }
        exit()

    result_dict = {
            'datetime': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'city': data['name'],
            'temperature': data['main']['temp'],
            'humidity': data['main']['humidity'],
            'description': data['weather'][0]['description']
        }
except Exception as e:
    print(f'Erro inesperado: {e}')
    exit()

# Adiciona tudo ao CSV
append_csv_row(filepath=filepath, data_dict=result_dict)

# Envia por email
# send_email(result_dict)
    
# FASE A — Dados para Relatório
# 1. Abrir o CSV
with open('weather_log.csv', 'r', encoding='utf-8') as file:
    # 2. Ler todas as linhas
    reader = csv.DictReader(file)
    # 3. Jogar em uma lista chamada rows
    rows = []
    for line in reader:
        rows.append(line)

# FASE A — converter a coluna datetime para objeto real
for line in rows:
    date = line['datetime']
    date = datetime.strptime(date, '%d/%m/%Y %H:%M:%S')
    line['datetime'] = date

# FASE A — ordenar pelos valores reais de datetime
rows.sort(key=lambda line: line['datetime'])

# FASE B — Criar Excel do Relatório
# 1. Criar workbook e worksheet
report_wb = Workbook()
report_ws = report_wb.active
# 2. Escrever cabeçalhos
for pos, header in enumerate(fieldnames, start=1):
    report_ws.cell(row=1, column=pos).value = header

# 3. Preencher dados ordenados
linha_atual = 2
for line in rows:
    for pos, header in enumerate(fieldnames, start=1):
        report_ws.cell(row=linha_atual, column=pos).value = line[header]
    linha_atual += 1

# 4. Estilizar cabeçalhos
header_font = Font(bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
header_alignment = Alignment(horizontal='center', vertical='center')
medium_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

for pos in range(1, len(fieldnames) + 1):
    cell = report_ws.cell(row=1, column=pos)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
    cell.border = medium_border

# 5. Estilizar dados (alinhamento, bordas, zebra)
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

fill_odd = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_even = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')

align_left = Alignment(horizontal='left', vertical='center')
align_center = Alignment(horizontal='center', vertical='center')

for row in range(2, report_ws.max_row + 1):
    for col in range(1, len(fieldnames) + 1):
        cell = report_ws.cell(row=row, column=col)
        cell.border = thin_border

        # Escolhe alinhamento com base no nome da coluna
        header = fieldnames[col - 1]
        if header in ['city', 'description']:
            cell.alignment = align_left
        else:
            cell.alignment = align_center

        # Linhas zebradas
        if row % 2 == 0:
            cell.fill = fill_even
        else:
            cell.fill = fill_odd


# 6. Ajustar largura das colunas
for col in range(1, len(fieldnames) + 1):
    max_length = 0
    for row in range(1, report_ws.max_row + 1):
        cell = report_ws.cell(row=row, column=col)
        if len(str(cell.value)) > max_length:
            max_length = len(str(cell.value))
    col_letter = get_column_letter(col)
    report_ws.column_dimensions[col_letter].width = max_length + 2 

# 7. Congelar linha do cabeçalho
report_ws.freeze_panes = 'A2'
# 8. Ativar filtros nas colunas
report_ws.auto_filter.ref = report_ws.dimensions
# 9. Salvar como "weather_report.xlsx"
report_wb.save('weather_report.xlsx')


# FASE C — Criar PDF do Relatório

def create_pdf_report(rows):
    # 1 - Criar o objeto PDF
    pdf = FPDF()
    # 2 - Criar uma página
    pdf.add_page()
    # 3 - Definir fonte
    pdf.set_font(family='Arial', size=16)
    # 4 - Escrever o título do relatório
    pdf.cell(0, 14, 'Weather Report - Consultation History', ln=1, align='C')
    # 5 - Escrever data/hora de geração
    pdf.set_font(family='Arial', size=10)
    pdf.cell(0, 10, f'Generated in: {datetime.now().strftime("%d/%m/%Y %H:%M")}', ln=1)
    pdf.ln(5)
    # 6 - Escrever o resumo:
    # 6.1 - Total de registros no histórico
    pdf.cell(0, 10, f'Total records: {len(rows)}', ln=1)
    # 6.2 - Data da primeira medição
    first = rows[0]
    pdf.cell(0, 10, f'First consultation: {first["datetime"].strftime("%d/%m/%Y - %H:%M")}', ln=1)
    # 6.3 - Data da última medição
    last = rows[-1]
    pdf.cell(0, 10, f'Last consultation: {last["datetime"].strftime("%d/%m/%Y - %H:%M")}', ln=1)
    pdf.ln(4)
    # 6.4 - Últimas 3 medições (datetime, city, temp, description)
    last_consultations = []
    if len(rows) >= 1:
        last_consultations.append(rows[-1])
    if len(rows) >= 2:
        last_consultations.append(rows[-2])
    if len(rows) >= 3:
        last_consultations.append(rows[-3])

    pdf.set_font(family='Arial', size=14)
    pdf.cell(0, 10, 'Last three consultations', ln=1, align='C')
    pdf.set_font(family='Arial', size=10)

    for item in last_consultations:
        text = f'{item["datetime"].strftime("%d/%m/%Y - %H:%M")} | {item["city"]}'
        text2 = f'{item["temperature"]}°C | {item["humidity"]}% | {item["description"]}'
        pdf.cell(0, 10, text, ln=1 )
        pdf.cell(0, 10, text2, ln=1)
        pdf.ln(4)
    # 7 - Salvar arquivo "weather_report.pdf"
    pdf.output(f'weather_report_{datetime.now().strftime("%Y-%m-%d-%H-%M-%S")}.pdf')


create_pdf_report(rows)
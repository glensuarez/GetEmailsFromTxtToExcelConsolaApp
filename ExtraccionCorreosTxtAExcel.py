import re
import openpyxl

# Abrir archivo de texto y leer su contenido
with open('archivo.txt', 'r', encoding='utf-8') as file:
    data = file.read()

# Crear una expresión regular para encontrar correos electrónicos
email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
emails = re.findall(email_pattern, data)

# Crear un archivo Excel y agregar correos electrónicos
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = 'Correos Electrónicos'
for index, email in enumerate(emails):
    sheet.cell(row=index+1, column=1, value=email)

# Guardar archivo Excel
workbook.save('correos_electronicosGmail12Mayo2023.xlsx')


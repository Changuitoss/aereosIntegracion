#! python3

import datetime, openpyxl

aereosWb = openpyxl.load_workbook('O:\--- AEREOS INTEGRACION ---.xlsx')

sheets = aereosWb.sheetnames

now = datetime.datetime.now()


for e in sheets:
    sheet = aereosWb[e]
    for i in range(sheet['B6'].row, sheet.max_row + 1, 2):
        valor = sheet['B' + str(i)].value
        if valor != None:
            salida = datetime.datetime.strptime(sheet['B' + str(i)].value, '%d/%m/%Y')
            diasAntes30 = salida - datetime.timedelta(days=30)
        if diasAntes30 <= now:
            print('Revertir el aéreo ' + sheet.title + ' ' + str(salida))
        else:
            continue

#TODO: Evitar los aereos que ya fueron revertidos

#TODO: Enviar un mensaje de alguna forma a Cata

#TODO: Q lo chequee todos los dias automaticamente (el scheduler de windows no?)



        
    

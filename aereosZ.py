#! python3

import datetime, openpyxl, pyautogui, pyperclip

aereosWb = openpyxl.load_workbook('O:\--- AEREOS INTEGRACION ---.xlsx')

sheets = aereosWb.sheetnames

now = datetime.datetime.now()

direccion = 'cdandolo@amichi.com.ar'
asunto = 'Revertir *Z INTEGRACION'
aereosRevertir = []


for e in sheets:
    sheet = aereosWb[e]
    for i in range(sheet['B6'].row, sheet.max_row + 1, 2):
        valor = sheet['B' + str(i)].value
        estado = sheet['O' + str(i)].value
        if estado == 'REVERTIDO':
            continue
        if valor != None:
            salida = datetime.datetime.strptime(sheet['B' + str(i)].value, '%d/%m/%Y')
            diasAntes30 = salida - datetime.timedelta(days=30)
        if diasAntes30 <= now:
            aereosRevertir.append('Revertir el aéreo ' + sheet.title + ' ' + str(salida))
        else:
            continue
        
aereosWb.close()
        
#TODO: Evitar los aereos que ya fueron revertidos

def enviarMail():
    mailProc = subprocess.Popen('C:\Program Files\Microsoft Office\OFFICE11\OUTLOOK.EXE')
    pyautogui.click(36, 32)  #nuevo
    pyperclip.copy(direccion)
    pyperclip.paste(direccion)
    pyautogui.click(133, 157) #asunto
    pyperclip.copy(asunto)
    pyperclip.paste(asunto)
    pyautogui.click(66, 270)  #cuerpo
    pyperclip.copy(aereosRevertir)
    pyperclip.paste(aereosReverir)
    pyautogui.click(31, 55)
    
    
    


#TODO: Enviar un mensaje de alguna forma a Cata

#TODO: Q lo chequee todos los dias automaticamente (el scheduler de windows no?)



        
    

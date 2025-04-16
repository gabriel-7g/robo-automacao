import openpyxl 
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)


# Ler planilha e guardar informações
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # https://web.whatsapp.com/send?phone=555555&text=
    mensagem = f'Olá {nome}, o vencimento da sua internet ja passou do prazo com a data de {vencimento.strftime('%d/%m/%Y')} por favor pagar na seguinte chave pix '
    # aqui você pode colocar outra mensagem para complementar exemplo: pix = 'Mande nesse código pix 000000'
    mensagem_completa = mensagem # mensagem_completa = mensagem + pix 
    
    # Criar links personalizados do whatssap e enviar mensagens para cada cliente com base na planilha
    # Com base nos dados da planilha
    try:
        link_msg_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem_completa)}'
        webbrowser.open(link_msg_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.txt', 'a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')



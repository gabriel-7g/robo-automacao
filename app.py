import openpyxl 
from urllib.parse import quote
import webbrowser

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
    pix = "Pix cnpj " \
    "42514276000138 " \
    "MacielFibra ou Jenniffer mayara gomes tome maciel telecomunicacoes Banco Itau"
    mensagem_completa = mensagem + pix
    
    # Criar links personalizados do whatssap e enviar mensagens para cada cliente com base na planilha
    link_msg_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem_completa)}'


# Com base nos dados da planilha
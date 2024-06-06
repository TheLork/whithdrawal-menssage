import openpyxl
from urllib.parse import quote
import webbrowser
import time
import pyautogui

workbook = openpyxl.load_workbook('Retiradas_ Inad.xlsx')
pagina_clientes = workbook['Página1']

webbrowser.open('https://web.whatsapp.com')
time.sleep(10)

for linha in pagina_clientes.iter_rows(min_row=2):

    nome = linha[0].value
    telefone = linha[1].value
    veiculo = linha[2].value

    mensagem = f'Olá {nome}! Verificamos que o veículo {veiculo} ainda não passou pelo processo de retirada do equipamento. Qual seria o melhor dia e horário para agendarmos a visita do técnico?'
    
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        time.sleep(10)
        seta = pyautogui.locateCenterOnScreen('botao whatsapp.png')
        time.sleep(5)
        pyautogui.click(seta[0],seta[1])
        time.sleep(5)
        pyautogui.hotkey('ctrl','w')
        time.sleep(5)
    except:
        print(f'Não foi possível envviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}')
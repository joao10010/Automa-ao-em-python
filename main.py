import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os


webbrowser.open('https://web.whatsapp.com/')
print("Aguardando carregar o WhatsApp Web...")
sleep(15)


# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('Pasta1.xlsx')
pagina_clientes = workbook['Dia_5']


for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    mensagem = f'Olá {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.link_do_pagamento.com'

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        print(f"Enviando mensagem para {nome}...")
        sleep(15)  # Aumentei o tempo de espera
        seta = pyautogui.locateCenterOnScreen('enviar.png')
        print("Coordenadas do botão de enviar mensagem:", seta)  # Adicionei uma instrução de impressão
        if seta:
            # Modifiquei a linha abaixo para usar as coordenadas x e y do botão "Enviar"
            pyautogui.click(seta[0], seta[1])
            sleep(2)
            pyautogui.hotkey('ctrl','w')  # Fechar a aba do navegador
            print(f"Mensagem enviada para {nome}.")
        else:
            print(f"Não foi possível localizar o botão de enviar mensagem para {nome}.")
            raise Exception("Botão de enviar mensagem não encontrado.")
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {str(e)}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}\n')
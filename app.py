"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES, QUERO PODER MANDAR MENSAGENS DE COBRANÇA EM DETERMINADO DIA COM CLIENTES COM VENCIMENTO DIFERENTE
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
from tkinter import *

def iniciar_automacao():
  webbrowser.open('https://web.whatsapp.com')
  sleep(10)
  # Ler planilha e guardar informações sobre nome, telefone e data de vencimento
  workbook = openpyxl.load_workbook('clientes.xlsx')
  pagina_clientes = workbook['Sheet1']

  for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone e data de vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f"Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://www.link_do_pagamento.com"
  # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
  # com base nos dados da planilha
    
    try:
      link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
      webbrowser.open(link_mensagem_whatsapp)
      sleep(10)
      seta = pyautogui.locateCenterOnScreen('seta.png')
      sleep(5)
      pyautogui.click(seta[0], seta[1])
      sleep(5)
      pyautogui.hotkey('ctrl', 'w')
      sleep(5)
    except: 
      print(f'Não foi possivel enviar mensagem para {nome}')
      with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
        arquivo.write(f'{nome},{telefone}{os.linesep}')

def centralizar_janela(janela, largura, altura):
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


janela = Tk()
janela.title('Mensagens Automaticas Whatsapp')

largura_janela = 400
altura_janela = 300
janela.geometry(f"{largura_janela}x{altura_janela}")
centralizar_janela(janela, largura_janela, altura_janela)

texto_orientacao = Label(janela, text='Clique no botão para iniciar o envio de mensagens')
texto_orientacao.grid(column=0, row=0, padx=60, pady=70)

botao = Button(janela, text='Iniciar Envio de Mensagens', command=iniciar_automacao)
botao.grid(column=0, row=1, padx=60, pady=0)

janela.mainloop()
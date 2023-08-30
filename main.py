import pyautogui
from time import sleep
import pandas as pd
import PySimpleGUI as sg
import os

def load_file_excel(path, titulo):
    df = pd.read_excel('enderecos.xlsx')
    if 'e-mail' not in df.columns:
        sg.popup('A coluna `e-mail` não foi encontrada no arquivo excel selecionado', 
                 'Verifique o arquivo e execute o script novamente')
        exit()
    elif 'processo' not in df.columns:
        sg.popup('A coluna `processo` não foi encontrada no arquivo excel selecionado', 
                 'Verifique o arquivo e execute o script novamente')
        exit()
    else:
        for index, row in df.iterrows():
            enviar_mensagens(titulo + ' - Projeto ' + row['processo'], row['e-mail'])

def enviar_mensagens(titulo, email):
    pyautogui.click(170,213, duration=.5)
    pyautogui.click(1394,477, duration=.2)
    pyautogui.write(email, interval=.02)
    pyautogui.click(1280,520, duration=.2)
    sleep(.4)
    pyautogui.write(titulo, interval=.02)
    pyautogui.press('tab')
    sleep(.4)
    #colar mensagem
    pyautogui.hotkey('ctrl', 'v')

    #clicar em programar
    pyautogui.click(1354,999, duration=.2)
    pyautogui.click(1351,960, duration=.5)
    pyautogui.click(908,675, duration=.5)
    pyautogui.click(1126,735, duration=.5)

base_path = os.getcwd()
while True:
    sg.theme('TealMono')
    fonte = ("Arial", 16)
    layout = [
            [
                sg.Text('Programar envio de mensagens via mala direta'),
                sg.Button('Ajuda')
            ],[
                sg.Text('Selecione o arquivo Excel'), 
                sg.InputText(key='-FILE_EXCEL-'),
                sg.FileBrowse(initial_folder=base_path, file_types=[("Arquivos Excel", "*.xlsx")])
            ],[
                sg.Text('Título da mensagem'), 
                sg.InputText('Indicação de bolsista',key='-SUBJECT-'), 
            ],[
                sg.Button('Enviar mensagens', key='enviar_mensagens'), 
                sg.Button('Sair')
            ]
        ]
    janela = sg.Window('Teste', layout, font=fonte)

    eventos, valores = janela.read()
    janela.close()
    if eventos in (sg.WIN_CLOSED,'Sair'):
        print('Sair do aplicativo')
        janela.close()
        break
    elif eventos == 'Ajuda':
        file=open("ajuda.txt")
        text=file.read()
        sg.popup_scrolled(text, title="Scrolled Popup", size=(50,10))

    elif eventos == 'enviar_mensagens':
        if valores['-FILE_EXCEL-'] == '':
            sg.popup("Você deve selecionar um arquivo Excel")
        elif valores['-SUBJECT-'] == '':
            sg.popup("Você deve informar um título para a mensagem")
        else:
            print(valores)
            popup_result = sg.popup_ok_cancel('Antes de continuar, tenha certeza de ter preparado o ambiente. Siga as instuções:', 
                     'Deixe o e-mail do google aberto e em primeiro plano',
                     'Copie para a área de transferência a mensagem que será enviada para todos os destinatários.',
                     'Se tudo estiver pronto, pode clicar em OK',
                     'ou clique em Cancelar')
            if popup_result == 'OK':
                load_file_excel(valores['-FILE_EXCEL-'], valores['-SUBJECT-'])
            else:
                sg.popup('Você cancelou o envio das mensagens.')
            
            break


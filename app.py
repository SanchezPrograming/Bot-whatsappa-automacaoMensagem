import openpyxl
from urllib.parse  import quote
import webbrowser
from time import sleep
import pyautogui
import keyboard

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes  = workbook['Worksheet']
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    email = linha[1].value
    telefone = linha[2].value

    mensagem = f'Olá {nome}, Automação com Python utilizando a ferramenta PyAutoGUI https://github.com/SanchezPrograming'
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}' 
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10) # Esses 10s depende da velocidade do seu host e da sua Internet
    pyautogui.press('enter')  
    pyautogui.press('enter') 
    sleep(10) 
    pyautogui.hotkey('ctrl', 'w') # Aqui ele fara tudo de novo até voce cancelar
     
  
 

            








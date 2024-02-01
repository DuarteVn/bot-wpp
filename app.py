import openpyxl
import webbrowser
from urllib.parse import quote
from time import sleep
import pyautogui


webbrowser.open('https://web.whatsapp.com')
sleep(30)

workbook = openpyxl.load_workbook('cliente.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Bom dia {nome}, seu boleto vence dia {vencimento}. Favor enviar pix para este número 71991009828'
    # print(nome)
    # print(telefone)
    # print (vencimento)
    #Criar link personalizados do whatsapp e enviar mensagem para cada cliente
    
    try:
        link_msg_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_msg_wpp)
        sleep(15)  # Aumente o tempo de espera inicial, se necessário
        seta = pyautogui.locateCenterOnScreen('seta.PNG', confidence=0.8)
        if seta:
            pyautogui.click(seta[0], seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            sleep(5)
        else:
            print(f"Seta não encontrada. Verifique a imagem ou ajuste a sensibilidade.")
    except Exception as e:
        print(f"Não foi possível enviar mensagem para {nome}. Erro: {str(e)}")
    

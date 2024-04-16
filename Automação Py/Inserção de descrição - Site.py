import openpyxl
import pyautogui
import keyboard

workbook = openpyxl.load_workbook('gerar_descrição_chat_GPT.xlsx')
Arandelas_sheet = workbook['foco em mat elétrico']

for linha in Arandelas_sheet.iter_rows(min_row=2):


    if linha[11].value is not None:

        # Primeiro Processo - Gerar Descrição com IA

        pyautogui.click(1217,141, duration=0.5) 
        pyautogui.press('END')
        pyautogui.click(571,702, duration=0.5)
        keyboard.write(str(linha[11].value)) #Digitar a Para a IA
        pyautogui.press('ENTER') #Confirmar Perguntar
        pyautogui.sleep(25) #Esperar Resposta
        pyautogui.click(539,622, duration=0.5) #Copiar Resposta

        # Segundo Processo - Converter Para HTML

        pyautogui.click(370,16, duration=0.5) #Acessar Conversor
        pyautogui.click(122,48, duration=0.5) #Acessar Conversor
        pyautogui.click(721,393, duration=0.5) #Entrar no conversor
        pyautogui.hotkey('ctrl', 'v') #Colar SEO
        pyautogui.click(113,221, duration=0.5) #Entrar no Código-Fonte
        pyautogui.click(721,393, duration=0.5) #Entrar no conversor novamente
        pyautogui.hotkey('ctrl', 'home') #Acessar topo do Código
        pyautogui.press('down') #Descer Linha
        keyboard.write('&nbsp;<br />') #Adicionar o Espaçamento
        pyautogui.press('ENTER') #Adicionar Mais uma Linha
        keyboard.write('&nbsp;') #Adicionar o Espaçamento
        pyautogui.press('down', presses=2, interval=0.7) #Descer Duas Linhas
        keyboard.write('&nbsp;<br />') #Adicionar o Espaçamento
        pyautogui.press('ENTER') #Adicionar Mais uma Linha
        keyboard.write('&nbsp;') #Adicionar o Espaçamento
        pyautogui.press('down', presses=2, interval=0.7) #Descer Duas Linhas
        keyboard.write('&nbsp;<br />') #Adicionar o Espaçamento
        pyautogui.press('ENTER') #Adicionar Mais uma Linha
        keyboard.write('&nbsp;') #Adicionar o Espaçamento
        pyautogui.click(1995,53, duration=0.5) #Acessar HTML De Info. Tecnicas
        pyautogui.hotkey('ctrl', 'a') #Selecionar Todo o HTML de Info.
        pyautogui.hotkey('ctrl', 'c') #Copiar Todo o HTML de Info.
        pyautogui.click(721,393, duration=0.5) #Entrar no conversor novamente
        pyautogui.hotkey('ctrl', 'end') #Ir Para o Final do HTML
        pyautogui.press('ENTER') #Adiciona Mais Uma Linha
        pyautogui.hotkey('ctrl', 'v') #Colar o Complemento do HTML
        pyautogui.hotkey('ctrl', 'a') #Selecionar Todo o HTML
        pyautogui.hotkey('ctrl', 'c') #Copiar o HTML Pronto

        # Terceiro Processo - Fazer Upload para Wake

        pyautogui.click(570,16, duration=0.5) #Seleciona a Aba Correta da Wake
        pyautogui.click(732,385, duration=0.5) #Abre a Barra de Pesquisa
        keyboard.write(str(linha[1].value)) #Escreve Qual a Pesquisa desejada
        pyautogui.press('ENTER') #Confirma A Pesquisa Desejada
        pyautogui.sleep(2) #Espera Carregar Resultados de Pesquisa
        pyautogui.click(448,502, duration=0.5) #Seleciona o Produto Desejado
        pyautogui.sleep(1.5) #Espera Para Que Carregue a Página de Edição do Produto
        pyautogui.hotkey('ctrl', 'f') #Abre a Barra de Pesquisa
        keyboard.write('Vídeo', delay=0.5) #Escreve o Filtro de Pesquisa de Forma Lenta Para dar Tempo
        pyautogui.click(590,343, duration=0.5) #Abre no Código-Fonte
        pyautogui.click(934,580, duration=0.5) #Entra no Código-Fonte
        pyautogui.hotkey('ctrl', 'home') #Vai Para o Começo do Códiho
        pyautogui.press('ENTER', presses=2, interval=0.7) #Adicionar Duas Novas Linhas
        pyautogui.hotkey('ctrl', 'home') #Vai Para o Começo do Códiho
        pyautogui.hotkey('ctrl', 'v') #Cola o Código HTML Completo
        pyautogui.click(1168,286, duration=0.5)
        pyautogui.press('end') #Vai Para o Final da Página
        pyautogui.click(416,624, duration=0.5)
        pyautogui.sleep(2)
        pyautogui.click(137,11, duration=0.5)
        










        



        

        





        




        




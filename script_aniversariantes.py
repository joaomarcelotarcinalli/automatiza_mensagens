import pyautogui
import keyboard
import pandas
import time
from datetime import date

# realiza a leitura da planilha
planilha = pandas.read_excel('Planilha Aniversarios.xlsx')

# cria um loop para passar sobre todas as linhas
for linha in planilha.index:
    nome_cliente = planilha.loc[linha, "nome_cliente"] #declara a posição da coluna nome
    dia_nascimento = planilha.loc[linha, "dia_nasc"] # declara a posição da coluna data nascimento
    mes_nascimento = planilha.loc[linha, "mes_nasc"] # declara a posição da coluna telefone
    mensagem = planilha.loc[linha, "mensagem"] # declara a posição da coluna mensagem
    tel_cliente = planilha.loc[linha, "telefone_cliente"] # declara a posição da coluna telefone
    
    dt_atual = date.today() # cria a variavel da data atual

    dia_atual = dt_atual.day # declara o dia atual
    mes_atual = dt_atual.month # declara o mes atual

    if dia_nascimento == dia_atual and mes_nascimento == mes_atual: # se o dia atual for igual ao dia de nascimento, executa tal ação

        pyautogui.press('win')
        time.sleep(1)
        keyboard.write("canva")
        time.sleep(0.5)
        keyboard.press('enter')
        time.sleep(4)
        pyautogui.click(x=424, y=979) # clica no arquivo da imagem no canva
        time.sleep(4)
        pyautogui.click(clicks= 2, x=1214, y=414) # dois clique para alterar o texto da caixa de texto
        time.sleep(0.5) 
        pyautogui.hotkey("ctrl", "a") # seleciona tudo que está escrito na caixa de texto
        time.sleep(0.5)
        pyautogui.press('delete') # apaga tudo que está escrito
        time.sleep(1)
        keyboard.write(nome_cliente) # digita a mensagem que foi lida na planilha do excel
        time.sleep(2)
        pyautogui.click(x=1784, y=94) # clica no compartilhar
        time.sleep(2)
        pyautogui.click(x=1544, y=716) # clica em baixar
        time.sleep(3)
        pyautogui.click(x=1747, y=256)
        time.sleep(1)
        pyautogui.click(x=1879, y=322)
        time.sleep(2)
        pyautogui.click(x=1541, y=346)
        time.sleep(1)
        pyautogui.click(x=1672, y=682) # clica novamente na opção baixar
        time.sleep(5)
        pyautogui.click(x=489, y=65) # clica para pesquisar o caminho da pasta      
        time.sleep(0.5)
        keyboard.write(str('C:\\Users\\joaom\\OneDrive\\Documents')) # digita o caminho da pasta
        time.sleep(0.5)
        pyautogui.press('enter') # pesquisa a pasta
        time.sleep(2)
        pyautogui.click(x=475, y=463) # clica para dar nome ao arquivo
        time.sleep(0.5)
        keyboard.write(nome_cliente) # nome do arquivo é o foi lido da coluna nome da planilha
        time.sleep(1)
        pyautogui.press('enter') # salva o arquivo
        time.sleep(2)
        pyautogui.click(x=349, y=31) # fecha a aba de edição do canva
        time.sleep(1)
        pyautogui.click(x=1894, y=18) # fecha o canva app
        time.sleep(1)
        pyautogui.press('win') # presssiona a tecla windows
        time.sleep(1)
        pyautogui.write("google chrome") # digita google chrome
        time.sleep(1)
        pyautogui.press('enter') # pressiona tecla enter
        time.sleep(2)
        pyautogui.click(x=610, y=79) # refencia onde deve ser clicada
        pyautogui.write("http://app3.mktzap.com.br") # digita a url do site para pesquisa
        pyautogui.press('enter') # pressiona a tecla enter
        time.sleep(3)
        pyautogui.click(x=1260, y=516) # referencia do botao usuario
        pyautogui.write("jmtarcinalli@concretoimoveis.com.br") # digita o usuario
        time.sleep(1)
        pyautogui.click(x=1280, y=606) # referencia do botao de senha
        pyautogui.write("Conc1234") # digita a senha
        time.sleep(1)
        pyautogui.click(x=1319, y=700) # referencia do botao para entrar
        time.sleep(5)
        pyautogui.click(x=1801, y=295) # referencia do botao enviar
        time.sleep(1)
        pyautogui.click(x=1707, y=343) # referencia do botao mensagens whatsapp
        time.sleep(3)
        pyautogui.click(x=1403, y=406) # referencia para selecionar qual canal sera enviada a mensagem
        time.sleep(1)
        pyautogui.click(x=764, y=490) # seleciona a opção do canal ativo
        time.sleep(2)
        pyautogui.click(x=917, y=705) # clica no campo para inserir o telefone do cliente
        pyautogui.write(str(tel_cliente)) # lê os dados da planilha e insere os dados da coluna tel_cliente
        time.sleep(2)
        pyautogui.click(x=578, y=799) # clica no campo para inserir a mensagem
        keyboard.write(mensagem) # lê os dados da planilha e insere os dados da coluna mensagem
        time.sleep(3)
        pyautogui.press('tab') # pressiona a tecla tab
        time.sleep(1)
        pyautogui.press('enter') # pressiona a tecla enter
        time.sleep(2)
        pyautogui.click(x=489, y=65) # clica para pesquisar o caminho da pasta
        time.sleep(0.5)
        keyboard.write(str('C:\\Users\\joaom\\OneDrive\\Documents')) # pesquisa o caminho da pasta
        time.sleep(0.5)
        pyautogui.press('enter') # pressiona a tecla enter
        time.sleep(2)
        pyautogui.click(x=405, y=518) # clica para pesquisar o nome do arquivo
        time.sleep(0.5)
        keyboard.write(nome_cliente) # seleciona o arquivo assinatura(aguardar gerar o arquivo de aniversariantes)
        time.sleep(2)
        pyautogui.press('enter') # pressiona a tecla enter
        time.sleep(2)
        pyautogui.click(x=1356, y=895) # clica no botao para enviar a mensagem
        time.sleep(3)
        pyautogui.click(x=110, y=953) # clica no botao para abrir opção de deslogar
        time.sleep(1)
        pyautogui.click(x=92, y=897) #clica no botao para deslogar
        time.sleep(3)
        pyautogui.click(x=1882, y=17) # fecha aba do chrome

else:
    print("Fim do programa")

'''
    -> Automação de Sistemas e Processos com Python
        Para controle de custo, todo dia, seu chefe pede um relatório com todas as compras de mercadoria da empresa. O seu trabalho,
    como analista, é enviar um email para ele, assim que começar a trabalhar, com o gasto total, a quantidade de produtos compradas e
    o preço médio dos produtos.

'''
'''
    Passo a Passo: Lógica do nosso programa
        (1) Entrar no sistema da empresa
            Principais funções que utilizaremos
                . pyautogui.click
                . pyautogui.write
                . pyautogui.press 
                . pyautogui.hotkey
        (2) Fazer o login
        (3) Exportar a base de dados
        (4) Calcular os indicadores
        (5) Enviar um e-mail para meu chefe

    Pyautogui: permite automatizar o mause, teclado, monitor. Ou sej,a podemos automatizar tarefas diárias feitas pelo usuário
'''

#Importando as bibliotecas a serem usadas
import pyautogui
import time
import pandas as pd
import pyperclip

pyautogui.PAUSE = 1             #Tempo entre cada aplicação


#Passo (1)
#Entrando no microsoft edge
pyautogui.press('win')
pyautogui.write('Microsoft Edge')
pyautogui.press('enter')
pyautogui.hotkey('ctrl',"t")

#Entrando no site
pyautogui.write('https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema')
pyautogui.press('enter')

'''
    Criando uma ferramenta para fazer esperar um tempo para o site carregar
    time.sleep(5)           Vai criar uma janela de espera de 5s para ir pro proximo passo
'''


#Entrando na conta, colocando login e senha

# Passo(2)
pyautogui.click(x=1172,y=382)           #Devemos mudar para cada tipo de tela   #Clicando no input do login
pyautogui.write('Meu_login')            #Preenchendo o login
pyautogui.press('tab')                  #Mudando para a senha
pyautogui.write('Senha')                #Preenchendo a senha


pyautogui.click(x=1215,y=525)           #Clicando no botão de login     

'''
 Se precisar de scrollar a pagina
    pyautogui.scroll()
Pegando a posição: Podemos fazer o seguinte
    (1) Colocar o código de sleep para dar tempo para posicionarmos nosso mouse no local que queremos
        time.sleep(5)
        print(pyautogui.position())                 #diz a posição que meu mouse estar quando estou rodando
'''

# Passo(3)
pyautogui.click(x=366, y=291)   #Clicando no arquivo
pyautogui.click(x=2300, y=185)  #Clicando no 3 pontinhos
pyautogui.click(x=2055, y=627)  #Clicando para fazer donwload
time.sleep(5) 

'''
# Passo(4)
    Sempre que usar um caminho devemos colcoar antes das aspas um r
    Usando o separador de coluna compo ;
'''
tabela = pd.read_csv(r"C:\Users\Admin\Downloads\Compras.csv", sep=";")           
print(tabela)

#Adaptando o arquivo, mudando os valores de string para number (python usa . como decimal)
tabela["ValorFinal"] = tabela["ValorFinal"].replace(",", ".", regex=True).astype(float)

#TODO Criando os cálculos
total_gasto = tabela["ValorFinal"].sum()        #Fazenda a soma da coluna ValorFinal
quantidade = tabela["Quantidade"].sum()         #Fazenda a soma da coluna Quantidade
preco_medio= total_gasto/quantidade

print(total_gasto,quantidade)
print(quantidade)
print(preco_medio)

# Passo(5)

#Entrando no email:
pyautogui.press('win')
pyautogui.write('Microsoft Edge')
pyautogui.press('enter')
pyautogui.hotkey('ctrl',"t")
pyautogui.write('https://mail.google.com/mail/u/2/#inbox')
pyautogui.press('enter')

time.sleep(3)
pyautogui.click(x=149, y=203)       #Clicando na parte de criar um e-mail
pyautogui.click(x=1916, y=472)      #Clicando na parte de colocar um e-mail (Para)

pyautogui.write('matheusavila1006@gmail.com')          #Digitando o email
pyautogui.press('tab')                                 #Escolhendo o email 

'''
    Como não conseguimos escrever caracterres especiais com o .write vamos usar uma outra extensão do pyautogui para fazer isso
'''
pyautogui.press('tab')                                 #Alterando para o assunto
pyperclip.copy('Relatório de vendas')           
pyautogui.hotkey('ctrl',"v")                           #Digitando o assunto
  

pyautogui.press('tab')                                 #Alterando para o corpo do email
texto =  f'''
    Prezados,
    Segue o relatório de compras:
        Total Gasto = R${total_gasto:,.2f}
        Quantidade de Produto = {quantidade:,}
        Preço Médio = R${preco_medio:,.2f}
    
    Qualquer dúvida, é so falar.
    Att..

    Matheus Augusto
'''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl',"v") 


pyautogui.hotkey('ctrl',"enter")

# Proximos Desafios para pesquisar: 
# (1) Python rodar código automaticamente
# (2) Python criando um executável

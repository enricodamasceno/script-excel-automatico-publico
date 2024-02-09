from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# Ler a planilha excel
caminho_planilha = 'Caminho/Para/A/Planilha'
dados_planilha = pd.read_excel(caminho_planilha)

# Configurar o caminho para o executável do WebDriver
# driver_path = 'Caminho/Para/O/Executável'

# Inicializar o navegador
driver = webdriver.Chrome()  # para Chrome

# Navegar até a página de login
driver.get('http://meuendereço.com.html')

######################################
# Caso a página possua login e senha #
######################################

# Esperar até que o campo de login esteja clicável
campo_login = WebDriverWait(driver, 10).until(  
    EC.element_to_be_clickable((By.ID, 'id_username'))
)

login = 'meuLogin'

# Limpar o campo de login (se necessário)
driver.execute_script("arguments[0].value = '';", campo_login)

# Definir o valor do campo de login usando JavaScript
driver.execute_script("arguments[0].value = arguments[1];", campo_login, login)

# Encontrar o campo de senha e inserí-la
campo_senha = driver.find_element(By.ID, 'id_password')
senha = 'minhaSenha'
campo_senha.send_keys(senha)

# O botão entrar é o único botão da página, encontramos pelo nome e clicamos
botao_entrar = driver.find_element(By.NAME, 'button')
botao_entrar.click()
 
###########################################################################################
# No meu caso, uma vez que foi feito o login, eu podia acessar qualquer página do sistema #
###########################################################################################

# Navegar até a página do sistema web para o cadastro de itens da planilha
driver.get('http://endereçodecadastro.com.html')

# A minha planilha possuia esses valores abaixo. O interessante
# desse exercício é adaptar o driver para percorrer uma página
# web específica, com diferentes botões, inputs, etc. No meu
# caso, era necessário fazer login, cadastrar um novo nome e
# alterar dados de forma específica para cada caixa de texto.

# Iterar sobre as linhas da planilha
for indice, linha in dados_planilha.iterrows():
    nome = linha['NOME']
    valor = str(linha['VALOR'])
    data_pagamento = linha['PAGAMENTO']
    data_vencimento = linha['VENCIMENTO']
    descricao = linha['DESCRIÇÃO']
    
    # Adquirindo e clicando no botão de cadastrar nome
    botao_cadastrar_nome = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'saidaForm:j_idt87'))
    )
    botao_cadastrar_nome.click()

    # Adquirindo o campo nome
    campo_nome = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'dialogForm:descricaoInputText'))
    )

    # Preencher o campo nome
    campo_nome.send_keys(nome)

    # Adquirindo e clicando no botão salvar
    botao_salvar_nome = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'dialogForm:j_idt189'))
    )
    botao_salvar_nome.click()
    
    # Adquirindo os outros campos do formulário para preenchimento
    campo_valor = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'saidaForm:valorTotalInputText'))
    )
    
    # Selecionar todo o texto no campo de valor
    campo_valor.send_keys(Keys.END)
    for _ in range(15):
        # Apagar o texto
        campo_valor.send_keys(Keys.BACKSPACE)
        time.sleep(0.1)

    # Ir ao final do campo e digitar o valor
    campo_valor.send_keys(Keys.END)
    valor = valor.replace('.', ',')
    for char in valor:
        campo_valor.send_keys(char)
        time.sleep(0.1)

    # Selecionar o campo de data de pagamento
    campo_data_pagamento = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@id='saidaForm:dataRealizacaoPagamentoCalendar_input']"))
    )
    
    # Selecionar todo o texto no campo de data de pagamento
    campo_data_pagamento.send_keys(Keys.END)  # Mover o cursor para o final do texto
    for _ in range(10):
        campo_data_pagamento.send_keys(Keys.BACKSPACE)  # Pressionar a tecla Backspace para excluir o texto existente
        time.sleep(0.1)
    
    # Preencher a data de pagamento
    for char in data_pagamento:
        campo_data_pagamento.send_keys(char)
        time.sleep(0.1)
    
    # Selecionar o campo de data de vencimento
    campo_data_vencimento = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@id='saidaForm:dataSaidaInputText_input']"))
    )
    
    # Selecionar todo o texto no campo de data de vencimento
    campo_data_vencimento.send_keys(Keys.END)  # Mover o cursor para o final do texto
    for _ in range(10):
        campo_data_vencimento.send_keys(Keys.BACKSPACE) 
        time.sleep(0.1)

    # Preencher a data de vencimento
    for char in data_vencimento:
        campo_data_vencimento.send_keys(char)
        time.sleep(0.1)
    
    # Preencher o campo de descrição
    campo_descricao = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'saidaForm:descricaoInputTextarea'))
    )
    campo_descricao.clear()
    campo_descricao.send_keys(descricao)
 
    # Clicar no botão para enviar o formulário
    botao_enviar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'saidaForm:j_idt175'))
    )
    botao_enviar.click()

    # Voltar à página inicial para iterar novamente
    time.sleep(1)
    driver.get('http://meuendereço.com.html')

# Fechar o navegador ao final do script
driver.quit()
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime

# *Configurações de Data
hoje = datetime.date.today()
hoje_formatado = hoje.strftime('%d/%m/%Y')

# 1. Abrir o GPM
navegador = webdriver.Chrome()
navegador.get('https://endiconpa.gpm.srv.br/index.php')

# 1.1 Inserir credenciais de acesso - LOGIN
usuario = navegador.find_element_by_xpath('//*[@id="idLogin"]')
usuario.send_keys('SAMUEL.MENEZES')
time.sleep(1)

# 1.2 Inserir credenciais de acesso - SENHA
senha = navegador.find_element_by_xpath('//*[@id="idSenha"]')
senha.send_keys('trocar@senha')
time.sleep(1)

# 1.3 Inserir credenciais de acesso - LOGAR
botao_login = navegador.find_element_by_xpath('//*[@id="form_login"]/input[5]')
botao_login.click()
time.sleep(1)

# 2. Navegar até a aba de cadastro de veículos
url_cadastro = 'https://endiconpa.gpm.srv.br/cadastro/geral/veiculo.php'
navegador.get(url_cadastro)
time.sleep(2)

# 3. Acessar Planilha de Cadastro
workbook = openpyxl.load_workbook('C:\\temp\\atualizar_cadastro.xlsx')
planilha = workbook.active

# 4. Criar iteração do cadastros
for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]
    status = coluna[1]
    #velocidade = coluna[2]

# 5. Indicar os dados para serem alterados
# 5.1 Buscar placa do veículos
    buscar_veiculo = navegador.find_element_by_xpath('//*[@id="form"]/table/tbody/tr[2]/td[2]/div/input')
    buscar_veiculo.send_keys(veiculo)
    time.sleep(1)
# 5.2 Filtrar Status
    filtrar_status = navegador.find_element_by_xpath('//*[@id="form"]/table/tbody/tr[5]/td[4]/div/input')
    filtrar_status.send_keys('Todos')
# 5.2 Buscar modelo do veículo
    modelo = navegador.find_element_by_xpath('//*[@id="form"]/table/tbody/tr[2]/td[4]/div/input')
    modelo.send_keys(Keys.ENTER)
# 5.3 Abrir cadastro
    abrir_cadastro = navegador.find_element_by_xpath('//*[@id="tab_resultados"]/tbody/tr/td[1]/a[1]/img')
    abrir_cadastro.click()
# 5.4 Indicar status
    buscar_status = navegador.find_element_by_xpath('//*[@id="form"]/table[1]/tbody/tr[10]/td[2]/div/input')
    buscar_status.send_keys(status)
    time.sleep(2)
# 5.5 Velocidade chassis
    chassis = navegador.find_element_by_xpath('//*[@id="chassi"]')
    chassis.clear()
    chassis.send_keys(Keys.TAB)
    time.sleep(2)
# 5.6 Limpar Renavam
    renavam = navegador.find_element_by_xpath('//*[@id="renavam"]')
    renavam.clear()
    renavam.send_keys(Keys.ENTER)
    time.sleep(2)


















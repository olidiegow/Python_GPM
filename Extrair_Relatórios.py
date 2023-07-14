import time
import os
import shutil
import datetime
import rarfile
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

navegador = webdriver.Chrome()
navegador.maximize_window()
pasta_download = 'C:\\Users\\diego.oliveira\\Downloads'

# Abrir o GPM
navegador.get('https://endiconpa.gpm.srv.br/index.php')

# Inserir credenciais de Acesso
usuario = navegador.find_element_by_xpath('//*[@id="idLogin"]')
usuario.send_keys('SAMUEL.MENEZES')
time.sleep(1)

senha = navegador.find_element_by_xpath('//*[@id="idSenha"]')
senha.send_keys('trocar@senha')
time.sleep(1)

botao_login = navegador.find_element_by_xpath('//*[@id="form_login"]/input[5]')
botao_login.click()
time.sleep(1)

# Acessar aba de relatório de Segurança
acessar_aba = navegador.find_element_by_xpath('//*[@id="7000"]/a')
acessar_aba.click()
time.sleep(1)

# Acessar Checklist diario em nova aba
url_checklist = 'https://endiconpa.gpm.srv.br/gpm/geral/checklist_perguntas_respostas.php'
navegador.execute_script("window.open('https://endiconpa.gpm.srv.br/gpm/geral/checklist_perguntas_respostas.php', '_blank');")
time.sleep(1)
identificar_abas = navegador.window_handles
navegador.switch_to.window(identificar_abas[1])

# Preencher formulario
hoje = datetime.date.today()
hoje_formatado = hoje.strftime('%d/%m/%Y')
amanha = hoje + datetime.timedelta(days=1)
amanha_formatado = amanha.strftime('%d/%m/%Y')
ontem = hoje - datetime.timedelta(days=1)
ontem_formatado = ontem.strftime('%d/%m/%Y')
ultima_semana = hoje - datetime.timedelta(days=7)
semana_formatada = ultima_semana.strftime('%d/%m/%Y')
fim_semana = hoje - datetime.timedelta(days=4)
fds_formatado = fim_semana.strftime('%d/%m/%Y')
parametro_data = fds_formatado


data_inicio = navegador.find_element_by_id('id_data_in')
data_inicio.send_keys(parametro_data)
time.sleep(1)

data_fim = navegador.find_element_by_id('id_data_out')
data_fim.send_keys(hoje_formatado, Keys.TAB, Keys.TAB)
time.sleep(1)

opcoes_finalidade = navegador.find_element_by_xpath('//*[@id="final_chosen"]')
opcoes_finalidade.click()
escolher_finalidade = navegador.find_element_by_xpath('//*[@id="final_chosen"]/div/div/input')
escolher_finalidade.send_keys('1', Keys.ENTER)
time.sleep(1)

opcoes_tpchecklist = navegador.find_element_by_xpath('//*[@id="id_tp_form_chosen"]')
opcoes_tpchecklist.click()
escolher_tpchecklist = navegador.find_element_by_xpath('//*[@id="id_tp_form_chosen"]/div/div/input')
escolher_tpchecklist.send_keys('CAMINHÃO', Keys.ENTER)
time.sleep(1)

# Exportar Para o Excel
exportar_excel = navegador.find_element_by_xpath('/html/body/div[1]/span/a')
exportar_excel.click()
time.sleep(5)

arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]

ultimo_download = max(arquivos, key=lambda x: x[1])[0]
destino_indisponibilidade = 'Z:\\07.PowerBI\\Dados\\Frota\\3.GPM\\4.Abertura'
#shutil.move(ultimo_download, destino_indisponibilidade)
print("Arquivo Indisp.: ", ultimo_download)


# Acessar Tratamento de Pendencias Frota
url_pendencias = 'https://endiconpa.gpm.srv.br/gpm/geral/consulta_check_pendencia.php?tip=C'
navegador.get(url_pendencias)
time.sleep(2)
escolher_finalidade2 = navegador.find_element_by_xpath('//*[@id="tab_form"]/tbody/tr[1]/td[2]/div/input')
escolher_finalidade2.send_keys('Veicular - turno ', Keys.TAB)
time.sleep(1)

# Filtrar Pendencias
data_inicio2 = navegador.find_element_by_xpath('//*[@id="data_exec_inicial"]')
data_inicio2.send_keys(parametro_data, Keys.TAB)
time.sleep(1)

data_fim2 = navegador.find_element_by_xpath('//*[@id="data_exec_final"]')
data_fim2.send_keys(ontem_formatado, Keys.TAB)
time.sleep(1)

# Exportar Pendencias
botao_pesquisar = navegador.find_element_by_xpath('//*[@id="tab_form"]/tbody/tr[11]/td/input')
botao_pesquisar.click()
time.sleep(5)

exportar_excel2 = navegador.find_element_by_xpath('//*[@id="tab_resultados_wrapper"]/div[1]/button[3]')
exportar_excel2.click()
time.sleep(5)

# Salvar Pendencias
arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]
ultimo_download = max(arquivos, key=lambda x: x[1])[0]
destino_pendencias = 'Z:\\07.PowerBI\\Dados\Frota\\3.GPM\\1.Pendencias'
#shutil.move(ultimo_download, destino_pendencias)
print("Arquivo Pendencias: ", ultimo_download)

# Acessar Multas
url_multas = 'https://endiconpa.gpm.srv.br/cadastro/geral/cad_multa_veiculo.php'
navegador.get(url_multas)
time.sleep(2)

# Exportar Multas
botao_pesquisar2 = navegador.find_element_by_xpath('//*[@id="form1"]/div[2]/input')
botao_pesquisar2.click()
time.sleep(1)

exportar_excel3 = navegador.find_element_by_xpath('//*[@id="tab_resultados_wrapper"]/div[1]/button[3]')
exportar_excel3.click()
time.sleep(1)

# Salvar Multas
arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]
ultimo_download = max(arquivos, key=lambda x: x[1])[0]
destino_multas = 'Z:\\07.PowerBI\\Dados\\Frota\\3.GPM\\3.Multas'
shutil.move(ultimo_download, destino_multas)
print("Arquivo Multa: ", ultimo_download)


# Acessar Avarias
url_avarias = 'https://endiconpa.gpm.srv.br/gpm/geral/consulta_avaria.php'
navegador.get(url_avarias)
time.sleep(3)

selecionar_todos = navegador.find_element_by_xpath('//*[@id="fomulario"]/table[2]/tbody/tr[6]/td[2]/label[3]/input')
selecionar_todos.click()

botao_pesquisar3 = navegador.find_element_by_xpath('//*[@id="fomulario"]/div/input')
botao_pesquisar3.click()

exportar_excel4 = navegador.find_element_by_xpath('//*[@id="tab_resultados_wrapper"]/div[1]/button[3]')
exportar_excel4.click()
time.sleep(3)

# Salvar avarias
arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]
ultimo_download = max(arquivos, key=lambda x: x[1])[0]
destino_avarias = 'Z:\\07.PowerBI\\Dados\\Frota\\3.GPM\\2.Avarias'
shutil.move(ultimo_download, destino_avarias)
print("Arquivo Avarias: ", ultimo_download)


# Acessar Tratamento de Pendencias Segurança
pendencias_sesmt = 'https://endiconpa.gpm.srv.br/gpm/geral/consulta_check_pendencia.php?tip=C'
navegador.get(pendencias_sesmt)
time.sleep(2)
escolher_finalidade3 = navegador.find_element_by_xpath('//*[@id="tab_form"]/tbody/tr[1]/td[2]/div/input')
escolher_finalidade3.send_keys('Seguranca - turno', Keys.TAB)
time.sleep(1)

# Filtrar Pendencias de Segurança
data_inicio2 = navegador.find_element_by_xpath('//*[@id="data_exec_inicial"]')
data_inicio2.send_keys(parametro_data, Keys.TAB)
time.sleep(1)

data_fim2 = navegador.find_element_by_xpath('//*[@id="data_exec_final"]')
data_fim2.send_keys(ontem_formatado, Keys.TAB)
time.sleep(1)

# Exportar Pendencias de Segurança
botao_pesquisar1 = navegador.find_element_by_xpath('//*[@id="tab_form"]/tbody/tr[11]/td/input')
botao_pesquisar1.click()
time.sleep(5)

exportar_excel5 = navegador.find_element_by_xpath('//*[@id="tab_resultados_wrapper"]/div[1]/button[3]')
exportar_excel5.click()
time.sleep(5)


# Salvar Pendencias
arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]
ultimo_download = max(arquivos, key=lambda x: x[1])[0]
destino_pendencias_sesmt = 'Z:\\07.PowerBI\\Dados\\Segurança\\Pendencias'
shutil.move(ultimo_download, destino_pendencias_sesmt)
print("Arquivo Pendencias SESMT: ", ultimo_download)


navegador.quit()
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from PIL import Image
from io import BytesIO
import tempfile
import pandas as pd
import webbrowser
import datetime
import time
import os


usuario = input("Qual diretório será utilizado, person1 ou person2?")

person1 = os.path.expanduser("~") + r"\Business_name\person1\@atuais\Relatorios\Indústrias"
person2 = os.path.expanduser("~") + r"\Business_name\person2\@atuais\Relatorios\Indústrias"

if usuario.lower() == "person1":
    caminho = person1
else: 
    caminho = person2
    

cliente = str(input("Qual nome do cliente?"))


# Descobrindo o primeiro dia do mês anterior
data_atual = datetime.date.today()

if data_atual.month == 1:
    ano_atual = data_atual.year - 1
    mes_atual = 12
else:
    ano_atual = data_atual.year
    mes_atual = data_atual.month

ultimo_dia_mes_anterior = datetime.date(ano_atual, mes_atual, 1) - datetime.timedelta(days=1)
primeiro_dia_mes_anterior = datetime.date(ultimo_dia_mes_anterior.year, ultimo_dia_mes_anterior.month, 1)

dia_01 = primeiro_dia_mes_anterior.strftime("%d/%m/%Y")
dia_30 = ultimo_dia_mes_anterior.strftime("%d/%m/%Y")
print(dia_01)
print(dia_30)


# Definindo nome do cliente e caminho para descobrir o URL correto
base_cliente = caminho + "\\" + cliente + "\\2023" + "\\" + cliente + " 2023.xlsm"
aba_dashboard = pd.read_excel(base_cliente, sheet_name="Dashboard")
url = aba_dashboard.iloc[24,2]


# Definindo produto a ser pesquisado (USAR A 7ª COLUNA COMO REFERÊNCIA 2ª LINHA)
base_cliente = caminho + "\\" + cliente + "\\2023" + "\\" + cliente + " 2023.xlsm"
top_vitrines = pd.read_excel(base_cliente, sheet_name="Top Vitrines", skiprows=2)



# Configurando o Selenium para utilizar o Google Chrome
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')  # Define a janela maximizada
driver = webdriver.Chrome(options=options)

# Definindo site correto para acessar
standout = "https://www.mybusiness.com.br" + url + "/admin/site/login"
driver.get(standout)


# fazendo login 
usuario = '//*[@id="LoginForm_username"]'
driver.find_element('xpath', usuario).send_keys('eduarda.andrade')

senha = '//*[@id="LoginForm_password"]'
driver.find_element('xpath', senha).send_keys('80976759')

click_login = '//*[@id="login-form"]/button'
driver.find_element('xpath', click_login).click()

time.sleep(40)


#Acessando a página para geração do mapa de calor
relatorios = '//*[@id="menu"]/li[2]/a'
driver.find_element('xpath', relatorios).click()
time.sleep(3)

mp_calor = '//*[@id="frmReport"]/div/div/li[15]/a/span'
driver.find_element('xpath', mp_calor).click()
time.sleep(4)

for i, produto in enumerate(top_vitrines.iloc[:, 7]):

    campo_nome = '//*[@id="s2id_product"]/a/span'
    driver.find_element('xpath', campo_nome).click()
    time.sleep(2)
    driver.find_element('xpath', '//*[@id="select2-drop"]/div/input').send_keys(produto)
    driver.find_element('xpath', '//*[@id="select2-drop"]/ul/li[1]/div').click()
    time.sleep(1)

    data_i = '//*[@id="startDate"]'
    driver.find_element('xpath', data_i).click()
    driver.find_element('xpath', data_i).send_keys(dia_01)
    driver.find_element('xpath', data_i).send_keys(Keys.TAB)

    data_f = '//*[@id="endDate"]'
    driver.find_element('xpath', data_f).click()
    driver.find_element('xpath', data_f).send_keys(dia_30)
    driver.find_element('xpath', data_f).send_keys(Keys.TAB)

    gerar = '//*[@id="btGenerate"]'
    driver.find_element('xpath', gerar).click()
    time.sleep(10)

    # Obter as dimensões da página inteira
    width = driver.execute_script("return document.documentElement.scrollWidth")
    height = driver.execute_script("return document.documentElement.scrollHeight")

    # Criando uma lista vazia para armazenar as imagens
    screenshot_images = []  

    # Definindo a altura atual como zero e definindo a altura que está visível na página
    scroll_height = 0
    window_height = driver.execute_script("return window.innerHeight")


    while scroll_height < height:

        # tira print da página
        screenshot = driver.get_screenshot_as_png()
        screenshot_image = Image.open(BytesIO(screenshot)) #cria uma instancia a partir dos dados da imagem capturada

        # Adiciona a captura à lista
        screenshot_images.append(screenshot_image)

        # Executa um script JS para rolar a página verticalmente 
        driver.execute_script(f"window.scrollTo(0, {scroll_height + window_height});")
        time.sleep(1)

        # Atualiza a altura de rolagem 
        scroll_height += window_height


    # Remove a última imagem duplicada
    screenshot_images.pop()

    #  calcula a altura total da imagem final somando a altura de cada imagem capturada na lista
    total_height = sum(image.height for image in screenshot_images)

    # Cria uma nova imagem com as dimensões corretas
    final_image = Image.new('RGB', (width, total_height))
    y_offset = 0

    #cola as imagens de forma vertical
    for image in screenshot_images:
        final_image.paste(image, (0, y_offset))
        y_offset += image.height

    element_xpath = '//*[@id="heatmapArea"]/canvas'  
    element = driver.find_element('xpath', element_xpath)
    location = element.location
    size = element.size

    # Calcular as coordenadas de corte com base na posição e tamanho do elemento
    crop_left = location['x']
    crop_top = location['y']
    crop_right = crop_left + size['width']
    crop_bottom = crop_top + size['height']

    # Ajustar as coordenadas de corte se ultrapassarem as dimensões da imagem final
    crop_right = min(crop_right, width)
    crop_bottom = min(crop_bottom, total_height)

    # Corta a imagem baseado nas coordenadas calculadas
    cropped_image = final_image.crop((crop_left, crop_top, crop_right, crop_bottom))
    
    
    cropped_image.show()

    
    # Pegando caminho para salvar o arquivo, e nome correto do arquivo a ser salvo
    referencia = top_vitrines.iloc[i,6]
    delimitador = "\\"
    caminho_arquivo = referencia.rsplit(delimitador, 1)[0].strip()
    referencia = top_vitrines.iloc[i,6]
    nome_arquivo = referencia.rsplit(delimitador, 1)[1].strip()

    
    save_path = os.path.join(caminho_arquivo, nome_arquivo + ".jpg")
    cropped_image.save(save_path, 'JPEG')

    
    # Retornando para página anterior
    driver.back()
    
    limpar_data = driver.find_element('xpath', data_i)
    limpar_data.clear()
    
    limpar_data = driver.find_element('xpath', data_f)
    limpar_data.clear()

driver.quit()

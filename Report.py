import pandas as pd
import time
import os
import shutil
import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

usuario = input("Qual diretorio deverá ser utilizado, Duda ou Raia?")

duda = os.path.expanduser("~") + r"\Business_name\person1\@atuais\Relatorios\Indústrias"
raia = os.path.expanduser("~") + r"\Business_name\person2\@atuais\Relatorios\Indústrias"

if usuario.lower() == "person1":
    caminho = person1
else:
    caminho = person2
    

tabela = pd.read_excel(caminho + r'\Base_Bot.xlsx', sheet_name='Sheet1')

if 'Status' not in tabela.columns:
    tabela['Status'] = ""
    
    
primeira_linha = tabela[tabela['Status'] != 'OK'].index

if not primeira_linha.empty:
    i = primeira_linha[0]
   
    navegador = webdriver.Chrome()
    navegador.get('https://www.mybusiness.com.br/admin/')


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

    dia_15 = datetime.date(ultimo_dia_mes_anterior.year, ultimo_dia_mes_anterior.month, 15)
    dia_15 = dia_15.strftime("%d/%m/%Y")

    dia_16 = datetime.date(ultimo_dia_mes_anterior.year, ultimo_dia_mes_anterior.month, 16)
    dia_16 = dia_16.strftime("%d/%m/%Y")

    time.sleep(50)

    navegador.find_element("xpath", '//*[@id="menu"]/li[2]/a').click()
    time.sleep(1)
    navegador.find_element("xpath", '//*[@id="content"]/div[6]/div/ul/li[1]/a').click()
    time.sleep(1)



    for i in tabela[tabela['Status'] != 'OK'].index:
        try:
            cliente = tabela.at[i, 'Nome']        
            navegador.find_element("xpath", '//*[@id="s2id_customer"]/a/span').click()
            time.sleep(1)
            navegador.find_element("xpath", '//*[@id="select2-drop"]/div/input').send_keys(cliente)
            navegador.find_element("xpath", '//*[@id="select2-drop"]/ul/li/div').click()

            # Data inicial
            navegador.find_element("xpath", '//*[@id="startDate"]').click()
            navegador.find_element("xpath", '//*[@id="startDate"]').send_keys(str(dia_01))

            # Data final
            data_final = '//*[@id="endDate"]'
            navegador.find_element("xpath", data_final).click()
            navegador.find_element("xpath", data_final).send_keys(str(dia_30))
            navegador.find_element("xpath", data_final).send_keys(Keys.ENTER)
            time.sleep(3)

            # Clique na setinha
            seta = '//a[contains(@class, "btn") and contains(@class, "btn-primary") and contains(@class, "pull-right") and @data-toggle="dropdown"]'
            navegador.find_element("xpath", seta).click()

            # Clique em exportar
            navegador.find_element("xpath", '//*[@id="divExport"]/ul/li/a').click()
            time.sleep(90)

            # limpando os campos de data
            limpar_data1 = navegador.find_element("xpath", '//*[@id="startDate"]')
            limpar_data1.clear()

            limpar_data2 = navegador.find_element("xpath", '//*[@id="endDate"]')
            limpar_data2.clear()

            pasta_destino = tabela.iloc[i, 2]  # pega a mesma linha de referência da variável cliente. Porém com número de coluna travado

            # Caminho da pasta de downloads
            downloads_path = os.path.expanduser("~") + "/Downloads"

            # Capturando último arquivo baixado
            arquivos_downloads = os.listdir(downloads_path)
            arquivos_csv = [arquivo for arquivo in arquivos_downloads if arquivo.endswith(".csv")]
            arquivos_ordenados = sorted(arquivos_csv, key=lambda x: os.path.getmtime(os.path.join(downloads_path, x)),reverse=True)
            if arquivos_ordenados:

                arquivos_ordenados != None

                primeiro_arquivo = arquivos_ordenados[0]

                # Caminho completo do arquivo baixado
                caminho_arquivo = os.path.join(downloads_path, primeiro_arquivo)

                # Renomeando o arquivo movido
                nome_arquivo = tabela.iloc[i, 3]  # Obtém o nome do arquivo da coluna correspondente
                novo_caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
                shutil.move(caminho_arquivo, novo_caminho_arquivo)

                limpar_data1 = navegador.find_element("xpath", '//*[@id="startDate"]')
                limpar_data1.clear()

                limpar_data2 = navegador.find_element("xpath", '//*[@id="endDate"]')
                limpar_data2.clear()

                tabela.at[i, 'Status'] = 'OK'
                tabela.to_excel(caminho + r'\Base_Bot.xlsx', index=False)      

                verificacao = tabela[tabela['Status'] != 'OK'].index

                if not verificacao.empty:
                    i = verificacao[0]
                else:
                    exit()

            else:

                campo_cliente = '//*[@id="startDate"]'
                navegador.find_element("xpath", campo_cliente).click()
                navegador.find_element("xpath", campo_cliente).send_keys(cliente)

                navegador.find_element("xpath", '//*[@id="startDate"]').click()
                navegador.find_element("xpath", '//*[@id="startDate"]').send_keys(str(dia_01))

                data_final = '//*[@id="endDate"]'
                navegador.find_element("xpath", data_final).click()
                navegador.find_element("xpath", data_final).send_keys(str(dia_15))
                navegador.find_element("xpath", data_final).send_keys(Keys.ENTER)

                seta = '//a[contains(@class, "btn") and contains(@class, "btn-primary") and contains(@class, "pull-right") and contains(@class, "dropdown-toggle") and @data-toggle="dropdown"]'
                navegador.find_element("xpath", seta).click()

                # Clique em exportar
                navegador.find_element("xpath", '//*[@id="divExport"]/ul/li/a').click()
                time.sleep(60)

                pasta_destino = tabela.iloc[i, 2]  # pega a mesma linha de referência da variável cliente. Porém com número de coluna travado

                # Caminho da pasta de downloads
                downloads_path = os.path.expanduser("~") + "/Downloads"

                # Capturando último arquivo baixado
                arquivos_downloads = os.listdir(downloads_path)
                arquivos_csv = [arquivo for arquivo in arquivos_downloads if arquivo.endswith(".csv")]
                arquivos_ordenados = sorted(arquivos_csv,key=lambda x: os.path.getmtime(os.path.join(downloads_path, x)),reverse=True)
                primeiro_arquivo = arquivos_ordenados[0]

                # Caminho completo do arquivo baixado
                caminho_arquivo = os.path.join(downloads_path, primeiro_arquivo)

                # Renomeando o arquivo movido
                nome_arquivo = tabela.iloc[i, 3]  # Obtém o nome do arquivo da coluna correspondente
                parte1 = nome_arquivo[0:9] + "a.csv"
                novo_caminho_arquivo = os.path.join(pasta_destino, parte1)
                shutil.move(caminho_arquivo, novo_caminho_arquivo)

                limpar_data1 = navegador.find_element("xpath", '//*[@id="startDate"]')
                limpar_data1.clear()

                limpar_data2 = navegador.find_element("xpath", '//*[@id="endDate"]')
                limpar_data2.clear()

                print(f"arquivo {parte1} do cliente {cliente} movido com sucesso")


                # REINICIANDO PARA BAIXAR A OUTRA METADE DO MÊS

                navegador.find_element("xpath", campo_cliente).send_keys(cliente)
                navegador.find_element("xpath", campo_cliente).click()
                navegador.find_element("xpath", campo_cliente).send_keys(cliente)

                # Data inicial
                navegador.find_element("xpath", '//*[@id="startDate"]').click()
                navegador.find_element("xpath", '//*[@id="startDate"]').send_keys(str(dia_16))

                data_final = '//*[@id="endDate"]'
                navegador.find_element("xpath", data_final).click()
                navegador.find_element("xpath", data_final).send_keys(str(dia_30))
                navegador.find_element("xpath", data_final).send_keys(Keys.ENTER)
                time.sleep(3)

                seta = '//a[contains(@class, "btn") and contains(@class, "btn-primary") and contains(@class, "pull-right") and contains(@class, "dropdown-toggle") and @data-toggle="dropdown"]'
                navegador.find_element("xpath", seta).click()

                # Clique em exportar
                navegador.find_element("xpath", '//*[@id="divExport"]/ul/li/a').click()
                time.sleep(60)

                pasta_destino = tabela.iloc[i, 2]  # pega a mesma linha de referência da variável cliente. Porém com número de coluna travado

                # Caminho da pasta de downloads
                downloads_path = os.path.expanduser("~") + "/Downloads"

                # Capturando último arquivo baixado
                arquivos_downloads = os.listdir(downloads_path)
                arquivos_csv = [arquivo for arquivo in arquivos_downloads if arquivo.endswith(".csv")]
                arquivos_ordenados = sorted(arquivos_csv,key=lambda x: os.path.getmtime(os.path.join(downloads_path, x)),reverse=True)
                primeiro_arquivo = arquivos_ordenados[0]

                # Caminho completo do arquivo baixado
                caminho_arquivo = os.path.join(downloads_path, primeiro_arquivo)

                # Renomeando o arquivo movido
                nome_arquivo = tabela.iloc[i, 3]  # Obtém o nome do arquivo da coluna correspondente
                parte2 = nome_arquivo[0:9] + "b.csv"
                novo_caminho_arquivo = os.path.join(pasta_destino, parte2)
                shutil.move(caminho_arquivo, novo_caminho_arquivo)


                limpar_data1 = navegador.find_element("xpath", '//*[@id="startDate"]')
                limpar_data1.clear()

                limpar_data2 = navegador.find_element("xpath", '//*[@id="endDate"]')
                limpar_data2.clear()

                tabela.at[i, 'Status'] = 'OK'
                tabela.to_excel(caminho + r'\Base_Bot.xlsx', index=False)      

                verificacao = tabela[tabela['Status'] != 'OK'].index

                if not verificacao.empty:
                    i = verificacao[0]
                else:
                    exit()

                print(f"arquivo {parte2} do cliente {cliente} movido com sucesso")


        except FileNotFoundError as file_not_found_err:
            print(f'Erro ao encontrar arquivo: {file_not_found_err}')
            input('Pressione ENTER para continuar...')
            limpar_data1 = navegador.find_element("xpath", '//*[@id="startDate"]')
            limpar_data1.clear()
            limpar_data2 = navegador.find_element("xpath", '//*[@id="endDate"]')
            limpar_data2.clear()
            continue


        except shutil.Error as shutil_err:

            print(f'Erro ao mover arquivo: {shutil_err}')
            input('Pressione ENTER para continuar...')
            limpar_data1 = navegador.find_element("xpath", '//*[@id="startDate"]')
            limpar_data1.clear()
            limpar_data2 = navegador.find_element("xpath", '//*[@id="endDate"]')
            limpar_data2.clear()
            continue 


        except Exception as e:
            print(f'Ocorreu um erro no cliente {cliente}')
            input('Deseja continuar? Se sim, pressione ENTER')

            limpar_data1 = navegador.find_element("xpath", '//*[@id="startDate"]')
            limpar_data1.clear()

            limpar_data2 = navegador.find_element("xpath", '//*[@id="endDate"]')
            limpar_data2.clear()
            continue
            
else:
    print("Todas os Reports foram baixados. O BOT será encerrado.")
    exit()
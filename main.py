### INCIANDO O SCRIPT ###
import traceback
import pandas as pd
import csv
import re
import zipfile
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import urllib.request
import whisper
import time


try:
    list_ = []
    dataframes_ = []
    url = 'https://cbo.mte.gov.br/cbosite/pages/downloads.jsf'

    # Iniciando o navegador
    chrome_options = Options()
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)
    time.sleep(3)

    # Acessando pagina de download do cbo
    xpath = '/html/body/div[1]/div[3]/div/div[2]/form/span/p[1]/a'
    element = driver.find_element(By.XPATH, xpath).click()
    time.sleep(3)

    # clicando no checkbox do reCaptcha
    frames = driver.find_element(By.TAG_NAME, 'iframe')
    driver.switch_to.frame(frames)
    driver.find_element(By.CLASS_NAME, 'recaptcha-checkbox-border').click()
    time.sleep(3)

    # clicando no botão de audio
    xpath = '/html/body/div[3]/div[4]'
    driver.switch_to.default_content()
    frames = driver.find_element(
        By.XPATH, xpath).find_element(By.TAG_NAME, 'iframe')
    driver.switch_to.frame(frames)
    time.sleep(3)

    driver.find_element(By.ID, 'recaptcha-audio-button').click()
    time.sleep(3)

    # reproduzindo o audio
    driver.switch_to.default_content()
    frames = driver.find_element(
        By.XPATH, xpath).find_element(By.TAG_NAME, 'iframe')
    driver.switch_to.frame(frames)
    time.sleep(3)

    xpath = '/html/body/div/div/div[3]/div/button'
    element = driver.find_element(By.XPATH, xpath).click()
    time.sleep(3)

    # capturando o link de download do audio
    src = driver.find_element(By.ID, 'audio-source').get_attribute('src')
    print(f"[INFO] Audio src: {src}")

    # Caminho para a pasta de downloads
    caminho_origem = os.path.expanduser('~/Downloads/')

    # baixando o audio
    urllib.request.urlretrieve(src, caminho_origem+'audio.mp3')
    print(f'baixou o com sucesso.')

    # Convertendo audio para texto com IA
    original_file_path = os.path.join(caminho_origem, 'audio.mp3')
    model = whisper.load_model("base")
    time.sleep(2)
    result = model.transcribe(original_file_path, fp16=False)
    key = result["text"]
    print(f'password code: {key}')

    # escrevendo o texto no input do reCaptcha
    driver.find_element(By.ID, 'audio-response').send_keys(key.lower())
    time.sleep(1)
    driver.find_element(By.ID, 'audio-response').send_keys(Keys.ENTER)
    time.sleep(2)

    # reCaptcha burlado, baixando o .zip do CBO
    driver.switch_to.default_content()
    xpath = '/html/body/div[1]/div[3]/div/div[2]/form/span/div[2]/input[1]'
    element = driver.find_element(By.XPATH, xpath).click()
    time.sleep(5)
    driver.close()

    # Listar todos os arquivos na pasta Downloads
    conteudo_pasta = os.listdir(caminho_origem)

    # Filtrar apenas os arquivos .zip
    arquivos_zip = [
        arquivo for arquivo in conteudo_pasta if arquivo.endswith(".ZIP")
    ]

    if arquivos_zip:
        nome_arquivo_zip = arquivos_zip[0]
        caminho_arquivo_zip = os.path.join(caminho_origem, nome_arquivo_zip)

        # Pasta de destino para a extração (substitua pelo caminho desejado)
        caminho_destino = os.path.expanduser("~/Downloads/Extracao/")

        # Criar a pasta de destino, se ainda não existir
        os.makedirs(caminho_destino, exist_ok=True)

        # Extrair o arquivo .zip para a pasta de destino
        with zipfile.ZipFile(caminho_arquivo_zip, 'r') as zip_ref:
            zip_ref.extractall(caminho_destino)

        print(
            f"Arquivo '{nome_arquivo_zip}' extraído para '{caminho_destino}'.")
    else:
        print("Nenhum arquivo .zip encontrado na pasta de origem.")

    # listar os arquivos extraidos
    csv_files = os.listdir(caminho_destino)

    # Dicionario com os arquivos csv
    data = {
        'cbo2002Familia': caminho_destino + csv_files[0],
        'cbo2002GrandeGrupo': caminho_destino + csv_files[1],
        'cbo2002Ocupacao': caminho_destino + csv_files[2],
        'cbo2002Sinonimo': caminho_destino + csv_files[4],
        'cbo2002SubGrupoPrincipal': caminho_destino + csv_files[5],
        'cbo2002SubGrupo': caminho_destino + csv_files[6]
    }

    # arquivo csv cbo2002PerfilOcupacional para tratamento de dados
    df_cbo2002PerfilOcupacional = caminho_destino+csv_files[3]

    # abrindo o csv, criando os indices
    with open(df_cbo2002PerfilOcupacional, 'r') as file:
        reader = csv.reader(file)
        for i, row in enumerate(reader):
            list_.append((i, row))

    # percorrer cada indice e remover os ; por / dentro de parenteses
    for i, row in list_:
        for j in range(len(row)):
            # Usando expressão regular para encontrar o padrão "(...)"
            match = re.search(r'\((.*?)\)', row[j])
            if match:
                # Substituindo os pontos e vírgulas por barras apenas dentro dos parênteses
                new_value = re.sub(r';', '/', match.group(1))
                # Atualizando o valor na lista original
                list_[i][1][j] = re.sub(r'\((.*?)\)', f'({new_value})', row[j])

    # percorrendo a nova lista e removendo a linha com "coleta(bags;" pois ele não trocou ; por /
    for item in list_:
        if any("coleta(bags;" in value for value in item[1]):
            list_.remove(item)
            new_index = len(list_)
            new_item = [
                '5;51;519;5192;519205;A;COLETAR MATERIAL RECICLÁVEL E REAPROVEITÁVEL;7;Fornecer recipientes para coleta de bags, conteineres, etc.']
            list_.append((new_index, new_item))

    # Removendo o índice da lista
    df_list = [item[1] for item in list_]
    df = pd.DataFrame(df_list)

    # Juntar todas as colunas em uma única coluna
    df['combined'] = df.apply(lambda row: ' '.join(map(str, row)), axis=1)

    # Dropar as colunas de 0 até 9
    df = df.drop(df.columns[0:10], axis=1)

    # Substituir valores 'None' por espaços em branco na coluna 'combined'
    df['combined'] = df['combined'].replace(r'\bNone\b', ' ', regex=True)

    # Remover espaços em branco extras no final da string
    df['combined'] = df['combined'].str.strip()

    # Separar os dados por ponto e vírgula (;) e criar colunas separadas
    df = df['combined'].str.split(';', expand=True)

    # Definir a primeira coluna como o título
    df.columns = df.iloc[0]

    # Remover a primeira linha (título original)
    cbo2002PerfilOcupacional = df[1:]

    # Adicionar o dataframe ao dicionário dataframes_
    dataframes_['cbo2002PerfilOcupacional'] = cbo2002PerfilOcupacional

    # percorrendo cada csv, lendo como dataframe e adicionando na lista
    for key, value in data.items():
        df_ = pd.read_csv(value, encoding='cp1252', delimiter=';')
        dataframes_[key] = df_

    dataframes_
except Exception as e:
    traceback_str = traceback.format_exc()
    print(f"Ocorreu um erro:\n{traceback_str}, {e}")

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests

# Abrir a planilha com os dados de CEPs
file_path = r"C:\Users\BR05307045\Escritorio\kms_ceps.xlsx"
df = pd.read_excel(file_path, header=None)  # Lê sem considerar cabeçalhos, já que não tem nomes de colunas

# Verificar se as colunas 0 (cep_inicio) e 1 (cep_fim) existem
if df.shape[1] < 2:
    raise KeyError("A planilha precisa ter pelo menos duas colunas com os dados de 'cep_inicio' e 'cep_fim'!")

# Configuração do Microsoft Edge
options = Options()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-web-security")
options.add_argument("--disable-features=IsolateOrigins,site-per-process")

# Usando o webdriver_manager para garantir o EdgeDriver correto
servico = Service()
navegador = webdriver.Edge(service=servico, options=options)

# Função para buscar a distância entre dois CEPs no Google Maps
def buscar_distancia(cep_inicio, cep_fim):
    try:
        # A URL para buscar a direção no Google Maps
        url = f"https://www.google.com/maps/dir/{cep_inicio}/{cep_fim}"
        navegador.get(url)
        
        # Esperando o carregamento da página e do elemento de distância
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ivN21e"))
        )

        # Obtendo o conteúdo da página
        soup = BeautifulSoup(navegador.page_source, 'html.parser')

        # Tentando extrair a distância
        distance_div = soup.find('div', class_='ivN21e')
        if distance_div:
            distance = distance_div.get_text(strip=True)
            return distance, url
        else:
            return "Não encontrado", url
    except Exception as e:
        return f"Erro: {str(e)}", f"https://www.google.com/maps/dir/{cep_inicio}/{cep_fim}"

# Função para pegar o link do thumbnail do mapa
def get_thumbnail(url):
    try:
        # Definir cabeçalhos HTTP (User-Agent)
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        # Realiza a requisição HTTP para pegar o conteúdo da página com os cabeçalhos
        response = requests.get(url, headers=headers, verify=False)
        response.raise_for_status()  # Verifica se a requisição foi bem-sucedida

        # Faz o parsing do HTML da página
        soup = BeautifulSoup(response.content, 'html.parser')

        # Tenta encontrar a imagem de thumbnail com Open Graph (og:image)
        og_image = soup.find("meta", property="og:image")
        if og_image:
            return og_image["content"]

        # Se não encontrar, tenta buscar pelo link de ícone (favicon)
        icon_link = soup.find("link", rel="icon")
        if icon_link:
            return icon_link["href"]

        # Caso não encontre nada, retorna None
        return None
    except Exception as e:
        print(f"Erro ao acessar a página para pegar o thumbnail: {e}")
        return None

# Iterar sobre cada linha da planilha
for index, row in df.iterrows():
    cep_inicio = row[0]  # Coluna A (0-indexed)
    cep_fim = row[1]     # Coluna B (0-indexed)
    
    # Buscar a distância e o link para cada par de CEPs
    km, link = buscar_distancia(cep_inicio, cep_fim)
    
    # Obter o link do thumbnail
    thumbnail = get_thumbnail(link)
    
    # Atualizar a coluna 'kms' (coluna C), 'link' (coluna D) e 'thumbnail' (coluna E) na planilha
    df.at[index, 2] = km  # Coluna C (2-indexed)
    df.at[index, 3] = link  # Coluna D (3-indexed)
    df.at[index, 4] = thumbnail if thumbnail else "Não encontrado"  # Coluna E (4-indexed)
    
    print(f"Distância entre {cep_inicio} e {cep_fim}: {km}")
    print(f"Link para o Google Maps: {link}")
    print(f"Thumbnail: {thumbnail}")

# Salvar a planilha com os valores atualizados
output_path = r"C:\Users\BR05307045\Escritorio\kms_ceps_atualizado.xlsx"
df.to_excel(output_path, index=False, header=False)  # Não incluir cabeçalho na saída

# Fechar o navegador após a execução
navegador.quit()

print(f"Planilha salva em: {output_path}")

# Solução para Buscar Distâncias e Thumbnails de Mapas entre CEPs

## Descrição

Este repositório contém uma solução em Python para buscar a distância entre dois CEPs no Google Maps e obter a imagem de thumbnail do trajeto, salvando os dados em uma planilha Excel. O código utiliza o Selenium para navegar no Google Maps e o BeautifulSoup para extrair informações do HTML da página. Além disso, o link do Google Maps e a imagem de thumbnail também são extraídos e salvos na planilha.

## Funcionalidades

- **Busca de Distância**: Utiliza o Google Maps para calcular a distância entre dois CEPs fornecidos.
  
- **Link para o Google Maps**: Gera e salva o link direto do Google Maps para o trajeto entre os dois CEPs.
  
- **Thumbnail**: Extrai o link da imagem de thumbnail do mapa utilizando o protocolo Open Graph (`og:image`).
  
- **Saída em Excel**: Os resultados são salvos em uma planilha Excel, contendo o CEP de origem, o CEP de destino, a distância, o link para o Google Maps e o link da thumbnail do mapa.

## Tecnologias Utilizadas

- **Python**: A linguagem principal utilizada para implementar a solução.
- **Selenium**: Para automação do navegador e interações com o Google Maps.
- **BeautifulSoup**: Para parseamento e extração de dados da página HTML.
- **Pandas**: Para manipulação e armazenamento dos dados em formato Excel.
- **Requests**: Para fazer requisições HTTP e obter o conteúdo das páginas do Google Maps para extrair a thumbnail.

## Como Funciona

1. **Leitura da Planilha**  
   A solução começa lendo uma planilha Excel contendo duas colunas: uma com os CEPs de origem e outra com os CEPs de destino.

2. **Navegação no Google Maps**  
   Com os CEPs fornecidos, a solução utiliza o Selenium para abrir o Google Maps e calcular a distância entre os dois pontos.

3. **Extração da Distância e Link do Mapa**  
   A distância entre os dois CEPs é extraída da página do Google Maps, e o link do trajeto no Google Maps é gerado.

4. **Extração da Thumbnail do Mapa**  
   A imagem de thumbnail do mapa é extraída usando o BeautifulSoup, que pega a URL da imagem de thumbnail presente na página do Google Maps. Caso a thumbnail não seja encontrada, o código retorna "Não encontrado".

5. **Armazenamento dos Resultados**  
   Os resultados são armazenados em uma planilha Excel com 5 colunas:
   - **CEP Início**
   - **CEP Fim**
   - **Distância (km)**
   - **Link do Google Maps**
   - **Thumbnail do Mapa**

6. **Salvamento da Planilha**  
   A planilha é salva com os resultados em um novo arquivo Excel.

## Como Usar

### Pré-requisitos

Instalar as dependências:

Você precisará do Python 3 instalado em sua máquina.

### Instale as dependências necessárias utilizando pip:

```bash
pip install selenium pandas requests beautifulsoup4 openpyxl

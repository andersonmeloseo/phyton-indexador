# phyton-indexador
script phyton para indexar urls
# Google Search Console Indexing Script

Este script faz uma análise de indexação de URLs usando a API do Google Search Console e a API de Indexação do Google, além de enviar URLs para a fila de indexação.

## Configuração

1. Crie um projeto no Google Cloud e habilite as APIs:
   - Google Search Console API
   - Google Indexing API
2. Baixe o arquivo JSON de credenciais da conta de serviço.
3. Renomeie o arquivo para `personal.json` e mova-o para o diretório do script.

## Instalação

Instale as bibliotecas necessárias:

```bash
pip install oauth2client google-api-python-client openpyxl requests
===

1. Preparação Inicial
Antes de começar a rodar o script, você precisará configurar algumas bibliotecas e credenciais.

a. Criar o projeto no Google Cloud e configurar as APIs:
Acesse o Google Cloud Console.
Crie um projeto.
Habilite as APIs:
Google Search Console API
Google Indexing API
Vá em "Credenciais" e crie uma conta de serviço.
Baixe o arquivo JSON de credenciais da conta de serviço. Nomeie-o como personal.json e mova-o para o diretório onde seu script Python estará localizado.

b. Instalar as bibliotecas necessárias:
Use o pip para instalar as bibliotecas requeridas:

bash
Copiar código
pip install oauth2client google-api-python-client openpyxl requests
2. Criação do Script Python
Salve o código abaixo em um arquivo Python, por exemplo, indexacao_google_search_console.py.

python
Copiar código
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import openpyxl
from openpyxl import Workbook
import requests
import xml.etree.ElementTree as ET
import datetime
import re

# Configuração inicial
SCOPES = ["https://www.googleapis.com/auth/indexing", "https://www.googleapis.com/auth/webmasters"]
JSON_KEY_FILE = "personal.json"

# Solicitar a URL do domínio a ser analisado
def validate_domain_input(domain):
    if not domain.startswith("https://") and not domain.startswith("http://"):
        domain = "https://" + domain
    domain = domain.replace("https//", "https://").replace("http//", "http://")
    domain = domain.rstrip("/")
    return domain

DOMAIN = input("Digite a URL do domínio que deseja analisar (exemplo: https://cbmadvs.com.br/): ").strip()
DOMAIN = validate_domain_input(DOMAIN)

# Autenticação com as credenciais
credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scopes=SCOPES)
search_service = build('searchconsole', 'v1', credentials=credentials)
indexing_service = build('indexing', 'v3', credentials=credentials)

# Funções (get_urls_from_sitemap, get_all_urls_from_sitemaps, get_index_status, send_url_to_indexing, get_performance_comparison, create_excel_report)
# (Inclua as funções que já estão no seu script)
Certifique-se de que o arquivo personal.json está no mesmo diretório que o script.

3. Rodar o Script
Abra o terminal, navegue até o diretório onde está o script, e execute o comando:

bash
Copiar código

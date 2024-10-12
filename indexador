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
    # Adiciona 'https://' se estiver faltando
    if not domain.startswith("https://") and not domain.startswith("http://"):
        domain = "https://" + domain
    # Corrigir 'https//' para 'https://'
    domain = domain.replace("https//", "https://").replace("http//", "http://")
    # Remove barras adicionais no final da URL
    domain = domain.rstrip("/")
    return domain

DOMAIN = input("Digite a URL do domínio que deseja analisar (exemplo: https://cbmadvs.com.br/): ").strip()
DOMAIN = validate_domain_input(DOMAIN)

# Autenticação com as credenciais
credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scopes=SCOPES)
search_service = build('searchconsole', 'v1', credentials=credentials)
indexing_service = build('indexing', 'v3', credentials=credentials)

# Função para obter URLs do sitemap
def get_urls_from_sitemap(sitemap_url):
    try:
        response = requests.get(sitemap_url)
        response.raise_for_status()
        urls = []
        root = ET.fromstring(response.content)

        # Extrair URLs do sitemap
        for url in root.findall(".//{http://www.sitemaps.org/schemas/sitemap/0.9}loc"):
            urls.append(url.text)
        
        print(f"Encontradas {len(urls)} URLs no sitemap: {sitemap_url}")
        return urls
    except requests.exceptions.RequestException as e:
        print(f"Erro ao acessar o sitemap {sitemap_url}: {e}")
        return []

# Função para obter todas as URLs dos sitemaps listados no Search Console
def get_all_urls_from_sitemaps(service, site_url):
    sitemap_list = service.sitemaps().list(siteUrl=site_url).execute()
    all_urls = []
    for sitemap in sitemap_list.get('sitemap', []):
        sitemap_url = sitemap['path']
        print(f"Processando sitemap: {sitemap_url}")
        all_urls.extend(get_urls_from_sitemap(sitemap_url))
    return list(set(all_urls))  # Remover duplicatas

# Função para obter o status de indexação e dados de desempenho das URLs com paginação
def get_index_status(service, site_url):
    all_rows = []
    start_row = 0
    page_size = 1000

    while True:
        request = service.searchanalytics().query(
            siteUrl=site_url,
            body={
                "startDate": "2023-01-01",
                "endDate": datetime.date.today().strftime('%Y-%m-%d'),
                "dimensions": ["page"],
                "startRow": start_row,
                "rowLimit": page_size
            }
        )
        response = request.execute()
        rows = response.get('rows', [])
        if not rows:
            break

        all_rows.extend(rows)
        start_row += len(rows)

        print(f"Obtidos {len(rows)} URLs. Total até agora: {len(all_rows)} URLs.")

        if len(rows) < page_size:
            break

    return all_rows

# Função para enviar uma URL para a fila de indexação com tratamento de erros
def send_url_to_indexing(service, url):
    body = {
        "url": url,
        "type": "URL_UPDATED"
    }
    try:
        response = service.urlNotifications().publish(body=body).execute()
        return response, "Enviado para indexação"
    except Exception as e:
        print(f"Erro ao enviar a URL: {url}")
        print(f"Detalhes do erro: {e}")
        return str(e), "Erro ao enviar para indexação"

# Função para obter o desempenho dos últimos 30 dias e comparar com os 30 dias anteriores
def get_performance_comparison(service, site_url):
    end_date = datetime.date.today()
    start_date_30_days = end_date - datetime.timedelta(days=30)
    start_date_60_days = start_date_30_days - datetime.timedelta(days=30)

    # Dados dos últimos 30 dias
    request_30_days = service.searchanalytics().query(
        siteUrl=site_url,
        body={
            "startDate": start_date_30_days.strftime('%Y-%m-%d'),
            "endDate": end_date.strftime('%Y-%m-%d'),
            "dimensions": ["page"]
        }
    ).execute()

    # Dados dos 30 dias anteriores
    request_30_days_before = service.searchanalytics().query(
        siteUrl=site_url,
        body={
            "startDate": start_date_60_days.strftime('%Y-%m-%d'),
            "endDate": start_date_30_days.strftime('%Y-%m-%d'),
            "dimensions": ["page"]
        }
    ).execute()

    return request_30_days.get('rows', []), request_30_days_before.get('rows', [])

# Função para gerar relatório em Excel
def create_excel_report(rows, all_urls, indexed_urls, non_indexed_urls, sent_for_indexing, performance_30_days, performance_30_days_before):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório de Indexação"
    ws.append(["URL", "Status", "Ação", "Impressões"])

    # Adiciona todas as URLs com seus status e impressões
    for url, impressions in indexed_urls.items():
        ws.append([url, "Indexada", "", impressions])

    for url in non_indexed_urls:
        ws.append([url, "Rastreadas, mas não indexadas", "", 0])

    # Adiciona URLs enviadas para indexação
    for url, status in sent_for_indexing:
        ws.append([url, "Enviada para indexação", status, 0])

    # Aba de desempenho das URLs
    ws_performance = wb.create_sheet("Desempenho 12 meses")
    ws_performance.append(["URL", "Cliques", "Impressões", "CTR", "Posição Média"])

    for row in rows:
        page = row['keys'][0]
        clicks = row.get('clicks', 0)
        impressions = row.get('impressions', 0)
        ctr = row.get('ctr', 0)
        position = row.get('position', 0)
        ws_performance.append([page, clicks, impressions, ctr, position])

    # Aba de desempenho dos últimos 30 dias
    ws_performance_30_days = wb.create_sheet("Desempenho 30 dias")
    ws_performance_30_days.append([
        "URL", "Cliques (Últimos 30 dias)", "Impressões (Últimos 30 dias)", 
        "Posição Média (Últimos 30 dias)", "Cliques (30 dias anteriores)", 
        "Impressões (30 dias anteriores)", "Posição Média (30 dias anteriores)", 
        "Diferença Cliques (%)", "Diferença Impressões (%)", "Diferença Posição (%)"
    ])

    # Processa os dados dos últimos 30 dias e os 30 dias anteriores
    clicks_30_days = {row['keys'][0]: row.get('clicks', 0) for row in performance_30_days}
    impressions_30_days = {row['keys'][0]: row.get('impressions', 0) for row in performance_30_days}
    position_30_days = {row['keys'][0]: row.get('position', 0) for row in performance_30_days}

    clicks_30_days_before = {row['keys'][0]: row.get('clicks', 0) for row in performance_30_days_before}
    impressions_30_days_before = {row['keys'][0]: row.get('impressions', 0) for row in performance_30_days_before}
    position_30_days_before = {row['keys'][0]: row.get('position', 0) for row in performance_30_days_before}

    # Calcula as diferenças percentuais e adiciona ao relatório
    for url in set(clicks_30_days.keys()).union(clicks_30_days_before.keys()):
        clicks_now = clicks_30_days.get(url, 0)
        impressions_now = impressions_30_days.get(url, 0)
        position_now = position_30_days.get(url, 0)

        clicks_before = clicks_30_days_before.get(url, 0)
        impressions_before = impressions_30_days_before.get(url, 0)
        position_before = position_30_days_before.get(url, 0)

        change_clicks = ((clicks_now - clicks_before) / clicks_before * 100) if clicks_before > 0 else 100 if clicks_now > 0 else 0
        change_impressions = ((impressions_now - impressions_before) / impressions_before * 100) if impressions_before > 0 else 100 if impressions_now > 0 else 0
        change_position = ((position_now - position_before) / position_before * 100) if position_before > 0 else 0

        ws_performance_30_days.append([
            url, clicks_now, impressions_now, position_now,
            clicks_before, impressions_before, position_before,
            change_clicks, change_impressions, change_position
        ])

    # Aba com URLs do sitemap, adicionando "site:" antes de cada URL
    ws_sitemap = wb.create_sheet("URLs do Sitemap")
    ws_sitemap.append(["URL", "Status"])

    for url in all_urls:
        status = "Indexada" if url in indexed_urls else "Rastreadas, mas não indexadas"
        ws_sitemap.append([f"site:{url}", status])

    # Resumo
    ws_summary = wb.create_sheet("Resumo")
    total_urls = len(all_urls)
    indexed_count = len(indexed_urls)
    non_indexed_count = len(non_indexed_urls)
    sent_count = len(sent_for_indexing)

    ws_summary.append(["Total de URLs no Sitemap", total_urls])
    ws_summary.append(["URLs Indexadas", indexed_count])
    ws_summary.append(["URLs Rastreadas mas Não Indexadas", non_indexed_count])
    ws_summary.append(["URLs Enviadas para Indexação", sent_count])

    # Salva o arquivo Excel com data e hora
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    domain_name = DOMAIN.replace("https://", "").replace("http://", "").replace("/", "")
    filename = f"relatorio_indexacao_{domain_name}_{now}.xlsx"
    wb.save(filename)
    print(f"Relatório gerado: {filename}")

# Obter todas as URLs dos sitemaps para análise
all_urls = get_all_urls_from_sitemaps(search_service, DOMAIN)

# Obter status de indexação das URLs com paginação
rows = get_index_status(search_service, DOMAIN)

# Obter dados de desempenho dos últimos 30 dias e dos 30 dias anteriores
performance_30_days, performance_30_days_before = get_performance_comparison(search_service, DOMAIN)

# Separar URLs indexadas e não indexadas, incluindo URLs com 0 impressões
indexed_urls = {row['keys'][0]: row.get('impressions', 0) for row in rows}
non_indexed_urls = set(all_urls) - indexed_urls.keys()

# Ajuste para limitar o número de envios por dia
MAX_REQUESTS_PER_DAY = 200

# Enviar URLs não indexadas para a fila de indexação, respeitando o limite de solicitações diárias
sent_for_indexing = []
request_count = 0

for url in non_indexed_urls:
    if request_count >= MAX_REQUESTS_PER_DAY:
        print(f"Limite de {MAX_REQUESTS_PER_DAY} envios diários atingido. Tente novamente amanhã.")
        break

    response, status = send_url_to_indexing(indexing_service, url)
    sent_for_indexing.append((url, status))
    print(f"URL: {url}, Status: {status}")
    request_count += 1

# Gerar o relatório final
create_excel_report(rows, all_urls, indexed_urls, non_indexed_urls, sent_for_indexing, performance_30_days, performance_30_days_before)

import os
import io
import pickle
import pandas as pd
from openai import OpenAI
from openai import cli

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from PyPDF2 import PdfReader

# -------------------------
# CONFIGURAÇÕES INICIAIS
# -------------------------
# Escopo para acessar o Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
# Pasta local para salvar os PDFs baixados
DOWNLOAD_FOLDER = "pdf_downloads"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# Caminho para a planilha de desafios
CAMINHO_DESAFIOS = "trilha_ia_desafios.xlsx"

# Definindo a variável de ambiente OPENAI_API_KEY
os.environ["OPENAI_API_KEY"]  # Substitua pelo seu token OpenAI

# Inicializar o cliente OpenAI
openai_client = OpenAI()

# -------------------------
# FUNÇÕES DE AUTENTICAÇÃO E DOWNLOAD
# -------------------------
def authenticate_drive():
    """Autentica e retorna o serviço do Google Drive."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # Se não houver credenciais válidas, realize o login.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Salve as credenciais para a próxima execução.
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    return service


def list_pdf_files(service, folder_id):
    """Lista arquivos PDF dentro de uma pasta do Google Drive."""
    query = f"'{folder_id}' in parents and mimeType='application/pdf'"
    results = service.files().list(q=query, fields="files(id, name, createdTime)").execute()
    files = results.get('files', [])
    return files


def download_file(service, file_id, file_name, destination_folder):
    """Faz o download do arquivo a partir do Google Drive."""
    request = service.files().get_media(fileId=file_id)
    file_path = os.path.join(destination_folder, file_name)
    with io.FileIO(file_path, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}% concluído: {file_name}")
    return file_path

# -------------------------
# LEITURA DO DESAFIO ATIVO
# -------------------------
def get_active_desafio(planilha_path=CAMINHO_DESAFIOS):
    """
    Lê a planilha de desafios e retorna os dados do desafio ativo.
    A planilha deve conter as colunas: Etapa, Semana, Desafio, Critérios, Atual.
    Retorna um dicionário com as informações do desafio cuja coluna "Atual" seja "X".
    """
    try:
        df = pd.read_excel(planilha_path)
        ativo = df[df["Atual"] == "x"]
        if ativo.empty:
            raise ValueError("Nenhum desafio ativo encontrado na planilha.")
        # Se houver mais de um, pega o primeiro
        row = ativo.iloc[0]
        return row.to_dict()
    except Exception as e:
        print(f"Erro ao ler a planilha de desafios: {e}")
        return None

# -------------------------
# AVALIAÇÃO DO TRABALHO COM BASE NO DESAFIO
# -------------------------
def avaliar_trabalho(texto_trabalho, desafio_data):
    """
    Utiliza a API da OpenAI para avaliar o trabalho com base no desafio ativo.
    desafio_data é um dicionário com as colunas: Etapa, Semana, Desafio, Critérios.
    Retorna: "Bom", "Regular" ou "Ruim".
    """
    prompt = f"""
Você é um avaliador de trabalhos. Com base no desafio a seguir, avalie o trabalho de um aluno.

Etapa: {desafio_data.get('Etapa', 'N/A')}
Semana: {desafio_data.get('Semana', 'N/A')}
Desafio: {desafio_data.get('Desafio', 'N/A')}

Critérios:
{desafio_data.get('Critérios', 'N/A')}

Trabalho:
{texto_trabalho}

Avalie apenas com uma das palavras: Bom, Regular, Ruim.  Extraia o nome do aluno e o número do "RA" a partir da seguinte estrutura encontrada no pdf:

RA do Aluno/RA: 99999
Nome/ALUNO: xxxxxxxxxxxxxxxxxxxxxxxxxx

O retorno deve ser no formato de um dicionário com no exemplo a seguir: Ex: RA: "999999", ALUNO: "XXXXXXXXXXXXXXX", RESULTADO: "bom"     
    """
    try:
        response = openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=50,
            temperature=0.0
        )
        # Acessa o conteúdo da resposta
        #resultado = response.choices[0].message.content.strip()
        resultado = response.choices[0].message.to_dict()
        print("NOME ALUNO:", resultado)
        # if resultado not in ["Bom", "Regular", "Ruim"]:
        #     resultado = "Outro"
        return resultado
    except Exception as e:
        print(f"Erro na avaliação via OpenAI: {e}")
        return "Regular"


# -------------------------
# EXTRAÇÃO DO TEXTO DO PDF
# -------------------------
def extract_text_from_pdf(pdf_path):
    """Extrai o texto de um arquivo PDF."""
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        return text
    except Exception as e:
        print(f"Erro ao extrair texto de {pdf_path}: {e}")
        return ""


# -------------------------
# GERAÇÃO DA PLANILHA
# -------------------------
def gerar_planilha(resultados, nome_arquivo="avaliacoes.xlsx"):
    """
    Gera uma planilha Excel com os resultados da avaliação.
    resultados: lista de dicionários com as chaves: 'aluno', 'arquivo', 'avaliacao'
    """
    df = pd.DataFrame(resultados)
    print(df)
    df.to_excel(nome_arquivo, index=False)
    print(f"Planilha gerada: {nome_arquivo}")


# -------------------------
# FUNÇÃO PRINCIPAL
# -------------------------
def main():
    # Obtém o desafio ativo da planilha de desafios
    desafio_atual = get_active_desafio()
    if not desafio_atual:
        print("Não foi possível encontrar um desafio ativo. Verifique a planilha de desafios.")
        return

    print("Desafio ativo:")
    print(f"Etapa: {desafio_atual.get('Etapa')}, Semana: {desafio_atual.get('Semana')}")
    print(f"Desafio: {desafio_atual.get('Desafio')}")
    print(f"Critérios: {desafio_atual.get('Critérios')}\n")
    DRIVE_FOLDER_ID = desafio_atual.get('Directory_Key')  # ID da pasta do Google Drive que contém os PDFs

    service = authenticate_drive()
    arquivos = list_pdf_files(service, DRIVE_FOLDER_ID)
    print(f"Encontrados {len(arquivos)} arquivos PDF.")

    resultados_avaliacoes = []

    for arquivo in arquivos:
        file_id = arquivo['id']
        file_name = arquivo['name']
        print(f"Processando {file_name}...")

        pdf_path = download_file(service, file_id, file_name, DOWNLOAD_FOLDER)
        texto = extract_text_from_pdf(pdf_path)
        if not texto:
            print(f"Não foi possível extrair texto de {file_name}.")
            continue

        avaliacao = avaliar_trabalho(texto, desafio_atual)
        # Supondo que o nome do aluno esteja no nome do arquivo (sem extensão)
        aluno = os.path.splitext(file_name)[0]

        resultados_avaliacoes.append({
            "aluno": aluno,
            "arquivo": file_name,
            "avaliacao": avaliacao,
            "etapa": desafio_atual.get("Etapa"),
            "semana": desafio_atual.get("Semana")
        })

    gerar_planilha(resultados_avaliacoes)

if __name__ == "__main__":
    main()

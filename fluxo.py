import os
import shutil
from datetime import datetime
from visualizarAnexo import extrair_texto

PASTA_ENTRADA = r"C:\Users\pedro_moraes\OneDrive - YanmarGlobal\Área de Trabalho\Notas Fiscais" 
PASTA_SAIDA = r"C:\Users\pedro_moraes\OneDrive - YanmarGlobal\Área de Trabalho\Notas Fiscais" 

# Armazena a palavra_alvo no código
with open("palavra_alvo.txt", "r", encoding="utf-8") as f:
    PALAVRA_ALVO = [
        line.strip() for line in f
        if line.strip() and not line.strip().startswith("#")
    ]

def listar_pdfs(pasta):
    pdfs = []
    for arq in os.listdir(pasta):
        p = os.path.join(pasta, arq)
        if os.path.isfile(p) and arq.lower().endswith(".pdf"):
            pdfs.append(p)
    return pdfs


if __name__ == "__main__":
    # Lista PDFs da pasta
    pdfs = listar_pdfs(PASTA_ENTRADA)

    if len(pdfs) == 1:
        print(f"[INFO] Foi encontrado {len(pdfs)} PDF na pasta.")
    else:
        print(f"[INFO] Foram encontrados {len(pdfs)} PDFs na pasta.")
    if not pdfs:
        print("[INFO] Nenhum PDF encontrado.")
        raise SystemExit

    if len(pdfs) == 1:
        print("[INFO] Iniciando OCR do arquivo...")
    else:
        print("[INFO] Iniciando OCR dos arquivos...")
    
    # Extrai o texto de todos os PDFs
    textos_por_pdf = {}
    for pdf_path in pdfs:
        try:
            textos_por_pdf[pdf_path] = extrair_texto(pdf_path)
        except Exception as e:
            print(f"[OCR] Erro ao extrair texto de {pdf_path}: {e}")
            textos_por_pdf[pdf_path] = ""

    # Busca palavras e move
    for pdf_path in pdfs:
        texto = textos_por_pdf.get(pdf_path, "")
        if not texto:
            print(f"[OCR] Sem texto extraído: {pdf_path}")
            continue

        palavra_encontrada = next(
            (w for w in PALAVRA_ALVO if w.lower() in texto.lower()), None
        )

        # Data de hoje
        data_hoje = datetime.now().strftime("%d%m%Y")

        # Nome original com data na frente
        nome_original = os.path.basename(pdf_path)
        novo_nome= f"{data_hoje}_{nome_original}"

        # Define destino
        if palavra_encontrada:
            destino_dir = os.path.join(PASTA_SAIDA, palavra_encontrada)
        else:
            destino_dir = os.path.join(PASTA_SAIDA, "SEM_MATCH")

        os.makedirs(destino_dir, exist_ok=True)

        destino_path = os.path.join(destino_dir, novo_nome)

        # Move os PDFs de acordo com o nome
        try:
            shutil.move(pdf_path, destino_path)
            if palavra_encontrada:
                print(f"[OCR] '{palavra_encontrada}' ENCONTRADA — movido para: {destino_path}")
            else:
                print(f"[OCR] Nenhuma palavra-alvo — movido para: {destino_path}")
        except Exception as e:
            print(f"[OCR] Erro ao mover {pdf_path}: {e}")

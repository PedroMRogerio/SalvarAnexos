import os
import shutil
from datetime import datetime

from baixarAnexo import download_folder, folder, percorrer_pastas
from visualizarAnexo import extrair_texto

# Armazena a palavra_alvo no código
with open("palavra_alvo.txt", "r", encoding="utf-8") as f:
        PALAVRA_ALVO = [
            line.strip() for line in f
            if line.strip() and not line.strip().startswith("#")
        ]

if __name__ == "__main__":
    percorrer_pastas(folder)

    # Lista PDFs
    pdfs = []
    for f in os.listdir(download_folder):
        if f.lower().endswith(".pdf"):
            p = os.path.join(download_folder, f)
            if os.path.isfile(p):
                pdfs.append(p)

    # Extrai o texto de todos os PDFs
    textos_por_pdf = {}
    for pdf_path in pdfs:
        try:
            textos_por_pdf[pdf_path] = extrair_texto(pdf_path)
        except Exception as e:
            print(f"[OCR] Erro ao extrair texto de {pdf_path}: {e}")

    # Busca as palavras
    for pdf_path in pdfs:
        texto = textos_por_pdf.get(pdf_path, "")
        if not texto:
            print(f"[OCR] Sem texto extraído: {pdf_path}")
            continue

        # Encontra a primeira palavra
        palavra_encontrada = next(
            (w for w in PALAVRA_ALVO if w.lower() in texto.lower()), None
        )

        # Gera data de hoje no formato DDMMYYYY
        data_hoje = datetime.now().strftime("%d%m%Y")

        # Novo nome
        novo_nome = f"{data_hoje}.pdf"
        
        # Salva os destino
        if palavra_encontrada:
            destino_dir = os.path.join(download_folder, palavra_encontrada)
        else:
            destino_dir = os.path.join(download_folder, "SEM_MATCH")

        os.makedirs(destino_dir, exist_ok=True)
        destino_path = os.path.join(destino_dir, novo_nome)
        
        # Salva os anexos
        try:
            if os.path.exists(destino_path):
                os.remove(destino_path)  # Sobrescreve
            shutil.move(pdf_path, destino_path)
            if palavra_encontrada:
                print(
                    f"[OCR] '{palavra_encontrada}' ENCONTRADA — movido para: {destino_path}"
                )
            else:
                print(f"[OCR] Nenhuma palavra-alvo — movido para: {destino_path}")
        except Exception as e:
            print(f"[OCR] Erro ao mover {pdf_path}: {e}")

from doctr.io import DocumentFile
from doctr.models import ocr_predictor

# Extrai o texto dos PDFs salvos
def extrair_texto(pdf_path: str) -> str:
    model = ocr_predictor(pretrained=True)
    doc = DocumentFile.from_pdf(pdf_path)
    result = model(doc)
    return result.render() or ""

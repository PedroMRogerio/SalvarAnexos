## 📨 Script para salvar anexos do Outlook 

Um script em Python que automatiza:

1. a busca e download de anexos PDF de e-mails do Outlook,

2. a leitura do conteúdo via OCR (doctr),

3. e a organização automática dos arquivos em pastas, conforme palavras-chave detectadas no texto.

Ideal para automatizar fluxos de trabalho com notas fiscais, relatórios ou documentos digitalizados.

## ⚙️ Funcionalidades

### 🔍 Busca automática no Outlook (via win32com.client)

- Lê o caminho das pastas a partir de outlook_folder.txt

- Filtra e-mails por assunto (target_keyword.txt)

- Baixa anexos PDF não lidos

### 📄 Leitura OCR de PDFs (via doctr)

Extrai texto e identifica palavras definidas em palavra_alvo.txt

### 🗂️ Organização automática

Move os PDFs para subpastas nomeadas conforme a palavra encontrada

Renomeia os arquivos com a data atual (DDMMYYYY.pdf)

## 🪄 Como usar

### 1. Instale as dependências
```
pip install -r requirements.txt
```

### 2. Configure os arquivos .txt

Crie os seguintes arquivos no diretório raiz:

🗂️ outlook_folder.txt
```
# Exemplo com as pastasd o Outlook
Caixa de Entrada/Notas Fiscais
```
💬 target_keyword.txt
```
# Exemplo com a palavra a ser buscada no assunto do e-mail
Notas fiscal anexada
```
🔑 palavra_alvo.txt
```
# Exemplo de palavras usadas como filtro na OCR
Nota fiscal 1
Nota fiscal 2
...
Nota fiscal n
```
(Cada linha representa uma palavra que a OCR vai buscar nos PDFs.)

### 3. Execute o script principal
```
python fluxo.py
```
O script vai:

1. baixar anexos recentes da pasta configurada no Outlook;

2. executar OCR em todos os PDFs baixados;

3. mover e renomear cada arquivo conforme o conteúdo detectado.
   
### 4. 📦 Saída

Os PDFs processados ficam organizados dentro do diretório de download configurado em baixarAnexo.py (por padrão C:\temp\Teste):
```
C:\temp\Teste\
 ├── Nota fiscal 1\
 │   ├── DDMMYYYY.pdf
 ├── Nota fiscal 2\
 │   ├── DDMMYYYY.pdf
 └── SEM_MATCH\
     ├── DDMMYYYY.pdf
```
(Caso ele não encontre a palavra nas notas fiscais, ele criará uma pasta separada "SEM_MATCH")

## 🧠 Tecnologias usadas

Python 3.14

- win32com.client → integração com Microsoft Outlook

- docTR (Deep Optical Character Recognition) → OCR em PDFs

- datetime, shutil, os → automação e manipulação de arquivos

Exemplo:

<img width="942" height="240" alt="image" src="https://github.com/user-attachments/assets/31d07933-83a6-43ee-8c4e-9f5f4e6b7326" />

import os
from datetime import datetime, timedelta

import win32com.client

# Pasta root para salvar os anexos
download_folder = r"C:\temp\Teste"
# E-mails serão procurados até esses dias atrás
days_back = int(input("Até quantos dias atrás ele deve buscar e-mails?: "))
# Quantidade máxima de e-mails para baixar os anexos
max_emails = int(input("Quantos e-mails serão baixados?: "))
# Armazena a target_keyword no código
with open("target_keyword.txt", "r", encoding="utf-8") as f:
    for line in f:
        s = line.strip()
        target_keyword = s

# Configuração do Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
# Armazena o outlook_folder no código
with open("outlook_folder.txt", "r", encoding="utf-8") as f:
    path = f.read().strip()
folder = outlook.GetDefaultFolder(6)  # Inbox
for name in filter(None, path.replace("\\", "/").split("/")):
    folder = folder.Folders[name]

# Salva os anexos
def salvar_pdfs(message, folder_name):
    for i in range(1, message.Attachments.Count + 1):
        att = message.Attachments.Item(i)
        filename = att.FileName.lower()

        if filename.endswith(".msg"):  # .msg = Item do Outlook
            temp_path = os.path.join(download_folder, att.FileName)
            att.SaveAsFile(temp_path)
            try:
                msg_item = outlook.Application.CreateItemFromTemplate(temp_path)
                for j in range(1, msg_item.Attachments.Count + 1):
                    inner_att = msg_item.Attachments.Item(j)
                    if inner_att.FileName.lower().endswith(".pdf"):
                        inner_save = os.path.join(download_folder, inner_att.FileName)
                        inner_att.SaveAsFile(inner_save)
                        print(f"PDF baixado de '{folder_name}': {inner_att.FileName}")
            except Exception as e:
                print(f"Erro ao abrir .MSG '{att.FileName}': {e}")
            finally:
                try:
                    os.remove(temp_path)
                except Exception as e:
                    print(f"Exception {e}")
                    pass

# Percorre as pastas do Outlook e marca como lido
def percorrer_pastas(folder):
    try:
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        since_dt = datetime.now() - timedelta(days=days_back)
        since_str = since_dt.strftime("%m/%d/%Y %I:%M %p")

        restriction = f"[Unread] = True AND [ReceivedTime] >= '{since_str}'"
        filtered = messages.Restrict(restriction)

        alvo = []
        for message in filtered:
            if (
                message.Class == 43
                and target_keyword.lower() in (message.Subject or "").lower()
            ):
                alvo.append(message)
                if len(alvo) >= max_emails:
                    break

        for msg in alvo:
            salvar_pdfs(msg, folder.Name)
            msg.Unread = False
            msg.Save()

        print(
            f"\nProcessados {len(alvo)} e-mails não lidos contendo '{target_keyword}'.\n"
        )
    except Exception as e:
        print(f"Erro na pasta '{folder.Name}': {e}")

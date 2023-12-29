import win32com.client
from datetime import datetime, timedelta
import os
import zipfile
import shutil

# Configurar o cliente do Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Acessar a caixa de entrada
inbox = outlook.GetDefaultFolder(6)  # 6 é o código para a caixa de entrada

# Obter a data e hora atuais
now = datetime.now()
# Ajustar para o início do dia (meia-noite)
start_of_day = datetime(now.year, now.month, now.day, 0, 0, 0)

# Converter para string no formato de data/hora do Outlook
start_of_day_str = start_of_day.strftime('%m/%d/%Y %H:%M %p')

# Criar uma restrição para buscar e-mails a partir do início do dia
restriction = f"[ReceivedTime] >= '{start_of_day_str}'"

# Buscar e-mails com a restrição aplicada
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Ordenar por data de recebimento, do mais recente para o mais antigo
messages = messages.Restrict(restriction)

# Caminhos dos arquivos que serão substituídos
file_to_replace_1 = r"CAMINHO DO SEU ARQUIVO\Volumetria Backlog SN.xlsx"
file_to_replace_2 = r"CAMINHO DO SEU ARQUIVO\m2m_kb_task.xlsx" 
file_to_replace_3 = r"CAMINHO DO SEU ARQUIVO\Pesquisa SNOW.xlsx"
file_to_replace_4 = r"CAMINHO DO SEU ARQUIVO\task_sla.xlsx"

# Pasta temporária para extrair o arquivo zip
zip_temp_folder = r"CAMINHO DO SEU ARQUIVO\Testescript\temp_zip"
if not os.path.exists(zip_temp_folder):
    os.makedirs(zip_temp_folder)

# Iterar sobre os e-mails encontrados
for message in messages:
    # Verificar se o assunto corresponde ao desejado
    if message.Subject == 'Em aberto Total - Lista - (N1;N2;N4)':
        attachments = message.Attachments
        if attachments.Count > 0:
            attachment = attachments.Item(1)
            attachment.SaveAsFile(file_to_replace_1)
            print(f"Anexo '{attachment.FileName}' foi salvo como '{file_to_replace_1}'")
    elif message.Subject == 'QBC - KB - Base de conhecimento':
        attachments = message.Attachments
        if attachments.Count > 0:
            attachment = attachments.Item(1)
            attachment.SaveAsFile(file_to_replace_2)
            print(f"Anexo '{attachment.FileName}' foi salvo como '{file_to_replace_2}'")
    elif message.Subject == 'Pesquisa de Satisfação - N1-N2-N4 v2':
        attachments = message.Attachments
        if attachments.Count > 0:
            attachment = attachments.Item(1)
            attachment.SaveAsFile(file_to_replace_3)
            print(f"Anexo '{attachment.FileName}' foi salvo como '{file_to_replace_3}'")
    elif 'SLA N4 Chamados' in message.Subject:  # Verifique se o assunto do e-mail corresponde exatamente
        attachments = message.Attachments
        if attachments.Count > 0:
            attachment = attachments.Item(1)
            zip_file_path = os.path.join(zip_temp_folder, attachment.FileName)
            attachment.SaveAsFile(zip_file_path)
            print(f"Arquivo zip '{attachment.FileName}' foi salvo temporariamente como '{zip_file_path}'")
            
            # Extrair o arquivo específico do zip
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                # Nome do arquivo dentro do .zip
                nome_arquivo_dentro_zip = 'SLA N1-N2-N4 (Encerrados + Backlog) v2.xlsx'
                zip_ref.extract(nome_arquivo_dentro_zip, zip_temp_folder)
                
                # Caminho do arquivo extraído
                extracted_file_path = os.path.join(zip_temp_folder, nome_arquivo_dentro_zip)
                
                # Substituir o arquivo existente pelo extraído
                if os.path.exists(extracted_file_path):
                    shutil.move(extracted_file_path, file_to_replace_4)
                    print(f"Arquivo '{nome_arquivo_dentro_zip}' foi extraído e movido para '{file_to_replace_4}'")
else:
    print("Não foram encontrados e-mails correspondentes desde o início do dia atual.")

# Limpando a pasta temporária após a extração
if os.path.exists(zip_temp_folder):
    shutil.rmtree(zip_temp_folder)

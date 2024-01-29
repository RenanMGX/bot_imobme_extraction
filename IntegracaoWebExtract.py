from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.credential.carregar_credenciais import Credential
import os
from getpass import getuser
from Entities.sftp.sftp import TransferenciaSFTP
from Entities.registro.registro import Registro
from shutil import copy2
from datetime import datetime

if __name__ == "__main__":
    reg = Registro("IntegracaoWebExtract")
    crendential = Credential()
    entrada = crendential.credencial()
    if (entrada['usuario'] == None) or (entrada['senha'] == None):
       raise PermissionError("Credenciais Invalidas")

    down_path = f"{os.getcwd()}\\downloads_IntegracaoWeb\\"
             
    bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
    conversor = ImobmeExceltoConvert()

    bot_relatorio.obter_relatorios(["imobme_dados_contrato", "imobme_empreendimento"])

    arquivos = []
    for files in os.listdir(down_path):
       arquivos.append(down_path + files)

    caminho_temp = "temp_IntegracaoWeb"
    if not os.path.exists(caminho_temp):
       os.makedirs(caminho_temp)
    conversor.tratar_arquivos(arquivos, path_data="dados_IntegracaoWeb", tipo="csv_integra_web", copyto=caminho_temp)

    arquivos_do_temp = []
    for arqui in os.listdir(caminho_temp):
        arquivos_do_temp.append(caminho_temp +  "\\" + arqui)
    
    from credential_sftp import crendencial_sftp_apart
    crendencial_sftp = crendencial_sftp_apart
    
    #transferir arquivo via SFTP
    transfer = TransferenciaSFTP(crendencial_sftp)
    for arqui in arquivos_do_temp:
        nome_arquivo = arqui.split('\\')[-1]
        try:
            transfer.transferir(arqui, f"public_html/bases/{nome_arquivo}")
        except TimeoutError:
            reg.record("TimeoutError;Uma tentativa de conexão falhou porque o componente conectado não respondeu corretamente após um período de tempo ou a conexão estabelecida falhou porque o host conectado não respondeu")

    for arqui in arquivos_do_temp:
        copy2(arqui, r"C:\Users\renan.oliveira\Downloads")

    for arqui_del in os.listdir(caminho_temp):
        os.unlink(caminho_temp + "\\" + arqui_del)
    os.rmdir(caminho_temp)

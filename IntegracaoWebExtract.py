from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.dependencies.credenciais import Credential
import os
from getpass import getuser
from Entities.sftp.sftp import TransferenciaSFTP
#from Entities.registro.registro import Registro
from Entities.dependencies.logs import Logs
from shutil import copy2
from datetime import datetime
from Entities.dependencies.config import Config
import shutil
import traceback

TRANSFERIR_FTP:bool = True

if __name__ == "__main__":
    try:
        entrada:dict = Credential(Config()['credential']['crd']).load()
        if (entrada['login'] == None) or (entrada['password'] == None):
            raise PermissionError("Credenciais Invalidas")

        down_path = os.path.join(os.getcwd(), "downloads_IntegracaoWeb\\")
        
        ## extrair relatorio        
        bot_relatorio = BotExtractionImobme(user=entrada['login'],password=entrada['password'],download_path=down_path)

        bot_relatorio.start([
            "imobme_dados_contrato",
            "imobme_empreendimento"
            ])
        ## extrair relatorio  -- fim

        arquivos = []
        for files in os.listdir(down_path):
            arquivos.append(down_path + files)

        caminho_temp = os.path.join(os.getcwd(), "temp_IntegracaoWeb")
        if not os.path.exists(caminho_temp):
            os.makedirs(caminho_temp)
        else:
            for file in os.listdir(caminho_temp):
                file = os.path.join(caminho_temp, file)
                if os.path.isfile(file):
                    os.unlink(file)
                elif os.path.isdir(file):
                    shutil.rmtree(file)

        ImobmeExceltoConvert(path=down_path).extract_csv_integraWeb(caminho_temp)

        arquivos_do_temp = []
        for arqui in os.listdir(caminho_temp):
            arqui = os.path.join(caminho_temp, arqui)
            arquivos_do_temp.append(arqui)
    
        from credential_sftp import crendencial_sftp_apart
        crendencial_sftp = crendencial_sftp_apart
    
        #transferir arquivo via SFTP
        if TRANSFERIR_FTP:
            transfer = TransferenciaSFTP(crendencial_sftp)
            for arqui in arquivos_do_temp:
                nome_arquivo = arqui.split('\\')[-1]
                transfer.transferir(arqui, f"public_html/bases/{nome_arquivo}")
        else:
            for arqui in arquivos_do_temp:
                copy2(arqui, f"C:\\Users\\{getuser()}\\Downloads")
        Logs().register(status='Concluido', description='Extração de relatorios da Integração web foi concluida!')
    except Exception as err:
        Logs().register(status='Error', description='erro ao executar o extração dos relatorios da integração web', exception=traceback.format_exc())


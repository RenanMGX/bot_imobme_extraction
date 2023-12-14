from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.credential.carregar_credenciais import Credential
import os
from getpass import getuser
from Entities.registro.registro import Registro

if __name__ == "__main__":
    reg = Registro("SambaExtract.py")
    try:
        crendential = Credential()
        entrada = crendential.credencial()
        if (entrada['usuario'] == None) or (entrada['senha'] == None):
            raise PermissionError("Credenciais Invalidas")

        down_path = f"{os.getcwd()}\\downloads_samba\\"

        bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
        conversor = ImobmeExceltoConvert()

        bot_relatorio.obter_relatorios(["imobme_controle_vendas", "imobme_contratos_rescindidos"])

        arquivos = []
        for files in os.listdir(down_path):
            arquivos.append(down_path + files)

        conversor.tratar_arquivos(arquivos, path_data="dados_samba", copyto=f'C:\\Users\\{getuser()}\\OneLake - Microsoft\\DW_BI\\lake_house.Lakehouse\\Files\\jsons\\VendasContratos\\')
    
    except Exception as error:
        reg.record(f"{type(error)};{error}")

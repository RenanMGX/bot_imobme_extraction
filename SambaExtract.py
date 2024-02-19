import sys
sys.path.append("Entities")
from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.credential.carregar_credenciais import Credential
import os
from getpass import getuser
from Entities.registro.registro import Registro

if __name__ == "__main__":
    reg = Registro("SambaExtract.py")

    crendential = Credential()
    entrada = crendential.credencial()
    if (entrada['usuario'] == None) or (entrada['senha'] == None):
        raise PermissionError("Credenciais Invalidas")

    down_path = f"{os.getcwd()}\\downloads_samba\\"

    for x in range(3):
        try:
            bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)

            bot_relatorio.obter_relatorios(["imobme_controle_vendas_90_dias", "imobme_contratos_rescindidos_90_dias", "imobme_relacao_clientes_x_clientes"])

            final = ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\OneLake - Microsoft\\DW_BI\\lake_house.Lakehouse\\Files\\jsons\\VendasContratos\\')
            if final:
                break
        except Exception as error:
            reg.record(f"{type(error)};{error}")
        finally:
            try:
                bot_relatorio.navegador.close()
                del bot_relatorio
            except:
                pass

import sys
sys.path.append("Entities")
from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.credential.carregar_credenciais import Credential
import os
from getpass import getuser
from Entities.registro.registro import Registro
import traceback

if __name__ == "__main__":
    reg = Registro("SambaExtract.py")

    entrada: dict = Credential().credencial()
    if (entrada['usuario'] == None) or (entrada['senha'] == None):
        raise PermissionError("Credenciais Invalidas")

    down_path = f"{os.getcwd()}\\downloads_samba\\"

    for x in range(3):
        try:
            bot_relatorio = BotExtractionImobme(user=entrada['usuario'],password=entrada['senha'],download_path=down_path)

            bot_relatorio.start(["imobme_controle_vendas_90_dias", "imobme_contratos_rescindidos_90_dias", "imobme_relacao_clientes_x_clientes"])

            final = ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\OneLake - Microsoft\\BI_HML\\lake_house.Lakehouse\\Files\\jsons\\VendasContratos\\')
            if final:
                bot_relatorio.navegador.close()
                break
        except Exception as error:
                erro_trace = traceback.format_exc()
                print(erro_trace)
                erro_trace = erro_trace.replace("\n", "|||")
                reg.record(f"{type(error)};{error} traceback:  {erro_trace}")
        finally:
            try:
                bot_relatorio.navegador.close()
                del bot_relatorio
            except:
                pass

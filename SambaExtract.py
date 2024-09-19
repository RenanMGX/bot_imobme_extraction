import sys
sys.path.append("Entities")
from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.dependencies.credenciais import Credential
import os
from getpass import getuser
from Entities.dependencies.logs import Logs
import traceback
from datetime import datetime
from Entities.dependencies.config import Config

if __name__ == "__main__":
    try:
        reg = Logs()

        entrada: dict = Credential(Config()['credential']['crd']).load()
        if (entrada['login'] == None) or (entrada['password'] == None):
            raise PermissionError("Credenciais Invalidas")

        down_path = f"{os.getcwd()}\\downloads_samba\\"

        for x in range(3):
            try:
                bot_relatorio = BotExtractionImobme(user=entrada['login'],password=entrada['password'],download_path=down_path)

                bot_relatorio.start(["imobme_controle_vendas_90_dias", "imobme_contratos_rescindidos_90_dias", "imobme_relacao_clientes_x_clientes"])

                final = ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\OneLake - Microsoft\\BI_HML\\lake_house.Lakehouse\\Files\\jsons\\VendasContratos\\')
                if final:
                    bot_relatorio.navegador.close()
                    break
            except Exception as error:
                    erro_trace = traceback.format_exc()
                    print(erro_trace)
                    erro_trace = erro_trace.replace("\n", "|||")
                    reg.register(status='Error', description=str(error), exception=erro_trace)
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                except:
                    pass
    except Exception as error:
        path:str = "logs/"
        if not os.path.exists(path):
            os.makedirs(path)
        file_name = path + f"LogError_{datetime.now().strftime('%d%m%Y%H%M%Y')}.txt"
        with open(file_name, 'w', encoding='utf-8')as _file:
            _file.write(traceback.format_exc())
        raise error

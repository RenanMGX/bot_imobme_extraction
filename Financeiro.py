from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.dependencies.credenciais import Credential
import os
from getpass import getuser
#from Entities.registro.registro import Registro
from Entities.dependencies.logs import Logs
from datetime import datetime
import traceback
from Entities.dependencies.config import Config

def error_except(error):
    erro_trace = traceback.format_exc()
    print(erro_trace)
    erro_trace = erro_trace.replace("\n", " <br> ")
    Logs().register(status='Report', description=str(error), exception=erro_trace)
    

if __name__ == "__main__":
    tempo_agora = datetime.now()
    try:
        
        entrada: dict = Credential(Config()['credential']['crd']).load()
        if (entrada['login'] == None) or (entrada['password'] == None):
            raise PermissionError("Credenciais Invalidas")

        down_path = f"{os.getcwd()}\\downloads_financeiro\\"

        #primeira parte
        for x in range(5):
            try:   
                bot_relatorio = BotExtractionImobme(user=entrada['login'],password=entrada['password'],download_path=down_path)
                
                bot_relatorio.start([
                    "imobme_dados_contrato",
                    "imobme_relacao_clientes",
                    "imobme_controle_vendas",
                    "imobme_controle_estoque",
                    "imobme_contratos_rescindidos",
                    "imobme_cadastro_datas",
                    "imobme_empreendimento"
                ])
                ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                break
            except Exception as error:
                error_except(error)
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                except:
                    continue
        #fim primeira parte        
         
        #segunda parte        
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(user=entrada['login'],password=entrada['password'],download_path=down_path)
                
                bot_relatorio.start([
                    "recebimentos_compensados",
                    "imobme_relacao_clientes_x_clientes",
                    "imobme_previsao_receita"
                ])
                ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                break
            except Exception as error:
                error_except(error)
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                except:
                    pass
        #fim segunda parte
        Logs().register(status='Concluido', description='Extração de relatorios do Financeiro concluida!')
    except Exception as err:
        Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
        

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
import multiprocessing
from time import sleep

def execute(lista:list, download:str="fin"):
        entrada: dict = Credential(Config()['credential']['crd']).load()
        if (entrada['login'] == None) or (entrada['password'] == None):
            raise PermissionError("Credenciais Invalidas")

        path:str=f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\'
        #down_path = f"{os.getcwd()}\\downloads_financeiro\\"
        down_path = os.path.join(os.getcwd(), f"downloads_{download}") + "\\"
        for _ in range(5):
            try:   
                bot_relatorio = BotExtractionImobme(user=entrada['login'],password=entrada['password'],download_path=down_path)
                
                bot_relatorio.start(lista)
                ImobmeExceltoConvert(path=down_path).extract_json(path)
                break
            
            except TimeoutError as err_t:
                Logs().register(status='Report', description=str(err_t), exception=traceback.format_exc())
            except Exception as err:
                Logs().register(status='Report', description=str(err), exception=traceback.format_exc())
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                except:
                    continue
            if _ >= 5:
                Logs().register(status='Error', description=f"Numero de tentativas exedida ao extrair a os relatorios {lista}")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    tempo_agora = datetime.now()
    try:  
        
        primeiro = multiprocessing.Process(target=execute, args=([
            "imobme_dados_contrato",
            "imobme_relacao_clientes",
            "imobme_controle_vendas",
            "imobme_controle_estoque",
            "imobme_contratos_rescindidos",
            "imobme_cadastro_datas",
            "imobme_empreendimento"
        ],f'Relatorio_Imobme_Financeiro_1',))
        
        segundo = multiprocessing.Process(target=execute, args=([
            "recebimentos_compensados",
            "imobme_relacao_clientes_x_clientes"                    
        ],f'Relatorio_Imobme_Financeiro_2',))
        
        terceiro = multiprocessing.Process(target=execute, args=([
            "imobme_previsao_receita",                 
        ],f'Relatorio_Imobme_Financeiro_3',))
        
        
        primeiro.start()
        sleep(5)
        segundo.start()
        sleep(5)
        terceiro.start()
        sleep(5)
        
        primeiro.join()
        segundo.join()
        terceiro.join()
        
        
        Logs().register(status='Concluido', description='Extração de relatorios do Financeiro concluida!')
    except Exception as err:
        Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
        

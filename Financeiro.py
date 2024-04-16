from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.credential.carregar_credenciais import Credential
import os
from getpass import getuser
from Entities.registro.registro import Registro
from datetime import datetime
import traceback

if __name__ == "__main__":
    reg = Registro("Financeiro")
    tempo_agora = datetime.now()
    try:
        
        entrada: dict = Credential().credencial()
        if (entrada['usuario'] == None) or (entrada['senha'] == None):
            raise PermissionError("Credenciais Invalidas")

        down_path = f"{os.getcwd()}\\downloads_financeiro\\"

        #primeira parte
        for x in range(3):
            try:   
                bot_relatorio = BotExtractionImobme(user=entrada['usuario'],password=entrada['senha'],download_path=down_path)
                
                bot_relatorio.start([
                    "imobme_dados_contrato",
                    "imobme_relacao_clientes",
                    "imobme_controle_vendas",
                    "imobme_controle_estoque",
                    "imobme_contratos_rescindidos",
                    "imobme_cadastro_datas",
                    "imobme_empreendimento"
                ])
                final = ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                if final:
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
                    continue
        #fim primeira parte        
         
         
        #segunda parte        
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(user=entrada['usuario'],password=entrada['senha'],download_path=down_path)
                
                bot_relatorio.start([
                    "recebimentos_compensados",
                    "imobme_relacao_clientes_x_clientes",
                    "imobme_previsao_receita"
                ])
                ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
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
        #fim segunda parte
                
    except Exception as error:
        erro_trace = traceback.format_exc().replace("\n", "|||")
        reg.record(f"{type(error)};{error} traceback:  {erro_trace}")
    
    print(datetime.now() - tempo_agora)

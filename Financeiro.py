from Entities.tratar_arquivos_excel_imobme import ImobmeExceltoConvert
from Entities.extraction_imobme import BotExtractionImobme
from Entities.credential.carregar_credenciais import Credential
import os
from getpass import getuser
from Entities.registro.registro import Registro
from datetime import datetime

if __name__ == "__main__":
    reg = Registro(__file__)
    tempo_agora = datetime.now()
    try:
        crendential = Credential()
        entrada = crendential.credencial()
        if (entrada['usuario'] == None) or (entrada['senha'] == None):
            raise PermissionError("Credenciais Invalidas")

        down_path = f"{os.getcwd()}\\downloads_financeiro\\"

        for x in range(3):
            try:   
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                
                bot_relatorio.obter_relatorios([
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
                reg.record(f"{type(error)};{error}")
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                except:
                    pass
                
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                
                bot_relatorio.obter_relatorios([
                    "recebimentos_compensados"
                ])
                final = ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
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
                
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                
                bot_relatorio.obter_relatorios([
                    "imobme_previsao_receita",
                ])
                final = ImobmeExceltoConvert(path=down_path).extract_json(f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
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

    except Exception as error:
        reg.record(f"{type(error)};{error}")
    
    print(datetime.now() - tempo_agora)

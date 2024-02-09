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
        

        for x in range(5):
            try:   
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                conversor = ImobmeExceltoConvert()
                
                bot_relatorio.obter_relatorios([
                    "imobme_dados_contrato",
                    "imobme_relacao_clientes"
                ])
                arquivos = []
                for files in os.listdir(down_path):
                    arquivos.append(down_path + files)
                final = conversor.tratar_arquivos(arquivos, path_data="dados_financeiro", copyto=f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                if final:
                    break
            except Exception as error:
                reg.record(f"{type(error)};{error}")
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                    del conversor
                except:
                    pass
                
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                conversor = ImobmeExceltoConvert()
                
                bot_relatorio.obter_relatorios([
                    "recebimentos_compensados",
                    "imobme_controle_vendas"
                ])
                arquivos = []
                for files in os.listdir(down_path):
                    arquivos.append(down_path + files)
                final = conversor.tratar_arquivos(arquivos, path_data="dados_financeiro", copyto=f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                if final:
                    break
            except Exception as error:
                reg.record(f"{type(error)};{error}")
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                    del conversor
                except:
                    pass
                
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                conversor = ImobmeExceltoConvert()
                
                bot_relatorio.obter_relatorios([
                    "imobme_controle_estoque",
                    "imobme_contratos_rescindidos"
                ])
                arquivos = []
                for files in os.listdir(down_path):
                    arquivos.append(down_path + files)
                final = conversor.tratar_arquivos(arquivos, path_data="dados_financeiro", copyto=f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                if final:
                    break
            except Exception as error:
                reg.record(f"{type(error)};{error}")
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                    del conversor
                except:
                    pass
                
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                conversor = ImobmeExceltoConvert()
                
                bot_relatorio.obter_relatorios([
                    "imobme_cadastro_datas",
                    "imobme_empreendimento"
                ])
                arquivos = []
                for files in os.listdir(down_path):
                    arquivos.append(down_path + files)
                final = conversor.tratar_arquivos(arquivos, path_data="dados_financeiro", copyto=f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                if final:
                    break
            except Exception as error:
                reg.record(f"{type(error)};{error}")
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                    del conversor
                except:
                    pass
                
        for x in range(5):
            try:
                bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
                conversor = ImobmeExceltoConvert()
                
                bot_relatorio.obter_relatorios([
                    "imobme_previsao_receita",
                ])
                arquivos = []
                for files in os.listdir(down_path):
                    arquivos.append(down_path + files)
                final = conversor.tratar_arquivos(arquivos, path_data="dados_financeiro", copyto=f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')
                if final:
                    break
            except Exception as error:
                reg.record(f"{type(error)};{error}")
            finally:
                try:
                    bot_relatorio.navegador.close()
                    del bot_relatorio
                    del conversor
                except:
                    pass

    except Exception as error:
        reg.record(f"{type(error)};{error}")
    
    print(datetime.now() - tempo_agora)

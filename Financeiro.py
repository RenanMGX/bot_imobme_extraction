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

        bot_relatorio = BotExtractionImobme(usuario=entrada['usuario'],senha=entrada['senha'],caminho_download=down_path)
        conversor = ImobmeExceltoConvert()

        
        bot_relatorio.obter_relatorios([
            "imobme_previsao_receita",
            "imobme_dados_contrato",
            "imobme_relacao_clientes",
            "imobme_cadastro_datas",
            "imobme_empreendimento",
            "recebimentos_compensados",
            "imobme_contratos_rescindidos",
            "imobme_controle_vendas",
            "imobme_controle_estoque"

        ])

        arquivos = []
        for files in os.listdir(down_path):
            arquivos.append(down_path + files)

        conversor.tratar_arquivos(arquivos, path_data="dados_financeiro", copyto=f'C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Relatorio_Imobme_Financeiro\\')

        
    
    except Exception as error:
        reg.record(f"{type(error)};{error}")
    
    print(datetime.now() - tempo_agora)

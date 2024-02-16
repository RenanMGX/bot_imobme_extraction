import os
import sys
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
from datetime import datetime
from dateutil.relativedelta import relativedelta
try:
    from Entities.registro.registro import Registro
except ModuleNotFoundError:
    from registro.registro import Registro

class BotExtractionImobme():
    def __init__(self,usuario=None,senha=None, caminho_download=f"{os.getcwd()}\\downloads\\"):
        self.__registro_error = Registro("extraction_imobme")
        if (usuario == None) or (senha == None):
            raise ValueError("usuario ou senha estão vazios")
        
        self.caminho_download = caminho_download
        self.temp_variable = None
        self.relatorios_id = []

        self.login = [
            {'action' : self.escrever, 'kargs' : {'target' : '//*[@id="login"]', 'input' : usuario}}, #escreve o usuario no campo do usuario
            {'action' : self.escrever, 'kargs' : {'target' : '//*[@id="password"]', 'input' : senha}}, #escreve o senha no campo da senha
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="btnLogin"]'}}, #clica no botão logar
            {'action' : self.clicar, 'kargs' : {'target' : '/html/body/div[2]/div[3]/div/button[1]/span'}}, #se o usuario já estiver logado clica logoff
            {'action' : self.finalizador_de_emergencia, 'kargs' : {'target' : '/html/body/div[1]/div/div/div/div[2]/form/div/ul/li', 'verific' : {"regra" : "in", "texto": "Senha Inválida"}}}, # faz uma verificação se o login ou a senha está correta caso está incorreto ele encerra todo o script para não bloquear a conta
            {'action' : self.finalizar, 'kargs' : {'target' : '//*[@id="login"]', 'exist' : False}} # se não achar o campo do login ele finaliza o roteiro
        ]
        self.ir_relatorios = [
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Menu"]/ul/li[3]/a'}}, # clicar no icone dos relatorios
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Menu"]/ul/li[3]/div/ul/li/a'}}, # clica no botão gerar relatorios
            {'action' : self.finalizar, 'kargs' : {'target' : '//*[@id="result-table"]/tbody', 'exist' : True}} # finaliza o roteiro se achar a lista dos relatorios
        ]
        self.verificar_lista = [
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="btnProximaDefinicao"]'}}, #clica no botão de atualizar a lista
            {'action' : self.esperar, 'kargs' : {'segundos' : 2}}, # colocar uma espera
            {'action' : self.salvar, 'kargs' : {'target' : '//*[@id="result-table"]/tbody'}}, # salva o conteudo da tabela em uma instancia global self.temp_variable
            {'action' : self.finalizar, 'kargs' : {'target' : '//*[@id="result-table"]/tbody', 'exist' : True}} #finaliza apos verificar que a tabela realmente carregou
        ]
        self.gerar_relatorios = [
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}, # clique em selecionar Relatorios
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_13"]'}}, # clique em IMOBME - Empreendimento
            {'action' : self.esperar, 'kargs' : {'segundos' : 1}}, # colocar uma espera
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}, # clique em selecionar Emprendimentos
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input'}}, # clique em selecionar todos os empreendimentos
            {'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}, # clica em gerar relatorio
            {'action' : self.finalizar, 'kargs' : {'target' : '//*[@id="result-table"]/tbody', 'exist' : True}} # verifica para encerrar esté roteiro
        ]

    def obter_relatorios(self,relatorios=None):
        '''
        esté metodo monta um roteiro para o bot gerar o relatorio no imobme
        relatorios registrador no script são:
            imobme_empreendimento
            imobme_controle_vendas
            imobme_contratos_rescindidos
            imobme_dados_contrato
        '''
        if isinstance(relatorios, list):
            print(relatorios)
            self.relatorios = relatorios
            if len(relatorios) == 0:
                print("A lista de 'relatorios' não pode está vazia")
                return False
            
            self.gerar_relatorios = []
            verificar_se_tem_relatorios = 0
            for relatorio in relatorios:
                relatorio = str(relatorio)

                # para gerar o 'Relatorios de Empreendimento'
                if (relatorio.lower() == "imobme_empreendimento") or (relatorio.lower() == "empreendimentos"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_7"]'}}) # clique em IMOBME - Empreendimento
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}})  # clique em selecionar Emprendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input'}}) # clique em selecionar todos os empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}})  # clique em selecionar Emprendimentos novamente para sair
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                # para gerar o relatorio 'Contole de Vendas'
                elif (relatorio.lower() == "imobme_controle_vendas") or (relatorio.lower() == "vendas"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_5"]'}}) # clique em IMOBME - Contre de Vendas
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataInicio"]', 'input' : (datetime.now() - relativedelta(days=90)).strftime("%d%m%Y")  }})  # escreve a data de inicio padrao 01/01/2015
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataFim"]', 'input' : datetime.now().strftime("%d%m%Y")}})  # escreve a data hoje
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="TipoDataSelecionada_chzn"]/a'}}) # clique em Tipo Data
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="TipoDataSelecionada_chzn_o_0"]'}}) # clique em Data Lançamento Venda
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clique em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input'}}) # clique em Todos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clique em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                # para gerar o relatorio 'Contrator Rescindidos'
                elif (relatorio.lower() == "imobme_contratos_rescindidos") or (relatorio.lower() == "contratosrescindidos"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_3"]'}}) # clique em IMOBME - Contratos Rescindicos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataInicio"]', 'input' : (datetime.now() - relativedelta(days=90)).strftime("%d%m%Y")}})  # escreve a data de inicio padrao 01/01/2015
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataFim"]', 'input' : datetime.now().strftime("%d%m%Y")}})  # escreve a data hoje
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button'}}) # clique em Tipo de Contrato
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/ul/li[2]/a/label/input'}}) # clique em Todos
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button'}}) # clique em Tipo de Contrato
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em empreendimentos
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input'}}) # clique em Todos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                # para gerar o relatorio 'Dados do Contrato'
                elif (relatorio.lower() == "imobme_dados_contrato") or (relatorio.lower() == "dadoscontrato"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_6"]'}}) # clique em IMOBME - Dados de Contrato
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label'}}) # clica em todos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[1]/div'}}) # clica fora
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[3]/div[1]/div/button'}}) # clica em tipos de contrato
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[3]/div[1]/div/ul/li[3]/a/label'}}) # clica em PCV
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[3]/div[1]/div/ul/li[5]/a/label'}}) # clica em Cessao
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[1]/div'}}) # clica fora
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[3]/div[2]/div/button'}}) # clica em Status
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[3]/div[2]/div/ul/li[3]/a/label'}}) # clica em Ativo
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[1]/div'}}) # clica fora
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataBase"]', 'input' : datetime.now().strftime("%d%m%Y")}})  # escreve a data hoje
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                # para gerar o relatorio 'Previsão de Receita'
                elif (relatorio.lower() == "imobme_previsao_receita") or (relatorio.lower() == "previsaoreceita"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_8"]'}}) # clique em IMOBME - Previsão de Receita
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataInicio"]', 'input' : "01012015"}})  # escreve a data de inicio padrao 01/01/2015
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataFim"]', 'input' : (datetime.now() + relativedelta(years=25)).strftime("%d%m%Y") }})  # escreve a data de fim padrao com a data atual mais 25 anos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input'}}) # clica em todos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos novamente para sair
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataBase"]', 'input' : datetime.now().strftime("%d%m%Y") }})  # escreve a data de inicio padrao 01/01/2015
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[4]/div/div[2]/div/button'}}) # clica em Tipo Parcela
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[4]/div/div[2]/div/ul/li[2]/a/label/input'}}) # clica em Todos
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[4]/div/div[2]/div/button'}}) # clica em Tipo Parcela para sair
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                # para gerar o relatorio 'Relação de Clientes'
                elif (relatorio.lower() == "imobme_relacao_clientes") or (relatorio.lower() == "relacaoclientes"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_9"]'}}) # clique em IMOBME - Relação de Clientes
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataInicio"]', 'input' : "01012015"}})  # escreve a data de inicio padrao 01/01/2015
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataFim"]', 'input' : datetime.now().strftime("%d%m%Y") }})  # escreve a data de fim padrao com a data atual mais 25 anos
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                # para gerar o relatorio 'Cadastro de Datas'
                elif (relatorio.lower() == "imobme_cadastro_datas") or (relatorio.lower() == "cadastrodatas"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_2"]'}}) # clique em IMOBME - Previsão de Receita
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label'}}) # clica em todos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                # para gerar o relatorio 'Recebimentos Compensados'
                elif (relatorio.lower() == "recebimentos_compensados") or (relatorio.lower() == "recebimentoscompensados"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_15"]'}}) # clique em Recebimentos Compensados
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataInicio"]', 'input' : "01012020"}})  # escreve a data de inicio padrao 01/01/2015
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataFim"]', 'input' : datetime.now().strftime("%d%m%Y") }})  # escreve a data de fim padrao com a data atual mais 25 anos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label'}}) # clica em todos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="dvEmpreendimento"]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                # para gerar o relatorio 'Controle de Estoque'
                elif (relatorio.lower() == "imobme_controle_estoque") or (relatorio.lower() == "controleestoque"):
                    verificar_se_tem_relatorios += 1
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn"]/a'}}) # clique em selecionar Relatorios
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="Relatorios_chzn_o_4"]'}}) # clique em IMOBME - Controle de Estoque
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[2]/div[1]/div/div/button'}}) # clica em Empreendimentos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}}) # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[2]/div[1]/div/div/ul/li[2]/a/label/input'}}) # clica em todos
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}}) # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="parametrosReport"]/div[2]/div[1]/div/div/button'}}) # clica em Empreendimentos novamente para sair
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 1}}) # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.escrever, 'kargs' : {'target' : '//*[@id="DataBase"]', 'input' : datetime.now().strftime("%d%m%Y")}})  # escreve a data hoje
                    self.gerar_relatorios.append({'action' : self.clicar, 'kargs' : {'target' : '//*[@id="GerarRelatorio"]'}}) # clica em gerar relatorio
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.coletar_id_relatorio, 'kargs' : {'target' : '//*[@id="result-table"]/tbody/tr[1]/td[1]'}}) #  # salva o numero do relatorio gerado em uma instancia global self.relatorios_id
                    self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 2}})  # colocar uma espera
                    self.gerar_relatorios.append({'action' : self.finalizar_relatorio, 'kargs' : {'target' : '//*[@id="Content"]/section/div[2]/div[1]/div[1]/div'}}) # verifica para encerrar esté roteiro)
                
                
                self.gerar_relatorios.append({'action' : self.esperar, 'kargs' : {'segundos' : 3}})  # colocar uma espera                
            self.gerar_relatorios.append({'action' : self.finalizar, 'kargs' : {'target' : '//*[@id="result-table"]/tbody', 'exist' : True}})# verifica para encerrar esté roteiro

            if verificar_se_tem_relatorios > 0:
                return self.iniciar_navegador()
            else:
                print("Nenhum dos relatorios informados existe no script, o bot do imobme será encerrado!")
                return False


        else:
            print("A instancias 'relatorios' deve ser uma lista")
            return False

    def iniciar_navegador(self,debug=False):
        if not os.path.exists(self.caminho_download):
            os.mkdir(self.caminho_download)
        else:
            for arquivo in os.listdir(self.caminho_download):
                try:
                    os.unlink(self.caminho_download + arquivo)
                except:
                    pass
        
        prefs = {"download.default_directory" : self.caminho_download}
        
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", prefs)
        with webdriver.Chrome(options=chrome_options) as self.navegador:
            self.navegador.get("http://patrimarengenharia.imobme.com/Autenticacao/Login")
            sleep(1)
            self.navegador.get("http://patrimarengenharia.imobme.com/Autenticacao/Login")
            #fazer o login
            self.roteiro(self.login, emergency_break=5*60)
            sleep(1)

            #ir até o relatorios
            self.roteiro(self.ir_relatorios)
            sleep(1)

            if debug == True:
                input("###################  Esperar: ")
                return False

            #verificar a lista de relatorios
            self.roteiro(self.verificar_lista)
            lista_para_apagar = self.temp_variable.split("\n")
            self.temp_variable = None
            
            #deleta todos os relatorios se a execução for antes das 09:00 da manha
            if datetime.now().hour < 9:
                #deletar dados se ouver
                for indice,linha in enumerate(lista_para_apagar):
                    self.roteiro([
                                {'action' : self.clicar, 'kargs' : {'target' : f'//*[@id="result-table"]/tbody/tr[{indice + 1}]/td[12]/a'}},
                                {'action' : self.clicar, 'kargs' : {'target' : f'//*[@id="result-table"]/tbody/tr[{indice + 1}]/td[12]/a'}},
                                {'action' : self.clicar, 'kargs' : {'target' : f'/html/body/div[5]/div[3]/div/button[1]'}},
                                {'action' : self.finalizar, 'kargs' : {'target' : f'//*[@id="result-table"]/tbody/tr[{indice + 1}]/td[12]/a', 'exist' : False}}, 
                                    ])
            
            #gerar relatorio
            self.roteiro(self.gerar_relatorios)

            contador_saida = 0
            while True:
                if contador_saida > 2160:
                    self.__registro_error.record(f"saida emergencia acionada a espera da geração dos relatorios superou as 3 horas")
                    raise TimeoutError("saida emergencia acionada a espera da geração dos relatorios superou as 3 horas")
                else:
                    contador_saida += 1

                if len(self.relatorios_id) == 0:
                    break
                tbody = self.navegador.find_element(By.TAG_NAME, 'tbody')
                tr_s = tbody.find_elements(By.TAG_NAME, 'tr')
                for tr in tr_s:

                    id_temp = tr.find_element(By.CLASS_NAME, 'sorting_1').text
                    if id_temp in self.relatorios_id:
                        download_link = tr.find_elements(By.TAG_NAME, 'td')[10]
                        try:
                            down_bt = download_link.find_element(By.TAG_NAME, 'a')
                            #import pdb; pdb.set_trace()
                            down_bt.click()
                            self.relatorios_id.pop(self.relatorios_id.index(id_temp))
                            continue
                        except Exception as error:
                            continue

                
                self.navegador.find_element(By.XPATH, '//*[@id="btnProximaDefinicao"]').click()
                sleep(5)

            # #salvar os arquivos
            # self.roteiro(self.verificar_lista)
            # lista_para_salvar = self.temp_variable.split("\n")
            # self.temp_variable = None

            # #verifica se o botão do download está disponivel quando estiver disponivel ele clica nele e passa para verificar o proximo do download se houver
            # for indice,linha in enumerate(lista_para_salvar):
            #     self.roteiro([
            #                   {'action' : self.clicar, 'kargs' : {'target' : f'//*[@id="btnProximaDefinicao"]'}},
            #                   {'action' : self.esperar, 'kargs' : {'segundos' : 1}},
            #                   {'action' : self.clicar, 'kargs' : {'target' : f'//*[@id="result-table"]/tbody/tr[{indice + 1}]/td[11]/a'}},
            #                   {'action' : self.finalizar, 'kargs' : {'target' : f'//*[@id="result-table"]/tbody/tr[{indice + 1}]/td[11]/a', 'exist' : True}}, 
            #                   {'action' : self.esperar, 'kargs' : {'segundos' : 6}}
            #                     ], emergency_break =1000)
                

            while True:
                download_check = True
                if len(os.listdir(self.caminho_download)):
                    for x in os.listdir(self.caminho_download):
                        if not x.split(".")[-1] == "xlsx":
                            download_check = False
                else:
                    download_check = False
                
                if download_check:
                    break
                else:
                    sleep(1)
            
            sleep(2)
            texto_error = "relatorios: " + str(self.relatorios) + "; foram baixados e salvos no diretorio " + str(self.caminho_download)
            self.__registro_error.record(texto_error, tipo="Concluido")
            
            # caminho_dados = "dados"
            # if not os.path.exists(caminho_dados):
            #     os.makedirs(caminho_dados)
            # caminho_dos_arquivos = []
            # for arquivoss in os.listdir(self.caminho_download):
            #     arquivo_nome = arquivoss.split('\\')[-1:][0]
            #     novo_nome = f"{arquivo_nome.split('_')[0]}.xlsx"
            #     copy2(f"{self.caminho_download}{arquivoss}", f"{caminho_dados}\\{novo_nome}")

            #     caminho_dos_arquivos.append(f"{caminho_dados}\\{novo_nome}")

                
            # return caminho_dos_arquivos
        
    def roteiro(self, roteiro, emergency_break=15):
        contador = 0
        while True:
            for evento in roteiro:
                #print(evento['kargs'])
                eventos = evento['action'](evento['kargs'])
                if eventos == "saida":
                    #print("saindo")
                    return
                elif eventos == "emergencia":
                    print("saida de Emergencia")
                    self.__registro_error.record(f"saida emergencia acionada;{evento['kargs']}")
                    raise TimeoutError(f"saida emergencia acionada;{evento['kargs']}")
            
            if contador <= emergency_break:
                contador += 1
            else:
                print("saida de Emergencia")
                self.__registro_error.record("Exedeu o tempo limite do Emergency_break")
                raise TimeoutError("Exedeu o tempo limite do Emergency_break")
            sleep(1)

    def clicar(self, argumentos):
        try:
            target = self.navegador.find_element(By.XPATH, argumentos['target'])
            target.click()
        except:
            pass
    
    def esperar(self, argumentos):
        sleep(argumentos['segundos'])

    def debug_click(self, argumentos):
        try:
            target = self.navegador.find_element(By.XPATH, argumentos['target'])
            return "saida"
        except:
            return "saida"

    def escrever(self, argumentos):
        try:
            target = self.navegador.find_element(By.XPATH, argumentos['target'])
            target.send_keys(argumentos['input'])
        except:
            pass
    def finalizar(self, argumentos):
        '''
        explicando o termo "exist"
        True: ele vai finalizar a execução se o target for encontrado
        False: ele vai finalizar a execução caso não encontre mais o target 
        '''
        try:
            self.navegador.find_element(By.XPATH, argumentos['target'])
            if argumentos['exist'] == True:
                return "saida"
        except:
            if argumentos['exist'] == False:
                return "saida"
                
    def finalizar_relatorio(self, argumentos):
        '''
        explicando o termo "exist"
        True: ele vai finalizar a execução se o target for encontrado
        False: ele vai finalizar a execução caso não encontre mais o target 
        '''
        try:
            if self.navegador.find_element(By.XPATH, argumentos['target']).text != '':
                raise UnboundLocalError("error ao gerar relatorio")
        except UnboundLocalError:
            raise UnboundLocalError("error ao gerar relatorio")
        except:
            pass
    
    def finalizador_de_emergencia(self, argumentos):
        try:
            target = self.navegador.find_element(By.XPATH, argumentos['target'])
            if argumentos['verific']['regra'] == "in":
                if argumentos['verific']['texto'] in target.text:
                    print(target.text)
                    return  "emergencia"
        except:
            pass

    def finalizador_controlado(self, argumentos):
        return "saida"
    
    def salvar(self, argumentos):
        try:
            target = self.navegador.find_element(By.XPATH, argumentos['target'])
            self.temp_variable = target.text
        except:
            pass

    def coletar_id_relatorio(self, argumentos):
        try:
            target = self.navegador.find_element(By.XPATH, argumentos['target'])
            self.relatorios_id.append(target.text)
        except:
            pass


if __name__ == "__main__":
    
    #pass
    sys.path.append("Entities")

    from credential.carregar_credenciais import Credential
    
    inicio_tempo = datetime.now()
    crendenciais = Credential()
    creden = crendenciais.credencial()
    
    if (creden['usuario'] == None) or (creden['senha'] == None):
        raise PermissionError("Credenciais Invalidas")


    bot = BotExtractionImobme(usuario=creden['usuario'],senha=creden['senha'],caminho_download=f"{os.getcwd()}\\downloads_samba\\")
    arquivos = bot.obter_relatorios(["imobme_controle_estoque"])
    #     #arquivos = bot.obter_relatorios(["imobme_controle_vendas", "imobme_contratos_rescindidos"])
    #     #arquivos = bot.obter_relatorios(["imobme_empreendimento"])
    #     bot.obter_relatorios(["imobme_contratos_rescindidos"])

    # except Exception as error:
    #     print(f"{type(error)} ---> {error.with_traceback()}")
    
    # print(datetime.now() - inicio_tempo)


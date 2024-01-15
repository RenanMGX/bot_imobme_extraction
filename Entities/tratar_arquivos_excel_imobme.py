import sys
sys.path.append("Entities")
import pandas as pd
import xlwings as xw
import os
from datetime import datetime
from shutil import copy2
from time import sleep
from registro.registro import Registro

class ImobmeExceltoConvert():
    def __init__(self):
        self.__exten_excel = ['.xlsx','.xlsm','.xlsb', '.xltx']
        self.registrar_error = Registro("tratar_arquivos_excel_imobme")

    def tratar_arquivos(
            self,
            lista_arquivos,
            tipo="json",
            path_data="dados",
            copyto = "",
            apenas_nome = True
                        ):
        if copyto != "":
            if copyto[-2:] != "\\":
                copyto += "\\"

        if not isinstance(lista_arquivos, list):
            raise KeyError("apenas lista podem ser carregadas nesta classe")
        if not os.path.exists(path_data):
            os.makedirs(path_data)
        
        for arquivo in lista_arquivos:
            try:
                try:
                    arquivo_exten = f".{arquivo.split('.')[-1:][0]}"
                except Exception as error:
                    self.registrar_error.record(f"{arquivo};{type(error)};{error}")
                    continue

                
                if arquivo_exten in self.__exten_excel:

                    file_temp = "temp_" + arquivo.split("\\")[-1]
                    file_name = arquivo.split("\\")[-1].split(".")[0]
                    if ("_" in file_name) and (apenas_nome):
                        file_name = file_name.split("_")[0]

                    app = xw.App(visible=False)
                    with app.books.open(arquivo)as wb:
                        if len(wb.sheets) > 1:
                            wb.sheets[0].delete()
                        wb.save(file_temp)
                    for app in xw.apps:
                        app.quit()

                    df = pd.read_excel(file_temp)
                    df = df.replace(float('nan'), None)

                    if tipo == "csv":
                        csv_file = path_data + "\\" + file_name + ".csv"
                        df = df.replace('–', '-')
                        df.to_csv(csv_file , sep=';', index=False, encoding='latin1', errors='ignore', decimal=',')
                        copytocsv = copyto + datetime.now().strftime('%d-%m-%Y_') + csv_file.split("\\")[-1]
                        copy2(csv_file, copytocsv)
                        self.registrar_error.record(f"{arquivo};Salvo com sucesso no caminho {copyto}",tipo="Concluido")

                    if tipo == "csv_integra_web":
                        if "Empreendimentos" in file_name:
                            df = self.tratar_df_empreendimento(df)

                        csv_file = path_data + "\\" + file_name + ".csv"
                        df = df.replace('–', '-')
                        df.to_csv(csv_file , sep=';', index=False, encoding='latin1', errors='ignore', decimal=',')
                        copytocsv = copyto + datetime.now().strftime('%d-%m-%Y_') + csv_file.split("\\")[-1]
                        copy2(csv_file, copytocsv)
                        self.registrar_error.record(f"{arquivo};Salvo com sucesso no caminho {copyto}",tipo="Concluido")

                    elif tipo == "json":
                        arquivo_json = df.to_json(orient='records', date_format='iso')
                        file_json = f"{path_data}\\{file_name}.json"
                        with open(file_json, 'w')as arqui:
                            arqui.write(arquivo_json)
                        copytojson = copyto  + file_json.split("\\")[-1]
                        copy2(file_json, copytojson)
                        self.registrar_error.record(f"{arquivo};Salvo com sucesso no caminho {copyto}",tipo="Concluido")


                    os.unlink(file_temp)
                    
    
            except FileNotFoundError:
                print(f"{arquivo} --> Não encontrado")
                self.registrar_error.record(f"{arquivo};Arquivo não encontrado")
                continue
            except Exception as error:
                print(f"{arquivo} : {type(error)} ---> {error}")
                self.registrar_error.record(f"{arquivo};{type(error)};{error}")
            
            
        data_files = os.listdir(path_data)
        for data in data_files:
            os.unlink(path_data + "\\" + data)
        os.rmdir(path_data)

    def tratar_df_empreendimento(self, df):
        colunas_para_remover = [
            'Área Terreno',
            'Área Construída',
            'Valor Reconstrução',
            'Matrícula',
            'Registro',
            'PEP Empreendimento',
            'Número De Unidades',
            'Data Da Opção da Planta Fase',
            'Andar Início',
            'Andar Fim',
            'Data Opção Planta Bloco',
            'Data Habite-se Bloco',
            'Número Do Andar',
            'Código Do Final',
            'Número De Dormitórios',
            'Número De Banheiro',
            'PEP Personalização'
        ]
        colunas = df.columns.to_list()
        for remover in colunas_para_remover:
            colunas.pop(colunas.index(remover))
        
        return df[colunas]


if __name__ == "__main__":
    #pass    
    tratar = ImobmeExceltoConvert()

    #tratar.tratar_arquivos(["eqwdsa\\aboteste.xlsx", "te\\teste.xlsx"],path_data="dados_samba", copyto=r"C:\Users\renan.oliveira\Downloads")
    df = pd.read_excel("downloads_IntegracaoWeb\\Empreendimentos_23891_20240115-114346.xlsx", sheet_name="IMOBME - Empreendimento")
    print(df)
    #tratar.tratar_df_empreendimento(df)

    # lista_arquivos = os.listdir("dados")
    # cont = 0
    # for arquiv in lista_arquivos:
    #     arquiv_tratado = f"{arquiv.split('_')[0].split('.')[0]}.xlsx"
    #     os.rename(f"dados\\{arquiv}", f"dados\\{arquiv_tratado}")

    #     lista_arquivos[cont] = f"dados\\{arquiv_tratado}"
    #     cont += 1

    # arquivos = tratar.tratar_arquivos(lista_arquivos)

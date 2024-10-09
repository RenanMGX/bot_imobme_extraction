import os
#from registro.registro import Registro # type: ignore
from dependencies.logs import Logs
from dependencies.functions import P
import pandas as pd
import xlwings as xw # type: ignore
from copy import deepcopy
from time import sleep
from typing import Dict
from datetime import datetime
import traceback
from getpass import getuser

class ImobmeExceltoConvert():
    def __init__(self, path:str) -> None:
        """Metodo construtor da class

        Args:
            path (str): caminho de onde os arquivos estão
        """
        self.__files_path = path
        self.__files_path_temp = self.__files_path + "temp_files\\"
        print(P("Iniciando conversão de relatorios"))
        
    
    def _verificFiles(self) -> tuple:
        """verifica todos os arquivos da pasta e salva em uma tupla apenas os arquivos excel

        Returns:
            tuple: tuple com os arquivos excel encontrados
        """
        valid_exten: list = ['.xlsx','.xlsm','.xlsb', '.xltx']
        files_list:list = []
        
        files: list = os.listdir(self.__files_path)
        
        for file in files:
            for exten in valid_exten:
                if exten in file:
                    files_list.append(self.__files_path + file)
                    
        return tuple(files_list)
        
    
    def _extract_infor(self) -> Dict[str,pd.DataFrame|None]:
        """lê os arquivos Excel e extrai em um dataframe os dados da segunda coluna de cada arquivo

        Returns:
            Dict[str,pd.DataFrame]: dicionario com [caminho do arquivo, dataframe do excel]
        """
        files_dict: dict = {}
        for file in self._verificFiles():
            try:
                app = xw.App(visible=False)
                with app.books.open(file) as wb:
                    if len(wb.sheet_names) > 1:
                        sheet = wb.sheets[1]
                    else:
                        app.kill()
                        continue
                    
                    df = sheet['A1'].expand().options(pd.DataFrame, index=False, header=True).value
                    
                    files_dict[file] = df
                app.kill()
            except Exception as error:
                Logs().register(status='Report', description=str(error), exception=traceback.format_exc())
                files_dict[file] = None
            
        return files_dict
    
    def extract_json(self, copyto:str) -> bool:
        """salva os dados do dataframe em um arquivo json

        Args:
            copyto (str): caminho onde vai salvar o arquivo json
        """
        print(P(f"tranformando arquivos em json e salvando em {copyto}"))
        if copyto[-1] != "\\":
            copyto += "\\"
                    
        #relacao:Dict[str,pd.DataFrame|None] = self._extract_infor()
                    
        for path,df in self._extract_infor().items():
            if df is None:
                continue
            file_name = path.split('\\')[-1].split('_')[0]
            json_file = df.to_json(orient='records', date_format='iso')
            with open(((copyto + file_name) + ".json"), 'w')as _file:
                _file.write(json_file)
                
        return True
                
    def extract_csv(self, copyto:str) -> bool:
        """salva os dados do dataframe em um arquivo csv

        Args:
            copyto (str): caminho onde será salvo o arquivo csv
        """
        print(P(f"tranformando arquivos em csv e salvando em {copyto}"))
        if copyto[-1] != "\\":
            copyto += "\\"
                    
        for path,df in self._extract_infor().items():
            if df is None:
                continue
            file_name = path.split('\\')[-1].split('_')[0]
            df.to_csv(((copyto + file_name) + ".csv") , sep=';', index=False, encoding='latin1', errors='ignore', decimal=',')
            
        return True

    def extract_csv_integraWeb(self, copyto:str) -> bool:
        """funciona do mesmo jeito que o self.extract_csv porem contem algumas regras de tratativas de dados

        Args:
            copyto (str): caminho onde será salvo
        """
        if copyto[-1] != "\\":
            copyto += "\\"
                    
        for path,df in self._extract_infor().items():
            if df is None:
                continue
            file_name = path.split('\\')[-1].split('_')[0]
            
            if "Empreendimentos" in file_name:
                df = self.__integraWeb_empreendimentos_filtros(df)
            elif "DadosContrato" in file_name:
                df = self.__integraWeb_dadoscontrato_filtros(df)
            
            df.to_csv(((copyto + datetime.now().strftime('%d-%m-%Y_') + file_name) + ".csv") , sep=';', index=False, encoding='latin1', errors='ignore', decimal=',')
        
        return True
    
    def __integraWeb_empreendimentos_filtros(self, df:pd.DataFrame) -> pd.DataFrame:
        return TratamentoDF(df
                ).columns_to_remove(columns_to_remove=[
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
                ]).rows_to_remove(column='Status Da Unidade', value_in_rows=[
                    'Quitada',
                    #'Vendida',
                    #'Bloqueada',
                    'Permuta',
                    #'Análise de Crédito / Risco',
                    #'Contratos - Validação',
                    #'Em Efetivação',
                    #'Disponível'
                ]).rows_to_keep(column='Nome Do Empreendimento', value_in_rows=[
                    "Novolar Moinho",
                    "Novolar Atlanta",
                    "Novolar Jardins do Brito",
                    "Novolar Green Life"
                ]).df
                
    def __integraWeb_dadoscontrato_filtros(self, df: pd.DataFrame) -> pd.DataFrame:
        return TratamentoDF(df
                ).rows_to_keep(column='Tipo de Contrato', value_in_rows=[
                    'PCV',
                    'Cessão'
                ]).rows_to_keep(column='Status', value_in_rows=[
                    'Ativo',
                    'Quitado'
                ]).rows_to_keep(column='Empreendimento', value_in_rows=[
                    "Novolar Moinho",
                    "Novolar Atlanta",
                    "Novolar Jardins do Brito",
                    "Novolar Green Life"
                ]).df    
    
class TratamentoDF:
    def __init__(self, df: pd.DataFrame) -> None:
        self.df:pd.DataFrame = df            
    
    def columns_to_remove(self, *, columns_to_remove:list):
        columns = self.df.columns.to_list()
        for column in columns_to_remove:
            columns.pop(columns.index(column))
        self.df = self.df[columns]
        return self
    
    def rows_to_keep(self, *, column:str, value_in_rows:list):
        df_temp:pd.DataFrame = pd.DataFrame()
        for rows in value_in_rows:
            df_temp = pd.concat([df_temp, self.df[self.df[column] == rows]], ignore_index=True)
        self.df = df_temp
        return self
    
    def rows_to_remove(self, *, column:str, value_in_rows:list):
        for rows in value_in_rows:
            self.df = self.df[self.df[column] != rows]
        return self

if __name__ == "__main__":
    down_path = f"{os.getcwd()}\\downloads_samba\\"
    
    print(ImobmeExceltoConvert(path=down_path).extract_csv_integraWeb(f"C:\\Users\\{getuser()}\\Downloads"))
    
        
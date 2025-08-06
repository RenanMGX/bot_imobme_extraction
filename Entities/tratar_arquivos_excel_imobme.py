import os
#from registro.registro import Registro # type: ignore
from dependencies.logs import Logs
from dependencies.functions import P
import pandas as pd
import xlwings as xw # type: ignore
from copy import deepcopy
from time import sleep
from typing import Dict, Literal
from datetime import datetime
import traceback
from getpass import getuser
from registros import Registros

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
                    files_list.append(os.path.join(self.__files_path , file))
                    
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
            file_target_path = ((copyto + file_name) + ".json")
            if os.path.exists(file_target_path):
                try:
                    os.unlink(file_target_path)
                except Exception as error:
                    pass
            with open(file_target_path, 'w')as _file:
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

    def extract_csv_integraWeb(self, copyto:str, *, tag_test:Literal["", "_Teste_SP", "_Teste_MG"]="", empreendimentos:list=[]) -> bool:
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
                if empreendimentos:
                    df = TratamentoDF(df)
                    df.rows_to_keep(column='Nome Do Empreendimento', value_in_rows=empreendimentos)
                    df = df.df
            elif "DadosContrato" in file_name:
                df = self.__integraWeb_dadoscontrato_filtros(df)
                if empreendimentos:
                    df = TratamentoDF(df)
                    df.rows_to_keep(column='Empreendimento', value_in_rows=empreendimentos)
                    df = df.df
            
            df.to_csv(os.path.join(os.path.normpath(copyto), datetime.now().strftime(f'%d-%m-%Y_{file_name}{tag_test}.csv')) , sep=';', index=False, encoding='latin1', errors='ignore', decimal=',')
        
        return True
    
    
    
    def __integraWeb_empreendimentos_filtros(self, df:pd.DataFrame) -> pd.DataFrame:
        r = Registros('integraWeb_filtros_empreendimentos')
        lista_filtros = r.load_all()
        
        for key, value in lista_filtros.items():
            if value['type'] == 'columnsToRemove':
                df = TratamentoDF(df).columns_to_remove(columns_to_remove=value['columns']).df
            elif value['type'] == 'columnsToKeep':
                df = TratamentoDF(df).columns_to_keep(columns_to_keep=value['columns']).df
            elif value['type'] == 'rowsToKeep':
                df = TratamentoDF(df).rows_to_keep(column=value['column'], value_in_rows=value['value_in_rows']).df
            elif value['type'] == 'rowsToRemove':
                df = TratamentoDF(df).rows_to_remove(column=value['column'], value_in_rows=value['value_in_rows']).df
            elif value['type'] == 'rowsToRemoveByInclude':
                df = TratamentoDF(df).rows_to_remove_include(column=value['column'], text_include=value['text_include']).df

        return df
                
    def __integraWeb_dadoscontrato_filtros(self, df: pd.DataFrame) -> pd.DataFrame:
        r = Registros('integraWeb_filtros_dadoscontrato')
        lista_filtros = r.load_all()
        
        for key, value in lista_filtros.items():
            if value['type'] == 'columnsToRemove':
                df = TratamentoDF(df).columns_to_remove(columns_to_remove=value['columns']).df
            elif value['type'] == 'columnsToKeep':
                df = TratamentoDF(df).columns_to_keep(columns_to_keep=value['columns']).df
            elif value['type'] == 'rowsToKeep':
                df = TratamentoDF(df).rows_to_keep(column=value['column'], value_in_rows=value['value_in_rows']).df
            elif value['type'] == 'rowsToRemove':
                df = TratamentoDF(df).rows_to_remove(column=value['column'], value_in_rows=value['value_in_rows']).df
            elif value['type'] == 'rowsToRemoveByInclude':
                df = TratamentoDF(df).rows_to_remove_include(column=value['column'], text_include=value['text_include']).df
        
        return df
    
class TratamentoDF:
    def __init__(self, df: pd.DataFrame) -> None:
        self.df:pd.DataFrame = df            
    
    def columns_to_remove(self, *, columns_to_remove:list):
        columns = self.df.columns.to_list()
        for column in columns_to_remove:
            try:
                columns.pop(columns.index(column))
            except:
                pass
        self.df = self.df[columns]
        return self
    
    def columns_to_keep(self, *, columns_to_keep:list):
        columns_all = self.df.columns.to_list()
        new_columns = []
        for column in columns_to_keep:
            if column in columns_all:
                new_columns.append(column)
        
        self.df = self.df[new_columns]
        return self
    
    def rows_to_keep(self, *, column:str, value_in_rows:list):
        if column in self.df.columns:
            df_temp:pd.DataFrame = pd.DataFrame()
            for rows in value_in_rows:
                try:
                    df_temp = pd.concat([df_temp, self.df[self.df[column] == rows]], ignore_index=True)
                except:
                    pass
            self.df = df_temp
        return self
    
    def rows_to_remove(self, *, column:str, value_in_rows:list):
        if column in self.df.columns:
            for rows in value_in_rows:
                try:
                    self.df = self.df[self.df[column] != rows]
                except:
                    pass
        return self
    
    def rows_to_remove_include(self, *, column:str, text_include:str):
        if column in self.df.columns:
            self.df = self.df[
                ~self.df[column].str.contains(text_include, case=False)
            ]
        return self

if __name__ == "__main__":
    down_path = f"{os.getcwd()}\\downloads_samba\\"
    
    print(ImobmeExceltoConvert(path=down_path).extract_csv_integraWeb(f"C:\\Users\\{getuser()}\\Downloads"))
    
        
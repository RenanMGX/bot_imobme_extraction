import os
from registro.registro import Registro # type: ignore
import pandas as pd
import xlwings as xw # type: ignore
from copy import deepcopy
from time import sleep
from typing import Dict
from datetime import datetime

class ImobmeExceltoConvert():
    def __init__(self, path:str) -> None:
        """Metodo construtor da class

        Args:
            path (str): caminho de onde os arquivos estão
        """
        self.__error_log = Registro("tratar_arquivos_excel_imobme")
        self.__files_path = path
        self.__files_path_temp = self.__files_path + "temp_files\\"
        
    
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
        
    
    def _extract_infor(self) -> Dict[str,pd.DataFrame]:
        """lê os arquivos Excel e extrai em um dataframe os dados da segunda coluna de cada arquivo

        Returns:
            Dict[str,pd.DataFrame]: dicionario com [caminho do arquivo, dataframe do excel]
        """
        files_dict: dict = {}
        for file in self._verificFiles():
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
            
        return files_dict
    
    def extract_json(self, copyto:str) -> None:
        """salva os dados do dataframe em um arquivo json

        Args:
            copyto (str): caminho onde vai salvar o arquivo json
        """
        if copyto[-1] != "\\":
            copyto += "\\"
                    
        for path,df in self._extract_infor().items():
            file_name = path.split('\\')[-1].split('_')[0]
            json_file = df.to_json(orient='records', date_format='iso')
            with open(((copyto + file_name) + ".json"), 'w')as _file:
                _file.write(json_file)
                
        return True
                
    def extract_csv(self, copyto:str) -> None:
        """salva os dados do dataframe em um arquivo csv

        Args:
            copyto (str): caminho onde será salvo o arquivo csv
        """
        if copyto[-1] != "\\":
            copyto += "\\"
                    
        for path,df in self._extract_infor().items():
            file_name = path.split('\\')[-1].split('_')[0]
            df.to_csv(((copyto + file_name) + ".csv") , sep=';', index=False, encoding='latin1', errors='ignore', decimal=',')
            
        return True

    def extract_csv_integraWeb(self, copyto:str) -> None:
        """funciona do mesmo jeito que o self.extract_csv porem contem algumas regras de tratativas de dados

        Args:
            copyto (str): caminho onde será salvo
        """
        if copyto[-1] != "\\":
            copyto += "\\"
                    
        for path,df in self._extract_infor().items():
            file_name = path.split('\\')[-1].split('_')[0]
            
            if "Empreendimentos" in file_name:
                df = self.integraWeb_empreendimentos(df)
            
            df.to_csv(((copyto + datetime.now().strftime('%d-%m-%Y_') + file_name) + ".csv") , sep=';', index=False, encoding='latin1', errors='ignore', decimal=',')
        
        return True
            
    def integraWeb_empreendimentos(self, df:pd.DataFrame) -> pd.DataFrame:
        columns_to_remove: list = [
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
        columns = df.columns.to_list()
        for remove in columns_to_remove:
            columns.pop(columns.index(remove))
        df = df[columns]
        
        #remove rows of column 'Nome Do Empreendimento'
        rows_to_remove: list = [
            'Acqua Galleria Condomínio Resort - Condomínio 1',
            'Acqua Galleria Condomínio Resort - Condomínio 2',
            'Acqua Galleria Condomínio Resort - Condomínio 3',
            'Edifício Adelaide Santiago',
            'Edifício Adelaide Santiago - Avulsos',
            'Edifício Apogée - Avulsos',
            'Edifício Avignon - Avulsos',
            'Edifício Brooklyn',
            'Edifício Brooklyn - Avulsos',
            'Edifício Gioia Del Colle',
            'Edifício Jornalista Oswaldo Nobre',
            'Edifício Key Biscayne',
            "Edifício L'Essence - Avulsos",
            'Edifício Maura Valadares Gontijo',
            'Edifício Mayfair Offices',
            'Edifício Nashville',
            'Edifício Neuchâtel',
            'Edifício Neuchâtel - Avulsos',
            'Edifício Niagara Falls - Edifício Angel Falls - Edifício Victoria Falls',
            'Edifício Olga Chiari',
            'Edifício Professor Danilo Ambrósio',
            'Edifício Saint Emilion',
            'Edifício Saint Tropez',
            'Edifício Saint Tropez - Avulsos',
            'Edifício Soho Square',
            'Edifício Tribeca Square',
            'Edifício Vivaldi Moreira [Holiday Inn]',
            'Four Seasons Condomínio Resort',
            'Greenport Residences',
            'Greenwich Village',
            'Manhattan Square',
            'Manhattan Square - Avulsos',
            'Mia Felicitá Condomínio',
            'Olga Gutierrez - Avulsos',
            'Palo Alto Residences',
            'Palo Alto Residences - Avulsos',
            'Park Residence Condomínio Resort',
            'Park Residence Condomínio Resort - Avulsos',
            'Priorato Residence',
            'Quintas do Morro',
            'Residencial Porto Fino',
            'Residencial Ruth Silveira e Ruth Silveira Stores',
            'Residencial Springfield',
            'The Plaza',
            'The Plaza - Avulsos',
            'Union Square',
            'Unique - Avulsos',
            'Villaggio Gutierrez',
            'Villaggio Gutierrez - Avulsos'
        ]
        for rows in rows_to_remove:
            df = df[df['Nome Do Empreendimento'] != rows]
        ##
            
        #remove rows of column 'Status Da Unidade'
        rows_to_remove = [
            'Quitada',
            #'Vendida',
            #'Bloqueada',
            'Permuta',
            #'Análise de Crédito / Risco',
            #'Contratos - Validação',
            #'Em Efetivação',
            #'Disponível'
        ]
        for rows in rows_to_remove:
            df = df[df['Status Da Unidade'] != rows]
        ##
        
        return df

if __name__ == "__main__":
    down_path = f"{os.getcwd()}\\downloads_samba\\"
    
    print(ImobmeExceltoConvert(path=down_path).extract_csv_integraWeb(r"C:\Users\renan.oliveira\Downloads"))
    
        
import sys
sys.path.append('Entities')
import os
import pandas as pd
import xlwings as xw
from tkinter import filedialog
from datetime import datetime

def verificar_arquivo(caminho):
    exten_excel = ['.xlsx','.xlsm','.xlsb', '.xltx']
    for exten in exten_excel:
        if exten in caminho:
            return caminho
    
    raise TypeError("apenas arquivos Excel")


if __name__ == "__main__":
    pasta_temp = "IntComercial_temp\\"
    try:
        caminho_arquivo = verificar_arquivo(filedialog.askopenfilename())
        print(caminho_arquivo)

        df = pd.read_excel(caminho_arquivo, sheet_name='Base Estoque', decimal=',')
        df = df.replace(float("nan"), "").replace(" ", "")

        
        df["Área privativa"] = df["Área privativa"].astype(str).str.replace(".", ",")

        df["Preço"] = df["Preço"].replace("", float(0))
        df["Preço"] = df["Preço"].round(2)#.astype(str).str.replace(".", ",")
        #df["Preço"] = df["Preço"].replace(float(-99999), "")

        df["Reajuste"] = df["Reajuste"].replace("", float(0))
        df["Reajuste"] = df["Reajuste"].round(4)
        #df["Reajuste"] = df["Reajuste"].astype(float)
        #df = df[['Referência Tabela', 'Empreendimento', 'PEP', 'Torre/Bloco', 'Unidade/Casa', 'Preço']]

        df.to_csv(f"C:\\Users\\renan.oliveira\\Downloads\\{datetime.now().strftime('%d-%m-%Y_IntComercial.csv')}", sep=";", index=False, encoding='latin1', errors='ignore', decimal=',')



    
    except Exception as error:
        print(f"{type(error)} - {error}")
    
    finally:
        pass

        #input()


    

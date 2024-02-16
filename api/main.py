import pandas as pd
import json
from time import sleep
from datetime import datetime
from fastapi import FastAPI , File, UploadFile, Form
from fastapi.responses import HTMLResponse
from starlette.requests import Request
import uvicorn
from getpass import getuser
from io import BytesIO

def verificar_arquivo(caminho):
    exten_excel = ['.xlsx','.xlsm','.xlsb', '.xltx']
    for exten in exten_excel:
        if exten in caminho:
            return caminho
    
    raise TypeError("apenas arquivos Excel")

def tratar_arquivo_int(file, cotyto):
    #caminho_arquivo = verificar_arquivo(filedialog.askopenfilename())
    #print(caminho_arquivo)
    
    

    
    
    df = pd.read_excel(BytesIO(file), sheet_name='Base Estoque', decimal=',')
    df = df.replace(float("nan"), "").replace(" ", "")
        
    df["Área privativa"] = df["Área privativa"].astype(str).str.replace(".", ",")

    df["Preço"] = df["Preço"].replace("", float(0))
    df["Preço"] = df["Preço"].round(2)#.astype(str).str.replace(".", ",")
    #df["Preço"] = df["Preço"].replace(float(-99999), "")

    df["Reajuste"] = df["Reajuste"].replace("", float(0))
    df["Reajuste"] = df["Reajuste"].round(4)
    #df["Reajuste"] = df["Reajuste"].astype(float)
    #df = df[['Referência Tabela', 'Empreendimento', 'PEP', 'Torre/Bloco', 'Unidade/Casa', 'Preço']]

    df.to_csv(cotyto, sep=";", index=False, encoding='latin1', errors='ignore', decimal=',')

variaveis = {
    "data" : datetime.now,
}

def variables(f):
    def decorador(*args, **kargs):
        response = f(*args, **kargs)
        
        for key,value in variaveis.items():
            try:
                response = response.replace("{{"+ key +"}}", str(value()))
            except TypeError:
                response = response.replace("{{"+ key +"}}", str(value))
        
        return response
    return decorador

@variables
def load_page():
    with open("page_html/index.html", 'r', encoding="utf-8")as _file:
        page = _file.read()
    return page
    

app = FastAPI()

@app.get("/", response_class=HTMLResponse)
async def main():
    
    return load_page()


@app.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file:  UploadFile = File(...)):
    conteudo = await file.read()
    file_name = file.filename
    #print(f"{file_name}")
    if conteudo:
        tratar_arquivo_int(conteudo ,cotyto=f"\\\\server008\\G\\ARQ_PATRIMAR\\WORK\\TI - RPA\\API SGA\\{datetime.now().strftime('%d-%m-%Y_IntComercial.csv')}")
    
    return load_page()


if __name__ == "__main__":
    with open('api_config.json', 'r')as _file:
        config = json.load(_file)
    uvicorn.run(app, port=config["port"], host=config["host"])
    
    

from datetime import datetime
from time import sleep
from typing import Literal
import re

class Registro():
    def __init__(self,script):
        self.__file_name = "registro.csv"
        self.__script = script
 
    def record(self, text:str, tipo:Literal['Error', 'Concluido']="Error"):
        text = re.sub(r'\n', ' <br> ', text)
        for x in range(10):
            try:
                with open(self.__file_name, 'a')as arqui:
                    arqui.write(f"{datetime.now().isoformat()};{self.__script};{tipo};{text}\n")
                break
            except PermissionError:
                print(f"feche a planilha {self.__file_name}")
            sleep(1)

if __name__ == "__main__":
    pass
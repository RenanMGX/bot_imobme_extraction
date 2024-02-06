from datetime import datetime
from time import sleep

class Registro():
    def __init__(self,script):
        self.__file_name = "registro.csv"
        self.__script = script
 
    def record(self, text, tipo="Error"):
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
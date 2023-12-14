from datetime import datetime


class Registro():
    def __init__(self,script):
        self.__file_name = "registro.csv"
        self.__script = script
 
    def record(self, text, tipo="Error"):
        with open(self.__file_name, 'a')as arqui:
            arqui.write(f"{datetime.now().isoformat()};{self.__script};{tipo};{text}\n")

if __name__ == "__main__":
    pass
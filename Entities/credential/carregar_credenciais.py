import json
import os

class Credential():
    def __init__(self, credencial="credencial.json"):
        self.__caminho = credencial

    def credencial(self):
        default = {
            "usuario": None,
            "senha": None
        }
        try:
            with open(self.__caminho, 'r')as arqui:
                return json.load(arqui)
        except:
            with open(self.__caminho, 'w')as arqui:
                json.dump(default, arqui)
                return default

if __name__ == "__main__":
    pass

    
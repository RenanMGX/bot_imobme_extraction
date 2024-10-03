import paramiko
import os
from Entities.dependencies.config import Config
from Entities.dependencies.credenciais import Credential

class TransferenciaSFTP():
    def __init__(self, *args, **kwargs):
        
        credencias:dict = Credential(Config()['credential']['sftp']).load()
        
        self.__hostname = credencias['hostname']
        self.__port = credencias['port']
        self.__username = credencias['username']
        self.__password = credencias['password']

    def transferir(self, original_file=None, copy_file=None):
        if (not isinstance(original_file, str)) or (not isinstance(copy_file, str)):
            raise TypeError("only str type")

        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(hostname=self.__hostname, port=self.__port, username=self.__username, password=self.__password)
        sftp = client.open_sftp()
        sftp.put(original_file, copy_file)
        sftp.close()
        client.close()

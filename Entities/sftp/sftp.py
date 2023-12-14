import paramiko
import os

class TransferenciaSFTP():
    def __init__(self, credencias=None):
        
        if not isinstance(credencias, dict):
            raise TypeError("Apenas Dicionarios são permitidos nessa instancia")
        if not "hostname" in credencias:
            raise KeyError("'hostname' not found")
        elif not "port" in credencias:
            raise KeyError("'port' not found")
        elif not "username" in credencias:
            raise KeyError("'username' not found")
        elif not "password" in credencias:
            raise KeyError("'password' not found")
        
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

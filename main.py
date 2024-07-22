from ntpath import join
from os import getcwd
from Entities.extrair_rel import ExtrairSAP
from Entities.credenciais import Credential
from Entities.logs import Log
from Entities.functions import Functions, _print
from shutil import copy2
from getpass import getuser
import os

class Execute(ExtrairSAP):
    def __init__(self, *, user: str = "", password: str = "", ambiente: str = "") -> None:
        super().__init__(user=user, password=password, ambiente=ambiente)
    
    def start(self, destino:str):
        _print("Iniciando")
        destino = Functions.tratar_caminho(destino)
        
        if not os.path.exists(destino):
            raise FileNotFoundError(f"Caminho não encontrado: \n    {destino=}")        
        
        _print("Carregando Relatorios do SAP:")
        files:list = [
            self.relatorio(transacao='me5a compras'), 
            self.relatorio(transacao='me2m compras'), 
            self.relatorio(transacao='zmm009 compras'), 
            self.relatorio(transacao='zmm010 compras'),
            self.relatorio(transacao='me2n contratos'),
            self.relatorio(transacao='me5a contratos'),
            self.relatorio(transacao='zmm009 contratos'),
            self.relatorio(transacao='zmm009 contratos_zrfe'),
            self.relatorio(transacao='zmm010 contratos'),
            self.relatorio_sem_variante(transacao='mkvz contratos')
        ]
        
        self.finalizar_sap()
        
        for file in files:
            if os.path.exists(file):
                copy2(file, destino)
                _print(f"arquivo {os.path.basename(file)} copiado !")
                try:
                    os.unlink(file)
                except PermissionError:
                    if Functions.fechar_excel(file):
                        os.unlink(file)
        _print("Script Finalizado!")

if __name__ == "__main__":
    try:
        crd:dict = Credential('SAP_PRD').load()
        
        bot = Execute(user=crd['user'], password=crd['password'], ambiente=crd['ambiente'])
        
        bot.start(destino=f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\RPA - Documentos\\RPA - Dados\\Follow UP\\relatorios")
        
    except Exception as error:
        Log('main.py').register_error()
    
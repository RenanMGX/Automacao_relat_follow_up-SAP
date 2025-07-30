from Entities.extrair_rel import ExtrairSAP
from Entities.functions import Functions, _print
from shutil import copy2
from getpass import getuser
import os
import traceback
from patrimar_dependencies.sharepointfolder import SharePointFolders

class ExecuteAPP(ExtrairSAP):
    def __init__(self, *, user:str, password:str, ambiente:str) -> None:
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
            self.relatorio_sem_variante(transacao='mkvz contratos'),
            self.extrair_relatorio_base(),
            self.relatorio_datas()
        ]
        
        self.finalizar_sap()
        
        for file in files:
            try:
                if os.path.exists(file):
                    copy2(file, destino)
                    _print(f"arquivo {os.path.basename(file)} copiado !")
                    try:
                        os.unlink(file)
                    except PermissionError:
                        print(traceback.format_exc())
                        if Functions.fechar_excel(file):
                            os.unlink(file)
                else:
                    raise Exception(f"{file=} não existe!")
            except Exception as error:
                print(traceback.format_exc())
        _print("Script Finalizado com Sucesso!")

if __name__ == "__main__":
    from patrimar_dependencies.credenciais import Credential
    
    sap_crd:dict = Credential(
        path_raiz=SharePointFolders(r'RPA - Dados\CRD\.patrimar_rpa\credenciais').value,
        name_file="SAP_PRD"
    ).load()
    
    
    ExecuteAPP(
        user=sap_crd['user'], 
        password=sap_crd['password'], 
        ambiente=sap_crd['ambiente']
    ).start(
        destino=r"#material\testes"
    )
    
    
    
    # try:
    #     crd:dict = Credential(Config()['credential']['crd']).load()
        
    #     bot = Execute(user=crd['user'], password=crd['password'], ambiente=crd['ambiente'])
        
    #     bot.start(destino=Config()['paths']['destino'])
        
    #     Logs().register(status='Concluido', description="Extração Concluida!")
    # except Exception as err:
    #     Logs().register(status='Error', description='erro na extração dos relatorios do Follow UP', exception=traceback.format_exc())
    
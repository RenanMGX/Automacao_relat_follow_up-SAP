from typing import Literal
from .sap import SAPManipulation
from time import sleep
from datetime import datetime
import os
from getpass import getuser
from .functions import Functions, _print
from .logs import Log

class ExtrairSAP(SAPManipulation):
    @property
    def download_path(self) -> str:
        return self.__download_path
    
    @property
    def log(self) -> Log:
        return Log(self.__class__.__name__)
    
    @property
    def variante(self) -> dict:
        return {
                'compras': ['FOLLOW UP COMPRAS'],
                'contratos': ['FOLLOW UP CONTRATOS'],
                'contratos_zrfe': ['ZRFE FOLLOW UP CONTRATOS']
                }
    
    # @property 
    # def variante_contratos(self) -> list:
    #     return ['FOLLOW UP CONTRATOS']
    
    def __init__(self, *, user: str = "", password: str = "", ambiente: str = "", download_path:str=os.path.join(os.getcwd(), 'download_relatorios')) -> None:
        super().__init__(user=user, password=password, ambiente=ambiente)
        
        if not os.path.exists(download_path):
            os.makedirs(download_path)
        self.__download_path:str = download_path
        
        for file in os.listdir(self.download_path):
            if file.endswith('.xlsx'):
                file = os.path.join(self.download_path, file)
                try:
                    os.unlink(file)
                except PermissionError:
                    if Functions.fechar_excel(file):
                        os.unlink(file)
        
    def relatorio(self, *, 
                    transacao:Literal[
                        'me5a compras', 
                        'me2m compras', 
                        'zmm009 compras', 
                        'zmm010 compras',
                        'me2n contratos',
                        'me5a contratos',
                        'zmm009 contratos',
                        'zmm009 contratos_zrfe',
                        'zmm010 contratos',
                        ],
                    fechar_sap_no_final=False
                    ) -> str:
        
        transacao, tipo = transacao.split(' ')[0:2]#type: ignore
        
        variante:list = self.variante[tipo]
        file_name:str = f"{transacao.upper()} {tipo.upper()}.xlsx"
        caminho:str = self.download_path
        
        try:
            rel = self.__extrair_relatorio(transacao=transacao,variante=variante,caminho=caminho,file_name=file_name)
            _print(f"Relatorio '{file_name}' foi gerado e salvo!")
            return rel
        except Exception as error:
            Log(file_name).register_error()
            _print(f"erro ao gerar relatorio '{file_name}' vide pasta logs")
            return "None"

    def relatorio_sem_variante(self, *, transacao:Literal['mkvz contratos'], fechar_sap_no_final=False):
        transacao, tipo = transacao.split(' ')[0:2]#type: ignore
        
        file_name:str = f"{transacao.upper()} {tipo.upper()}.xlsx"
        caminho:str = self.download_path
        
        try:
            rel = self.__extrair_relatorio(transacao=transacao,variante=[],caminho=caminho,file_name=file_name, tem_variante=False)
            _print(f"Relatorio '{file_name}' foi gerado e salvo!")
            return rel
        except Exception as error:
            Log(file_name).register_error()
            _print(f"erro ao gerar relatorio '{file_name}' vide pasta logs")
            return "None"
        
        
    @SAPManipulation.start_SAP      
    def __extrair_relatorio(self, *, transacao:str, variante:list, caminho:str, file_name:str, tem_variante:bool=True) -> str:
        self.session.findById("wnd[0]/tbar[0]/okcd").text = f"/n {transacao}"
        self.session.findById("wnd[0]").sendVKey(0)
        #self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        
        if tem_variante:
            error:str|Exception = ""
            for var in variante:
                try:
                    self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
                    try:
                        self.session.findById("wnd[1]/usr/txtV-LOW").text = ""
                        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
                        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    except:
                        pass
                    self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
                    self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
                    self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").contextMenu()
                    self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectContextMenuItem("&FILTER")
                    self.session.findById("wnd[2]/tbar[0]/btn[2]").press()
                    self.session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
                    self.session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
                    self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = var
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
                    try:
                        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
                    except AttributeError:
                        error = Exception(f"não foi possivel localizar a variante '{str(variante)}'")
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
                        continue
                        
                    self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
                    error = ""
                except Exception as e:
                    error = e
                    
            if error != "":
                raise error
        
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press() #Executa a transação
        
        try:
            self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        except:
            try:
                self.session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[1]").select()
            except:
                self.session.findById("wnd[0]/usr/shell").pressToolbarContextButton("&MB_EXPORT")
                self.session.findById("wnd[0]/usr/shell").selectContextMenuItem("&XXL")
        
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        file_full_path = os.path.join(caminho, file_name)
        Functions.fechar_excel(file_full_path)
        return file_full_path
        #import pdb;pdb.set_trace()
    
    @SAPManipulation.start_SAP
    def __encerrar(self, fechar_sap_no_final:Literal[True]):
        pass
    
    def finalizar_sap(self):
        self.__encerrar(fechar_sap_no_final=True)
    
if __name__ == "__main__":
    pass

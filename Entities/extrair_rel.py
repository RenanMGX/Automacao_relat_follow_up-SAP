from typing import Literal
from .sap import SAPManipulation
from time import sleep
from datetime import datetime
import os
from getpass import getuser

class ExtrairSAP(SAPManipulation):
    def __init__(self, *, user: str = "", password: str = "", ambiente: str = "") -> None:
        super().__init__(user=user, password=password, ambiente=ambiente)
        
    @SAPManipulation.start_SAP
    def t_me5a(self, *, variante:Literal['FOLLOW UP COMPRAS', 'FOLLOW UP CONTRATOS'], fechar_sap_no_final=False):
        file_name = datetime.now().strftime(f"me5a_{"COMPRAS" if variante == 'FOLLOW UP COMPRAS' else "CONTRATOS" if variante == 'FOLLOW UP CONTRATOS' else "NONE"}_%d%m%Y%H%M%S.xlsx")
        caminho = f"C:\\Users\\{getuser()}\\Downloads"
        
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n me5a"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").contextMenu()
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectContextMenuItem("&FILTER")
        self.session.findById("wnd[2]/tbar[0]/btn[2]").press()
        self.session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
        self.session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = variante
        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
        try:
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        except AttributeError:
            raise Exception(f"não foi possivel localizar a variante '{variante}'")   
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        return os.path.join(caminho, file_name)
        #import pdb;pdb.set_trace()
        
    
    def listar(self, campo):
        cont = 0
        for child_object in self.session.findById(campo).Children:
            print(f"{cont}: ","ID:", child_object.Id, "| Type:", child_object.Type, "| Text:", child_object.Text)
            cont += 1
    
if __name__ == "__main__":
    pass

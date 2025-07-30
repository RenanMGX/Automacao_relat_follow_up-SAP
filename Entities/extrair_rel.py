from typing import Literal
from patrimar_dependencies.sap import SAPManipulation
import os
from .functions import Functions, _print
from datetime import datetime
import pandas as pd
import traceback
from botcity.maestro import * # type: ignore
from time import sleep

class ExtrairSAP(SAPManipulation):
    @property
    def download_path(self) -> str:
        return self.__download_path
        
    @property
    def variante(self) -> dict:
        return {
                'compras': ['FOLLOW UP COMPRAS'],
                'contratos': ['FOLLOW UP CONTRATOS'],
                'contratos_zrfe': ['ZRFE FOLLOW UP CONTRATOS']
                }
        
    @property
    def layout(self) -> dict:
        return {
                'compras': ['FOLLOW UP COMPRAS - RPA', 'FOLLOW UP - RPA'],
                'contratos': ['FOLLOW UP CONTRATOS - RPA', 'FOLLOW UP - RPA'],
                'contratos_zrfe': []
                }
    
    # @property 
    # def variante_contratos(self) -> list:
    #     return ['FOLLOW UP CONTRATOS']
    
    def __init__(self, *, user: str = "", password: str = "", ambiente: str = "", download_path:str=os.path.join(os.getcwd(), 'download_relatorios')) -> None:
        self.__maestro:BotMaestroSDK|None
        try:
            self.__maestro = BotMaestroSDK.from_sys_args()
            self.__maestro.get_execution()
        except:
            self.__maestro = None
        
        super().__init__(user=user, password=password, ambiente=ambiente, new_conection=True)
        
        if not os.path.exists(download_path):
            os.makedirs(download_path)
        self.__download_path:str = download_path
        for file in os.listdir(self.__download_path):
            file = os.path.join(self.__download_path, file)
            if os.path.isfile(file):
                try:
                    os.unlink(file)
                except:
                    pass
        
        
        for file in os.listdir(self.download_path):
            if file.endswith('.xlsx'):
                file = os.path.join(self.download_path, file)
                try:
                    os.unlink(file)
                except PermissionError:
                    sleep(2)
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
        layout:list = self.layout[tipo]
        file_name:str = f"{transacao.upper()} {tipo.upper()}.xlsx"
        caminho:str = self.download_path
        

        try:
            for x in range(5):
                erro:str|Exception = ""
                try:
                    rel = self.__extrair_relatorio(transacao=transacao,variante=variante,caminho=caminho,file_name=file_name, layout=layout)
                    erro = ""
                    break
                except Exception as error:
                    erro = error
            if erro != "":#type: ignore
                raise erro #type: ignore
                    
            _print(f"Relatorio '{file_name}' foi gerado e salvo!")
            return rel#type: ignore
        except Exception as error:
            if not self.__maestro is None:            
                self.__maestro.alert(
                    task_id=self.__maestro.get_execution().task_id,
                    title=f"erro ao gerar relatorio '{file_name}' vide pasta logs",
                    message=str(traceback.format_exc()),
                    alert_type=AlertType.ERROR
                )
            _print(f"erro ao gerar relatorio '{file_name}' vide pasta logs")
            return "None"

    def relatorio_sem_variante(self, *, transacao:Literal['mkvz contratos'], fechar_sap_no_final=False):
        transacao, tipo = transacao.split(' ')[0:2]#type: ignore
        
        file_name:str = f"{transacao.upper()} {tipo.upper()}.xlsx"
        caminho:str = self.download_path
        
        try:
            for x in range(5):
                erro = ""
                try:
                    rel = self.__extrair_relatorio(transacao=transacao,variante=[],caminho=caminho,file_name=file_name, layout=[])
                    erro = ""
                    break
                except Exception as error:
                    erro = error
            if erro != "":#type: ignore
                raise erro #type: ignore
            
            _print(f"Relatorio '{file_name}' foi gerado e salvo!")
            return rel #type: ignore
        except Exception as error:
            if not self.__maestro is None:            
                self.__maestro.alert(
                    task_id=self.__maestro.get_execution().task_id,
                    title=f"erro ao gerar relatorio '{file_name}' vide pasta logs",
                    message=str(traceback.format_exc()),
                    alert_type=AlertType.ERROR
                )
            _print(f"erro ao gerar relatorio '{file_name}' vide pasta logs")
            return "None"
        
    def relatorio_datas(self):
        agora = datetime.now()
        table = {
            "Data hora": [agora.isoformat()],
            "Data": [agora.strftime('%d/%m/%Y')],
            "Hora": [agora.strftime('%H:%M:%S.%f')]
        }
        file = os.path.join(self.download_path, "data_atualização.json")
        
        pd.DataFrame(table).to_json(file, index=False)
        return file
        
    @SAPManipulation.start_SAP      
    def __extrair_relatorio(self, *, transacao:str, variante:list, caminho:str, file_name:str, layout:list, timeout:int=3) -> str:
        self.session.findById("wnd[0]/tbar[0]/okcd").text = f"/n {transacao}"
        self.session.findById("wnd[0]").sendVKey(0)
        #self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
                
        #selecionar Variante
        if variante:
            error:str|Exception = ""
            for var in variante:
                try:
                    self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
                    try:
                        #se abrir a tela de variantes que precisa digitar nome e descrição ele vai limpar e executar para ir até a lista das variantes
                        self.session.findById("wnd[1]/usr/txtV-LOW").text = ""
                        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
                        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    except:
                        pass
                    #irá fazer uma pesquita pelo nome da variante para selecionar ela
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
                    break
                except Exception as e:
                    error = e
                            
            if error != "":
                raise error
                
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press() #Executa a transação
                
        #selecionar Layout
        if layout:
            error2:str|Exception = ""
            for lay in layout:
                try:
                    try:
                        #caso não encontre o botão para abrir layout ele vai ignorar essa parte e vai partir para gerar o excel
                        self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
                    except:
                        break
                    #irá fazer uma pesquisa de layout para selecionar ele
                    try:
                        self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").setCurrentCell(0,"TEXT")
                    except:
                        continue
                    self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
                    self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
                    self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FILTER")
                    self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "FOLLOW UP COMPRAS - RPA"
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
                    try:
                        self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
                        self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
                        error2 = ""
                        break
                    except AttributeError:
                        error = Exception(f"não foi possivel localizar a variante '{str(variante)}'")
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
                except Exception as e2:
                    error2 = e2
                    
            if error2 != "":
                raise error2
                
        #import pdb; pdb.set_trace()
                    
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
        sleep(10)
        Functions.fechar_excel(file_full_path)
        return file_full_path
        #import pdb;pdb.set_trace()
    
    @SAPManipulation.start_SAP
    def extrair_relatorio_base(self) -> str:
        caminho:str = self.download_path
        file_name:str = "BASE CENTROS.txt"
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n me5a"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").setFocus()
            self.session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").caretPosition = 0
            self.session.findById("wnd[0]").sendVKey(4)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]").sendVKey(14)
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[2]/usr/ctxtDY_PATH").text = caminho
            self.session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
            
            caminho = os.path.join(caminho, file_name)
            caminho = ExtrairSAP.tratar_base(caminho)
            _print(f"Relatorio 'relatorio_base' foi gerado e salvo!")
            return caminho
            
        except Exception as err:
            if not self.__maestro is None:            
                self.__maestro.alert(
                    task_id=self.__maestro.get_execution().task_id,
                    title=str(err),
                    message=str(traceback.format_exc()),
                    alert_type=AlertType.ERROR
                )
            
            return "None"
        
    @staticmethod
    def tratar_base(path:str) -> str:
        maestro:BotMaestroSDK|None
        try:
            maestro = BotMaestroSDK.from_sys_args()
            maestro.get_execution()
        except:
            maestro = None
        
        
        try:
            if not path.endswith('.txt'):
                try:
                    raise Exception(f"o arquivo não é .txt '{path}'")
                except Exception as err:
                    if not maestro is None:            
                        maestro.alert(
                            task_id=maestro.get_execution().task_id,
                            title=str(err),
                            message=str(traceback.format_exc()),
                            alert_type=AlertType.ERROR
                        )
                return "None"
                
            if os.path.exists(path):
                with open(path, 'r')as _file:
                    linhas:list = _file.read().split('\n')
                linhas = linhas[3:-2]
                dados = {
                    "Divisão": [],
                    "Descrição": [],
                    "CEP": [],
                    "Cidade": [],
                    "Responsável": []
                }
                for linha in linhas:
                    dados_temp = linha.split('|')
                    dados["Divisão"].append(dados_temp[1])
                    dados['Descrição'].append(dados_temp[4])
                    dados['CEP'].append(dados_temp[5])
                    dados['Cidade'].append(dados_temp[6])
                    dados['Responsável'].append(" ")
            
                new_file = path.replace('.txt', '.xlsx')
                os.unlink(path)
                pd.DataFrame(dados).to_excel(new_file, index=False)
                return new_file
            
            else:
                try:
                    raise Exception("arquivo não encontrado")
                except Exception as err:
                    if not maestro is None:            
                        maestro.alert(
                            task_id=maestro.get_execution().task_id,
                            title=str(err),
                            message=str(traceback.format_exc()),
                            alert_type=AlertType.ERROR
                        )
                return "None"
        except Exception as err:
            if not maestro is None:            
                maestro.alert(
                    task_id=maestro.get_execution().task_id,
                    title=str(err),
                    message=str(traceback.format_exc()),
                    alert_type=AlertType.ERROR
                )
            return "None"
    
    @SAPManipulation.start_SAP
    def __encerrar(self, fechar_sap_no_final:bool):
        pass
    
    def finalizar_sap(self, *, mostrar_na_tela:bool=True):
        self.__encerrar(fechar_sap_no_final=True)
        if mostrar_na_tela:
            _print("Relatorios finalizados SAP encerrado!")
    
    @SAPManipulation.start_SAP
    def test_in_sap(self):
        print()        

    
if __name__ == "__main__":
    pass

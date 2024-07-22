import os
from datetime import datetime
import traceback

def _print(*args, end="\n"):
    if not end.endswith("\n"):
        end += "\n"
    value = ""
    for arg in args:
        value += f"{arg} " 
    
    print(datetime.now().strftime(f"[%d/%m/%Y - %H:%M:%S] - {value}"), end=end)


class Log:
    @property
    def file_path(self) -> str:
        return self.__file_path
    
    def __init__(self, _name:str, *, path:str=os.path.join(os.getcwd(), ".logs")) -> None:
        name = f"-->TIME<--_Erro__{_name}__"
            
        if not os.path.exists(path):
            os.makedirs(path)
        self.__file_path:str = os.path.join(path, name)
    
    def register_error(self, *, __print:bool=False) -> None:
        _traceback:str = traceback.format_exc()
        file_path:str = self.file_path.replace("-->TIME<--", datetime.now().strftime("%Y%m%d-%H%M%S"))
        if not file_path.endswith(".txt"):
            file_path += ".txt"
        
        with open(file_path, 'w', encoding='utf-8')as _file:
            _file.write(_traceback)
        if __print:
            _print(_traceback)

if __name__ == "__main__":
    pass
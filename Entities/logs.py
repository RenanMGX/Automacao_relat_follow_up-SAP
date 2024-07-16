import os
from datetime import datetime
import traceback

class Log:
    @property
    def file_path(self) -> str:
        return self.__file_path
    def __init__(self, name:str, *, path:str=os.path.join(os.getcwd(), ".logs")) -> None:
        if not name.endswith(".txt"):
            name += ".txt"
            
        if not os.path.exists(path):
            os.makedirs(path)
        self.__file_path:str = os.path.join(path, name)
    
    def register_error(self) -> None:
        _traceback:str = traceback.format_exc()
        file_path:str = self.file_path.replace(".txt", datetime.now().strftime("_Error_%d%b%Y-%H%M%S.txt"))
        with open(file_path, 'w', encoding='utf-8')as _file:
            _file.write(_traceback)
        print(_traceback)

if __name__ == "__main__":
    pass
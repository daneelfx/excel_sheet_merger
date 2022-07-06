import os
from os import path as os_path, access, scandir

class Path:
  
  def __init__(self, path):
    self.path = path

  @property
  def path(self):
    return self._path

  @path.setter
  def path(self, path):
    if not os_path.exists(path):
      raise OSError(f'ERROR: La ruta {path} no fue encontrada')
    if not os_path.isdir(path):
      raise NotADirectoryError(f'ERROR: La ruta {path} no hace referencia a un directorio válido')
    if not access(path, mode = os.R_OK):
      raise OSError(f'ERROR: No fue posible acceder a la ruta {path}')
    self._path = path



class FilesExtractor:

  def __init__(self, path_instance):
    self.path_instance = path_instance

  @property
  def path_instance(self):
    return self._path_instance

  @path_instance.setter
  def path_instance(self, path_instance):
    if not isinstance(path_instance, Path):
      raise TypeError('El argumento ingresado no es de tipo Path')
    self._path_instance = path_instance


class ExcelFilesMerger(FilesExtractor):

  ALLOWED_EXTENSIONS = ('xlsx',)

  def __init__(self, path_instance):
    super().__init__(path_instance)

  
  def merge_sheets(self):
    

  




  


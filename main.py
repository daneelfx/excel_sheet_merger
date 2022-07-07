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



class PathContent:

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

  @property
  def content_structure(self):
    path_str = self.path_instance.path
    return PathContent._build_structure(path_str)

  @staticmethod
  def _build_structure(path_str):
    content = []
    for path in os.listdir(path_str):
      relative_path_str = f'{path_str}/{path}'
      is_dir = os_path.isdir(relative_path_str)
      last_modification_date = os_path.getmtime(relative_path_str)
      size = os_path.getsize(relative_path_str)

      path_info = {'path': relative_path_str, 'is_dir': is_dir, 'last_mod_date': last_modification_date, 'size': size}
      
      if is_dir:
        path_info.update({'content': PathContent._build_structure(relative_path_str)})

      content.append(path_info)

    return content

print(PathContent(Path('.')).content_structure)


# class ExcelFilesMerger:

#   ALLOWED_EXTENSIONS = ('xlsx',)

#   def __init__(self, path_instance):
#     super().__init__(path_instance)
  
#   def merge_sheets(self):
#     pass

#   @staticmethod
#   def recursive_search(path):
    
#     if os_path.isfile(path):
#       print(path)
#     elif os_path.isdir(path):
#       for dir_path in os.listdir(path):
#         print(dir_path)
#         ExcelFilesMerger.recursive_search(dir_path)

#   def paths(self):
#     path = self.path_instance.path

#     ExcelFilesMerger.recursive_search(path)
  
# ExcelFilesMerger(Path('./')).paths()

  




  


import os
from os import path as os_path, access, scandir
import shutil


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
      raise NotADirectoryError(f'ERROR: La ruta {path} no es un directorio')
    if not access(path, mode = os.R_OK):
      raise PermissionError(f'ERROR: No fue posible acceder a la ruta {path}')
    self._path = path



class PathContent:

  def __init__(self, path_instance):
    self.path_instance = path_instance

  @classmethod
  def create_path_content(cls):
    return cls

  @property
  def path_instance(self):
    return self._path_instance

  @path_instance.setter
  def path_instance(self, path_instance):
    if not isinstance(path_instance, Path):
      raise TypeError('El argumento ingresado no es de tipo Path')
    self._path_instance = path_instance

  @property
  def content_iterator(self):
    path_str = self.path_instance.path
    return PathContent.get_content_iterator(path_str)

  @staticmethod
  def get_content_iterator(path_str):

    for path in os.listdir(path_str):
      content = []
      relative_path_str = f'{path_str}/{path}'
      is_dir = os_path.isdir(relative_path_str)
      last_modification_date = os_path.getmtime(relative_path_str)
      size = os_path.getsize(relative_path_str)

      path_info = {'path': relative_path_str, 'is_dir': is_dir, 'last_mod_date': last_modification_date, 'size': size}
      
      if is_dir:
        path_info.update({'content': PathContent.get_content_iterator(relative_path_str)})

      content.append(path_info)

      yield path_info

  @property
  def flattened_content_iterator(self):
    return PathContent.get_flattened_content_iterator(self.content_iterator)

  @staticmethod
  def get_flattened_content_iterator(content_iterator):
    for path_item in content_iterator:
      yield path_item
      if path_item['is_dir']:
        yield from PathContent.get_flattened_content_iterator(path_item['content'])



class FileMerger:
  
  def __init__(self, source_path_instance, target_path_instance):
    self.source_path_instance = source_path_instance
    self.target_path_instance = target_path_instance

  @property
  def source_path_instance(self):
    return self._source_path_instance

  @source_path_instance.setter
  def source_path_instance(self, source_path_instance):
    if not isinstance(source_path_instance, Path):
      raise TypeError('El argumento ingresado no es de tipo Path')
    self._source_path_instance = source_path_instance

  @property
  def target_path_instance(self):
    return self._target_path_instance

  @target_path_instance.setter
  def target_path_instance(self, target_path_instance):
    if not isinstance(target_path_instance, Path):
      raise TypeError('El argumento ingresado no es de tipo Path')
    self._target_path_instance = target_path_instance

  def merge_files(self):
    raise NotImplementedError('ERROR: FileMerger no implementa este método. Use una clase que herede la clase FileMerger')



class ExcelFileMerger(PathContent, FileMerger):

  ALLOWED_FILE_EXTENSIONS = ('xlsx',)
  
  def __init__(self, source_path_instance, target_path_instance):
    super().__init__(source_path_instance)
    super(PathContent, self).__init__(self.path_instance, target_path_instance)

  def merge_files(self):
    if os.listdir(self.target_path_instance.path):
      shutil.rmtree(self.target_path_instance.path)
      os.mkdir(self.target_path_instance.path)

    self.build_content(self.content_iterator)


  def build_content(self, content_iterator):

    
    for path_item in content_iterator:
      full_output_path_str = f"{path_item['path'].replace(self.source_path_instance.path, self.target_path_instance.path)}"
      if path_item['is_dir']:
        
        os.mkdir(full_output_path_str)

        # LOGIC GOES HERE
        
        self.build_content(path_item['content'])
      else:  
        if path_item['path'].endswith('.xlsx'):
          print(path_item['path'])

  


# path = Path('./folder1')
# path_content = PathContent(path)

# print(list(path_content.flattened_content_iterator))

ExcelFileMerger(Path('./folder1'), Path('C:/Users/danee/Downloads/testfolder')).merge_files()







  




  


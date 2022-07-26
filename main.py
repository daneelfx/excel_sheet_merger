from calendar import month_name
from itertools import count
import os
from os import path as os_path, access
from datetime import datetime, date, timedelta
import logging
from sqlalchemy import except_

import xlwings as xw
import pandas as pd
import numpy as np
from pprint import pprint



class Path:
  
  def __init__(self, path):
    self.path = path

  @property
  def path(self):
    return self._path

  @path.setter
  def path(self, path):
    if not os_path.exists(path):
      logging.error(f"ERROR: La ruta '{path}'' no fue encontrada")
      raise OSError(f"ERROR: La ruta '{path}'' no fue encontrada")
    if not os_path.isdir(path):
      logging.error(f"ERROR: La ruta {path}' no es un directorio")
      raise NotADirectoryError(f"ERROR: La ruta {path}' no es un directorio")
    if not access(path, mode = os.R_OK):
      logging.error(f"ERROR: No fue posible acceder a la ruta '{path}'")
      raise PermissionError(f"ERROR: No fue posible acceder a la ruta '{path}'")
    self._path = path
    logging.info(f"La ruta '{self._path}' fue accedida correctamente")



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
      raise TypeError('ERROR: El argumento ingresado no es de tipo Path')
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
      creation_date = os_path.getctime(relative_path_str)
      size = os_path.getsize(relative_path_str)

      path_info = {'path': relative_path_str, 'is_dir': is_dir, 'creation_date': creation_date, 'last_mod_date': last_modification_date, 'size': size}
      
      if is_dir:
        path_info.update({'content': PathContent.get_content_iterator(relative_path_str)})

      content.append(path_info)

      yield path_info

  def _iterate_over_content(self, content_iterator, *, callback):
      for path_item in content_iterator:
        if path_item['is_dir']:
          self._iterate_over_content(path_item['content'], callback = callback)
        else:
          callback(path_item)

  def get_content_structure(self, *, file_extensions):
    filepaths_per_dir = {}

    def _content_builder_callback(path_item):
      nonlocal filepaths_per_dir

      if path_item['path'].split('.')[-1] in file_extensions:
        parts_reversed = path_item['path'][:: -1].split('/', 1)
        dir_path = parts_reversed[-1][:: -1]
        file_name = parts_reversed[0][:: -1]
        path_item_reduced = {'name': file_name, 'creation_date': path_item['creation_date'], 'last_mod_date': path_item['last_mod_date']}

        if dir_path in filepaths_per_dir:
          filepaths_per_dir[dir_path]['files'].append(path_item_reduced)        
        else:
          filepaths_per_dir[dir_path] = {'level': dir_path.replace(self.path_instance.path, '', 1).count('/'), 'files': [path_item_reduced]}

    self._iterate_over_content(self.content_iterator, callback = _content_builder_callback)

    return filepaths_per_dir

  def get_content_tree(self, file_extensions):
    root_path = self.path_instance.path
    level = 0

    def _build_tree(path_str):
      nonlocal level
      children = []

      for path in os.listdir(path_str):
        relative_path_str = f'{path_str}/{path}'
        is_dir = os_path.isdir(relative_path_str)
        last_modification_date = os_path.getmtime(relative_path_str)
        creation_date = os_path.getctime(relative_path_str)
        size = os_path.getsize(relative_path_str)

        path_info = {'level': relative_path_str.replace(root_path, '', 1).count('/') - 1, 'path': relative_path_str, 'is_dir': is_dir, 'creation_date': creation_date, 'last_mod_date': last_modification_date, 'size': size}
        
      
        if is_dir:
          path_info.update({'children': _build_tree(relative_path_str)})
          children.append(path_info)
        elif relative_path_str[:: -1].split('.', 1)[0][:: -1].lower() in file_extensions:
          children.append(path_info)
        

      return children
    return {'level': -1, 'path': root_path, 'children': _build_tree(root_path)}



class FileMerger:
  
  def __init__(self, source_path_instances, target_path_instance):
    self.source_path_instances = source_path_instances
    self.target_path_instance = target_path_instance

  @property
  def source_path_instances(self):
    return self._source_path_instances

  @source_path_instances.setter
  def source_path_instances(self, source_path_instances):
    for source_path_instance in source_path_instances:
      if not isinstance(source_path_instance, Path):
        raise TypeError('ERROR: El argumento ingresado no es de tipo Path')
    self._source_path_instances = source_path_instances

  @property
  def target_path_instance(self):
    return self._target_path_instance

  @target_path_instance.setter
  def target_path_instance(self, target_path_instance):
    if not isinstance(target_path_instance, Path):
      raise TypeError('ERROR: El argumento ingresado no es de tipo Path')
    self._target_path_instance = target_path_instance

  def merge_files(self):
    raise NotImplementedError('ERROR: FileMerger no implementa este método. Use una clase que herede la clase FileMerger')



class ExcelFileMerger(FileMerger):
  ALLOWED_FILE_EXTENSIONS = ('xls', 'xlsx', 'xlsm')
  MONTHS_MAPPING = {'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04', 'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08', 'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'}
  
  def __init__(self, *source_path_instances, target_path_instance):
    super().__init__(source_path_instances, target_path_instance = target_path_instance)


  def _build_dates_mapping(self):
    dates_mapping = pd.DataFrame(columns = ['type', 'source_path', 'file_name', 'target_year', 'target_month'])

    def _traverse_tree(tree_structure, struct_type):
      nonlocal dates_mapping

      for child in tree_structure['children']:
        if child['is_dir']:
          folder_name = child['path'].split('/')[-1]
          if struct_type == 'type_1' and child['level']:
            if folder_name.lower() in ExcelFileMerger.MONTHS_MAPPING:
              _traverse_tree(child, struct_type)
            else:
              logging.warning(f"El nombre de la carpeta '{child['path']}' no es el nombre de un mes (e.g. 'JUNIO')")
          else:              
            try:
              int(folder_name)

              if len(folder_name) == 4:
                _traverse_tree(child, 'type_1')

              if len(folder_name) == 6:
                _traverse_tree(child, 'type_2')

            except ValueError:
              if not child['level']:
                logging.warning(f"El nombre de la carpeta '{child['path']}' no es de la forma 'YYYY' o 'YYYYMM'")
              elif child['level'] == 1 and len(folder_name) == 4:
                logging.warning(f"El nombre de la carpeta '{child['path']}' no es de la forma 'MM'")
              

        if not child['is_dir'] and child['level']:
            path_parts = child['path'].split('/')
            source_path = '/'.join(path_parts[:-1])
            file_name = ''.join(path_parts[-1])
            
            if struct_type == 'type_1':
              year = path_parts[-(child['level'] + 1)]
              month = ExcelFileMerger.MONTHS_MAPPING[path_parts[-child['level']].lower()]
            
            if struct_type == 'type_2':
              year_month = path_parts[-(child['level'] + 1)]
              year = year_month[:4]
              month = year_month[-2:]

            dates_mapping = dates_mapping.append(pd.Series(data = [struct_type, source_path, file_name, year, month], index = ['type', 'source_path', 'file_name', 'target_year', 'target_month']), ignore_index = True)
            
    for source_path_instance in self._source_path_instances:
      path_content = PathContent(source_path_instance)
      path_content_tree = path_content.get_content_tree(file_extensions = ExcelFileMerger.ALLOWED_FILE_EXTENSIONS)   
      _traverse_tree(path_content_tree, None)

    return dates_mapping


  def merge_files(self, output_paths_dict):
    dates_mapping = self._build_dates_mapping()

    def _do_merging(dataframe):
      business_code = pd.unique(dataframe['business_code'])[0]

      if not business_code in output_paths_dict:
        logging.error(f"El negocio '{business_code}' no se encuentra especificado en el archivo 'rutas_negocios.xlsx'")
        return

      app = xw.App(visible = True)
      current_output_book = app.books.add()
      current_output_book.sheets[0].name = 'SHEET_TO_BE_DELETED'

      for _, file_props in dataframe.iterrows():
        source_filepath = f"{file_props['source_path']}/{file_props['file_name']}"

        try:
          current_input_book = xw.Book(source_filepath)

          for sheet in iter(current_input_book.sheets):
            sheet.copy(after = current_output_book.sheets[current_output_book.sheets.count - 1])

          logging.info(f"El archivo '{source_filepath}' fue leido correctamente y sera usado para consolidacion. Espere confirmacion")
        except:
          app.quit()
          logging.error(f"El archivo '{source_filepath}' no pudo ser leido")

      current_output_book.sheets[0].delete()

      months_mapping_reversed = {month_number: month_name.upper() for month_name, month_number in ExcelFileMerger.MONTHS_MAPPING.items()}
      business_props = output_paths_dict[business_code]
      year = pd.unique(dataframe['target_year'])[0]
      month = pd.unique(dataframe['target_month'])[0]

      output_dir = f"{self.target_path_instance.path}/{business_props['group']}/{business_code} - {business_props['name']}/{'/'.join(business_props['sublevels'])}/{year}/{month} - {months_mapping_reversed[month]}"

      try:
        os.makedirs(output_dir, exist_ok = True)
      except:
        logging.critical(f"No fue posible crear la carpeta '{output_dir}'")
        app.quit()
        return

      try:
        output_file_path = f"{output_dir}/{pd.unique(dataframe['business_code'])[0]}.xlsx"
        current_output_book.save(output_file_path)
        logging.info(f"El archivo '{output_file_path}' fue consolidado y guardado")
      except:
        logging.critical(f"No fue posible guardar el archivo '{output_file_path}'")

      current_output_book.app.quit()
         
    type_1_paths = dates_mapping[dates_mapping['type'] == 'type_1']
    type_1_paths = type_1_paths[type_1_paths['file_name'].apply(lambda path: len(path.split('-')) == 2)]
    type_1_paths['business_code'] = type_1_paths['file_name'].apply(lambda file_name: file_name.split('-')[0].strip())
    type_1_paths['paper_name'] = type_1_paths['file_name'].apply(lambda file_name: file_name.split('-')[1].split('.')[0].strip())
    type_1_paths = type_1_paths.drop_duplicates(subset = ['target_year', 'target_month', 'business_code', 'paper_name'], keep = 'last')

    type_2_paths = dates_mapping[dates_mapping['type'] == 'type_2']
    type_2_paths = type_2_paths[type_2_paths['file_name'].apply(lambda path: len(path.split('_')) == 5)]
    type_2_paths['business_code'] = type_2_paths['file_name'].apply(lambda file_name: file_name.split('_')[0].strip())
    type_2_paths['paper_name'] = type_2_paths['file_name'].apply(lambda file_name: file_name.split('_')[1].strip())
    type_2_paths = type_2_paths.drop_duplicates(subset = ['target_year', 'target_month', 'business_code', 'paper_name'], keep = 'last')

    both_types_paths = type_1_paths.append(type_2_paths)
    both_types_paths.groupby(by = ['target_year', 'target_month', 'business_code']).apply(_do_merging)


if __name__ == '__main__':

  logging.basicConfig(filename = 'log.txt', format = '%(levelname)s %(asctime)s %(message)s', level = logging.INFO)
  logging.info(f"{'*' * 8} NUEVA EJECUCION INICIALIZADA {'*' * 128}")

  try:
    inputs_output_excel = pd.read_excel('./entradas_salida.xlsx').dropna().apply(lambda row: row.astype(str).str.strip())
    counts = inputs_output_excel.groupby('TIPO').count()
    if inputs_output_excel.columns.tolist() != ['TIPO', 'RUTA'] or (counts.loc['salida'] != 1).bool() or (counts.loc['entrada'] < 1).bool():
      raise AttributeError
    logging.info("El archivo 'entradas_salida.xlsx' fue leido correctamente")
    inputs_output_excel['RUTA'] = inputs_output_excel['RUTA'].apply(lambda row: row.replace('\\', '/'))
  except:
    logging.critical("El archivo 'entradas_salida.xlsx' no existe, presenta un formato incorrecto o no tiene al menos una ruta de entrada y una de salida")
    logging.info(f"{'*' * 8} EJECUCION FINALIZADA {'*' * 128}")
    raise

  try:
    output_paths_dict = {}
    output_paths_excel = pd.read_excel('./rutas_negocios.xlsx')
    output_paths_excel = output_paths_excel.iloc[output_paths_excel[output_paths_excel.columns[:3]].dropna().index]
    output_paths_excel = output_paths_excel.fillna('').apply(lambda row: row.astype(str).str.strip()).drop_duplicates(subset = ['CODIGO'], keep = 'last')

    if output_paths_excel.columns.tolist() != ['CODIGO', 'NOMBRE', 'GRUPO', 'SUBNIVEL 1', 'SUBNIVEL 2', 'SUBNIVEL 3', 'SUBNIVEL 4', 'SUBNIVEL 5']:
      raise AttributeError
    for _, row in output_paths_excel.iterrows():
      output_paths_dict[row['CODIGO']] = {'name': row['NOMBRE'], 'group': row['GRUPO'], 'sublevels': list(filter(lambda column_name: len(column_name), row.to_list()[-5:]))}
  except:
    logging.critical(f"El archivo 'rutas_negocios.xlsx' no existe o presenta un formato incorrecto")
    raise

  output_path_instance = Path(inputs_output_excel[inputs_output_excel['TIPO'] == 'salida']['RUTA'].iloc[0])
  input_path_instances = tuple(Path(input_path) for _, input_path in inputs_output_excel[inputs_output_excel['TIPO'] == 'entrada']['RUTA'].iteritems())

  file_merger = ExcelFileMerger(*input_path_instances, target_path_instance = output_path_instance)
  file_merger.merge_files(output_paths_dict)

  logging.info(f"{'*' * 8} EJECUCION FINALIZADA {'*' * 128}")
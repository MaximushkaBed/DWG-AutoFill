import pandas as pd
import ezdxf
import os
import json
from typing import Dict, Any

class IOManager:
    """
    Модуль для надёжной загрузки/сохранения DWG/DXF и табличных данных (Excel, CSV, JSON).
    """

    def load_dwg(self, path: str) -> ezdxf.document.Drawing:
        """
        Загружает DWG/DXF файл.
        """
        if not os.path.exists(path):
            raise FileNotFoundError(f"Файл не найден: {path}")
        
        try:
            # ezdxf.readfile() поддерживает как DXF, так и DWG (через встроенный конвертер)
            doc = ezdxf.readfile(path)
            return doc
        except ezdxf.DXFStructureError as e:
            raise IOError(f"Ошибка структуры DXF/DWG файла: {e}")
        except Exception as e:
            raise IOError(f"Не удалось прочитать DWG/DXF файл: {e}")

    def read_table(self, path: str) -> pd.DataFrame:
        """
        Читает данные из Excel, CSV или JSON в pandas DataFrame.
        """
        if not os.path.exists(path):
            raise FileNotFoundError(f"Файл данных не найден: {path}")

        path_lower = path.lower()
        
        try:
            if path_lower.endswith(('.xlsx', '.xls')):
                # Используем openpyxl для поддержки xlsx
                df = pd.read_excel(path, engine='openpyxl')
            elif path_lower.endswith('.csv'):
                # Попытка прочитать с разными разделителями и кодировками
                try:
                    df = pd.read_csv(path, encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(path, encoding='cp1251')
            elif path_lower.endswith('.json'):
                df = pd.read_json(path)
            else:
                raise ValueError('Неподдерживаемый формат файла данных. Ожидается .xlsx, .csv или .json.')
            
            # Удаляем полностью пустые строки
            df.dropna(how='all', inplace=True)
            
            if df.empty:
                raise ValueError("Файл данных пуст или содержит только заголовки.")
                
            return df
            
        except Exception as e:
            raise Exception(f"Ошибка при чтении файла данных {path}: {e}")

    def save_dwg(self, doc: ezdxf.document.Drawing, out_path: str, version: str = 'AC1021') -> None:
        """
        Сохраняет изменённый DWG/DXF документ.
        """
        try:
            doc.saveas(out_path, version=version)
        except Exception as e:
            raise IOError(f"Ошибка при сохранении DWG/DXF файла в {out_path}: {e}")

    def ensure_directory(self, path: str) -> None:
        """
        Создает директорию, если она не существует.
        """
        os.makedirs(path, exist_ok=True)

    def get_dxf_attributes(self, doc: ezdxf.document.Drawing) -> Dict[str, Any]:
        """
        Извлекает уникальные имена атрибутов из всех блоков в документе.
        Возвращает словарь {tag_name: [list_of_block_names_using_this_tag]}
        """
        attributes = {}
        for block in doc.blocks:
            # Пропускаем стандартные блоки и блоки, не содержащие атрибутов
            if block.is_layout or block.is_anonymous or not block.has_attdefs:
                continue
            
            for attdef in block.get_attdefs():
                tag = attdef.dxf.tag
                if tag not in attributes:
                    attributes[tag] = []
                attributes[tag].append(block.name)
                
        return attributes

# Создание файла __init__.py для модуля src
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'w') as f:
    f.write('from .io_manager import IOManager\n')

if __name__ == '__main__':
    # Пример использования (для отладки)
    # Для реального тестирования нужны DWG/DXF и Excel файлы
    print("IOManager создан. Для тестирования требуются файлы DWG/DXF и Excel.")
    # io_manager = IOManager()
    # try:
    #     # Пример загрузки DWG (замените на реальный путь)
    #     # doc = io_manager.load_dwg("path/to/template.dwg")
    #     # print(f"Загружен DWG: {doc.dxfversion}")
    #     
    #     # Пример чтения таблицы (замените на реальный путь)
    #     # df = io_manager.read_table("path/to/data.xlsx")
    #     # print(f"Прочитана таблица, строк: {len(df)}")
    #     
    #     # Пример получения атрибутов
    #     # attrs = io_manager.get_dxf_attributes(doc)
    #     # print(f"Найденные атрибуты: {list(attrs.keys())}")
    #     pass
    # except Exception as e:
    #     print(f"Ошибка при тестировании: {e}")
    pass

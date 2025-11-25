import pandas as pd
import ezdxf
import os
import json
from typing import Dict, Any
from .logger import logger

class IOManager:
    """
    Модуль для надёжной загрузки/сохранения DWG/DXF и табличных данных (Excel, CSV, JSON).
    Улучшенная валидация и логирование.
    """

    def load_dwg(self, path: str) -> ezdxf.document.Drawing:
        """
        Загружает DWG/DXF файл.
        """
        if not os.path.exists(path):
            logger.error(f"Файл не найден: {path}", context={'path': path})
            raise FileNotFoundError(f"Файл не найден: {path}")
        
        try:
            logger.info(f"Загрузка DWG/DXF файла: {path}")
            doc = ezdxf.readfile(path)
            logger.info(f"DWG/DXF файл успешно загружен. Версия: {doc.dxfversion}")
            return doc
        except ezdxf.DXFStructureError as e:
            logger.error(f"Ошибка структуры DXF/DWG файла: {e}", context={'path': path})
            raise IOError(f"Ошибка структуры DXF/DWG файла: {e}")
        except Exception as e:
            logger.error(f"Не удалось прочитать DWG/DXF файл: {e}", context={'path': path})
            raise IOError(f"Не удалось прочитать DWG/DXF файл: {e}")

    def read_table(self, path: str) -> pd.DataFrame:
        """
        Читает данные из Excel, CSV или JSON в pandas DataFrame.
        """
        if not os.path.exists(path):
            logger.error(f"Файл данных не найден: {path}", context={'path': path})
            raise FileNotFoundError(f"Файл данных не найден: {path}")

        path_lower = path.lower()
        
        try:
            logger.info(f"Чтение файла данных: {path}")
            if path_lower.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(path, engine='openpyxl')
            elif path_lower.endswith('.csv'):
                # Попытка прочитать с разными разделителями и кодировками
                try:
                    df = pd.read_csv(path, encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(path, encoding='cp1251')
                except Exception as e:
                    logger.warning(f"Не удалось прочитать CSV с utf-8/cp1251: {e}. Попытка с автоматическим определением.", context={'path': path})
                    df = pd.read_csv(path) # Попытка с дефолтной кодировкой
            elif path_lower.endswith('.json'):
                # JSON может быть как списком объектов, так и объектом с данными
                try:
                    df = pd.read_json(path)
                except ValueError:
                    # Попытка прочитать как список объектов, если не удалось как таблицу
                    with open(path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    df = pd.DataFrame(data)
            else:
                logger.error(f"Неподдерживаемый формат файла данных: {path}", context={'path': path})
                raise ValueError('Неподдерживаемый формат файла данных. Ожидается .xlsx, .csv или .json.')
            
            # Удаляем полностью пустые строки
            df.dropna(how='all', inplace=True)
            
            if df.empty:
                logger.warning(f"Файл данных пуст: {path}", context={'path': path})
                raise ValueError("Файл данных пуст или содержит только заголовки.")
            
            logger.info(f"Файл данных успешно прочитан. Строк: {len(df)}, Колонок: {len(df.columns)}")
            return df
            
        except Exception as e:
            logger.error(f"Ошибка при чтении файла данных: {e}", context={'path': path})
            raise Exception(f"Ошибка при чтении файла данных {path}: {e}")

    def save_dwg(self, doc: ezdxf.document.Drawing, out_path: str, version: str = 'AC1021') -> None:
        """
        Сохраняет изменённый DWG/DXF документ.
        """
        try:
            logger.info(f"Сохранение DWG/DXF в {out_path} (версия {version})", context={'path': out_path, 'version': version})
            doc.saveas(out_path, version=version)
            logger.info(f"DWG/DXF успешно сохранен: {out_path}")
        except Exception as e:
            logger.error(f"Ошибка при сохранении DWG/DXF файла: {e}", context={'path': out_path})
            raise IOError(f"Ошибка при сохранении DWG/DXF файла в {out_path}: {e}")

    def ensure_directory(self, path: str) -> None:
        """
        Создает директорию, если она не существует.
        """
        os.makedirs(path, exist_ok=True)
        logger.info(f"Проверено/создано: {path}")

    def get_dxf_attributes(self, doc: ezdxf.document.Drawing) -> Dict[str, Any]:
        """
        Извлекает уникальные имена атрибутов из всех блоков в документе.
        Возвращает словарь {tag_name: [list_of_block_names_using_this_tag]}
        """
        attributes = {}
        for block in doc.blocks:
            if block.is_layout or block.is_anonymous or not block.has_attdefs:
                continue
            
            for attdef in block.get_attdefs():
                tag = attdef.dxf.tag
                if tag not in attributes:
                    attributes[tag] = []
                attributes[tag].append(block.name)
                
        logger.info(f"Извлечено {len(attributes)} уникальных атрибутов DWG.")
        return attributes

# --- Код для __init__.py удален, так как он теперь в отдельном файле ---

if __name__ == '__main__':
    # Тестирование io_manager
    logger.info("Тестирование io_manager...")
    pass

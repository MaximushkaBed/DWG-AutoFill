import ezdxf
from ezdxf.document import Drawing
from ezdxf.entities import Attrib
from typing import Dict, Any, List, Tuple
import pandas as pd
import os
import copy

# Тип для хранения информации об измененной сущности
ChangedEntity = Dict[str, Any]

class Filler:
    """
    Модуль для подстановки значений в DWG/DXF и сбора геометрии 
    изменённых сущностей для подсветки.
    """

    def _compute_attrib_bbox(self, attrib: Attrib, insert: ezdxf.entities.Insert) -> Tuple[float, float, float, float]:
        """
        Приблизительный расчет bounding box для атрибута.
        В ezdxf нет простого метода для точного bbox атрибута, 
        поэтому используем приблизительную оценку на основе точки вставки и высоты текста.
        """
        # Получаем точку вставки атрибута в мировых координатах
        insert_point = attrib.dxf.insert
        text_height = attrib.dxf.height
        
        # Приблизительная ширина текста (очень грубо, зависит от шрифта)
        # Для MVP примем, что ширина текста примерно в 5 раз больше высоты
        text_width = text_height * 5 
        
        # Учитываем выравнивание (очень упрощенно)
        # Для простоты MVP, будем считать, что bbox центрирован вокруг точки вставки
        # В реальном приложении нужно учитывать attrib.dxf.halign и attrib.dxf.valign
        
        xmin = insert_point.x - text_width / 2
        ymin = insert_point.y - text_height / 2
        xmax = insert_point.x + text_width / 2
        ymax = insert_point.y + text_height / 2
        
        # TODO: Учесть трансформацию блока (insert.dxf.insert, insert.dxf.rotation, insert.dxf.scale)
        # Для MVP, предполагаем, что атрибуты находятся в блоках без сложной трансформации.
        
        return (xmin, ymin, xmax, ymax)

    def fill_document(self, doc: Drawing, mapping: Dict[str, str], row: Dict[str, Any]) -> Tuple[Drawing, List[ChangedEntity]]:
        """
        Заполняет документ DWG/DXF данными из одной строки (row) согласно маппингу.
        Возвращает новый документ и список измененных сущностей.
        """
        # Создаем копию документа для безопасного заполнения
        # ezdxf не имеет простого copy(), поэтому будем работать с оригиналом, 
        # но в реальном приложении лучше перезагружать/копировать
        
        # В целях MVP и простоты, будем считать, что doc - это свежезагруженный документ
        # и мы работаем с ним напрямую. Для пакетной генерации doc должен быть перезагружен
        # для каждой строки.
        
        changed_entities: List[ChangedEntity] = []
        
        # Перебираем все вхождения блоков (INSERT) в Modelspace
        for insert in doc.modelspace().query('INSERT'):
            # Проверяем, есть ли у этого блока атрибуты, которые нужно заполнить
            for col_name, attrib_tag in mapping.items():
                
                # Получаем значение из строки данных
                value = row.get(col_name)
                if value is None:
                    continue # Пропускаем, если нет данных для этой колонки
                
                # Ищем атрибут по тегу
                if insert.has_attrib(attrib_tag):
                    attrib = insert.get_attrib(attrib_tag)
                    
                    # Преобразуем значение в строку
                    new_text = str(value)
                    
                    # Проверяем, изменилось ли значение
                    if attrib.dxf.text != new_text:
                        # Заполняем атрибут
                        attrib.dxf.text = new_text
                        
                        # Вычисляем bounding box
                        bbox = self._compute_attrib_bbox(attrib, insert)
                        
                        # Сохраняем информацию об изменении
                        changed_entities.append({
                            "handle": attrib.dxf.handle,
                            "bbox": bbox,
                            "attribute_tag": attrib_tag,
                            "new_value": new_text,
                            "block_name": insert.dxf.name
                        })
                        
        return doc, changed_entities

    def batch_fill(self, dwg_path: str, dataframe: pd.DataFrame, mapping: Dict[str, str], out_dir: str, io_manager) -> List[Dict[str, Any]]:
        """
        Пакетная генерация DWG-файлов.
        """
        results: List[Dict[str, Any]] = []
        
        # Убедимся, что выходная директория существует
        io_manager.ensure_directory(out_dir)
        
        for index, row_series in dataframe.iterrows():
            row = row_series.to_dict()
            status = "SUCCESS"
            error_message = None
            out_path = os.path.join(out_dir, f"output_{index+1}.dwg")
            
            try:
                # 1. Перезагружаем документ для каждой строки, чтобы избежать накопления состояния
                doc = io_manager.load_dwg(dwg_path)
                
                # 2. Заполняем документ
                doc_filled, changed_entities = self.fill_document(doc, mapping, row)
                
                # 3. Сохраняем результат
                # Формирование имени файла по шаблону (упрощенно)
                project_name = row.get('PROJECT_NAME', 'Project') # Предполагаем, что есть колонка PROJECT_NAME
                out_path = os.path.join(out_dir, f"{project_name}_{index+1}.dwg")
                
                io_manager.save_dwg(doc_filled, out_path)
                
            except Exception as e:
                status = "FAILED"
                error_message = str(e)
                
            results.append({
                "row_index": index,
                "status": status,
                "output_path": out_path if status == "SUCCESS" else None,
                "error": error_message,
                "changed_count": len(changed_entities) if status == "SUCCESS" else 0
            })
            
        return results

# Обновление src/__init__.py
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
    f.write('from .filler import Filler, ChangedEntity\n')

if __name__ == '__main__':
    # Тестирование модуля filler требует реальных DWG/DXF файлов, 
    # что невозможно в текущем окружении.
    print("Модуль Filler создан. Для тестирования необходимы реальные DWG/DXF файлы.")
    pass

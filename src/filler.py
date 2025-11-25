import ezdxf
from ezdxf.document import Drawing
from ezdxf.entities import Attrib
from typing import Dict, Any, List, Tuple
import pandas as pd
import os
import copy
from .logger import logger

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
        Используем улучшенную эвристику, учитывающую выравнивание.
        """
        # Получаем точку вставки атрибута в мировых координатах
        insert_point = attrib.dxf.insert
        text_height = attrib.dxf.height
        
        # Используем коэффициент 5 для ширины (типично для атрибутов)
        text_width = text_height * 5 
        
        # Горизонтальное выравнивание (0=Left, 1=Center, 2=Right)
        halign = attrib.dxf.halign
        if halign == 0: # Left
            xmin = insert_point.x
            xmax = insert_point.x + text_width
        elif halign == 1: # Center
            xmin = insert_point.x - text_width / 2
            xmax = insert_point.x + text_width / 2
        elif halign == 2: # Right
            xmin = insert_point.x - text_width
            xmax = insert_point.x
        else: # Default to Left
            xmin = insert_point.x
            xmax = insert_point.x + text_width

        # Вертикальное выравнивание (0=Baseline, 1=Bottom, 2=Middle, 3=Top)
        valign = attrib.dxf.valign
        if valign == 0 or valign == 1: # Baseline/Bottom
            ymin = insert_point.y
            ymax = insert_point.y + text_height
        elif valign == 2: # Middle
            ymin = insert_point.y - text_height / 2
            ymax = insert_point.y + text_height / 2
        elif valign == 3: # Top
            ymin = insert_point.y - text_height
            ymax = insert_point.y
        else: # Default to Baseline
            ymin = insert_point.y
            ymax = insert_point.y + text_height
        
        # TODO: Учесть трансформацию блока (insert.dxf.insert, insert.dxf.rotation, insert.dxf.scale)
        # Для MVP, предполагаем, что атрибуты находятся в блоках без сложной трансформации.
        
        return (xmin, ymin, xmax, ymax)

    def fill_document(self, doc: Drawing, mapping: Dict[str, str], row: Dict[str, Any]) -> Tuple[Drawing, List[ChangedEntity]]:
        """
        Заполняет документ DWG/DXF данными из одной строки (row) согласно маппингу.
        Возвращает заполненный документ и список измененных сущностей.
        """
        changed_entities: List[ChangedEntity] = []
        
        # Перебираем все вхождения блоков (INSERT) в Modelspace
        for insert in doc.modelspace().query('INSERT'):
            # Проверяем, есть ли у этого блока атрибуты, которые нужно заполнить
            for col_name, attrib_tag in mapping.items():
                
                # Получаем значение из строки данных
                value = row.get(col_name)
                if value is None or pd.isna(value): # Проверка на None и NaN из pandas
                    continue # Пропускаем, если нет данных для этой колонки
                
                # Ищем атрибут по тегу
                if insert.has_attrib(attrib_tag):
                    attrib = insert.get_attrib(attrib_tag)
                    
                    # Преобразуем значение в строку
                    new_text = str(value)
                    
                    # Проверяем, изменилось ли значение (игнорируем регистр для простоты)
                    if str(attrib.dxf.text).strip() != new_text.strip():
                        logger.info(f"Заполнение атрибута {attrib_tag} в блоке {insert.dxf.name}: '{attrib.dxf.text}' -> '{new_text}'")
                        
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

    def batch_fill(self, dwg_path: str, dataframe: pd.DataFrame, mapping: Dict[str, str], out_dir: str, io_manager) -> Dict[str, Any]:
        """
        Пакетная генерация DWG-файлов с отчетом.
        Возвращает словарь с отчетом о генерации.
        """
        
        report: Dict[str, Any] = {
            "total_rows": len(dataframe),
            "success_count": 0,
            "failed_count": 0,
            "results": []
        }
        
        logger.info(f"Начало пакетной генерации. Всего строк: {report['total_rows']}")
        
        # Убедимся, что выходная директория существует
        io_manager.ensure_directory(out_dir)
        
        for index, row_series in dataframe.iterrows():
            row = row_series.to_dict()
            row_id = index + 1 # Для удобства пользователя
            status = "SUCCESS"
            error_message = None
            changed_entities = []
            out_path = os.path.join(out_dir, f"output_{row_id}.dwg")
            
            try:
                # 1. Перезагружаем документ для каждой строки, чтобы избежать накопления состояния
                logger.info(f"Обработка строки {row_id}/{report['total_rows']}. Загрузка шаблона...", context={'row_id': row_id})
                doc = io_manager.load_dwg(dwg_path)
                
                # 2. Заполняем документ
                doc_filled, changed_entities = self.fill_document(doc, mapping, row)
                
                # 3. Сохраняем результат
                # Формирование имени файла по шаблону (упрощенно)
                # Используем PROJECT_NAME из данных, если есть, иначе 'Project'
                project_name = str(row.get('PROJECT_NAME', 'Project')).replace(' ', '_')
                out_path = os.path.join(out_dir, f"{project_name}_{row_id}.dwg")
                
                io_manager.save_dwg(doc_filled, out_path)
                report['success_count'] += 1
                
            except Exception as e:
                logger.error(f"Ошибка при обработке строки {row_id}: {e}", context={'row_id': row_id})
                report['failed_count'] += 1
                status = "FAILED"
                error_message = str(e)
                out_path = None
                
            report['results'].append({
                "row_index": row_id,
                "status": status,
                "output_path": out_path,
                "error": error_message,
                "changed_count": len(changed_entities)
            })
            
        logger.info(f"Пакетная генерация завершена. Успешно: {report['success_count']}, Ошибок: {report['failed_count']}")
        return report

# --- Код для __init__.py удален, так как он теперь в отдельном файле ---

if __name__ == '__main__':
    # Тестирование модуля filler требует реальных DWG/DXF файлов, 
    # что невозможно в текущем окружении.
    from .io_manager import IOManager
    from .mapper import Mapper
    from .logger import logger
    logger.info("Модуль Filler создан. Для тестирования необходимы реальные DWG/DXF файлы.")
    pass

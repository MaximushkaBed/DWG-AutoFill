import pandas as pd
from typing import List, Dict, Any
from rapidfuzz import process, fuzz
import re
import json
import os

class Mapper:
    """
    Модуль для автоматического и ручного сопоставления колонок таблицы 
    с именами атрибутов блоков в DWG.
    """
    
    # Словарь синонимов для улучшения автоподбора (пример)
    SYNONYMS = {
        "desc": "description",
        "addr": "address",
        "power": "kw",
        "project": "projectname",
        "num": "number",
        "no": "number",
    }

    def __init__(self, fuzzy_threshold: int = 70):
        self.fuzzy_threshold = fuzzy_threshold

    def _normalize(self, name: str) -> str:
        """
        Нормализация имени: нижний регистр, удаление пробелов/символов, транслитерация (упрощенно).
        """
        name = name.lower()
        name = re.sub(r'[^a-z0-9]', '', name)
        
        # Применение синонимов
        for k, v in self.SYNONYMS.items():
            if k in name:
                name = name.replace(k, v)
                
        return name

    def auto_map(self, df_columns: List[str], dxf_attributes: List[str]) -> Dict[str, str]:
        """
        Автоматическое сопоставление колонок DataFrame с атрибутами DWG.
        Возвращает словарь: {column_name: dxf_attribute_tag}
        """
        mapping: Dict[str, str] = {}
        
        # Нормализованные списки для поиска
        normalized_attrs = {self._normalize(attr): attr for attr in dxf_attributes}
        normalized_attr_list = list(normalized_attrs.keys())
        
        for col in df_columns:
            normalized_col = self._normalize(col)
            
            # 1. Exact match (по нормализованному имени)
            if normalized_col in normalized_attrs:
                mapping[col] = normalized_attrs[normalized_col]
                continue
            
            # 2. Fuzzy matching
            if normalized_attr_list:
                match_norm, score, _ = process.extractOne(
                    normalized_col, 
                    normalized_attr_list, 
                    scorer=fuzz.ratio
                )
                
                if score >= self.fuzzy_threshold:
                    # Находим оригинальное имя атрибута по нормализованному совпадению
                    original_attr = normalized_attrs[match_norm]
                    mapping[col] = original_attr
                    
        return mapping

    def interactive_map(self, df_columns: List[str], dxf_attributes: List[str], current_mapping: Dict[str, str] = None) -> Dict[str, str]:
        """
        Заглушка для ручного/интерактивного маппинга. 
        В реальном приложении это будет реализовано через GUI.
        """
        if current_mapping is None:
            current_mapping = {}
            
        # В GUI здесь будет логика отображения таблицы сопоставления
        # и возможность пользователя изменить/добавить/удалить связи.
        
        # Для целей разработки просто возвращаем текущий авто-маппинг
        return self.auto_map(df_columns, dxf_attributes)

    def save_map(self, mapping: Dict[str, str], path: str) -> None:
        """
        Сохраняет сопоставление в JSON файл.
        """
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, indent=4, ensure_ascii=False)

    def load_map(self, path: str) -> Dict[str, str]:
        """
        Загружает сопоставление из JSON файла.
        """
        if not os.path.exists(path):
            return {}
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

# Обновление src/__init__.py
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
    f.write('from .mapper import Mapper\n')

if __name__ == '__main__':
    # Пример использования (для отладки)
    mapper = Mapper()
    
    # Пример колонок из Excel
    excel_cols = ["Project_Name", "Address", "Power_kW", "Description", "Date_of_Issue", "Unused_Column"]
    
    # Пример атрибутов из DWG
    dwg_attrs = ["PROJECTNAME", "ADDRESS", "POWER_KW", "DESC", "ISSUE_DATE", "REVISION_NO"]
    
    print("--- Тест авто-маппинга ---")
    auto_map_result = mapper.auto_map(excel_cols, dwg_attrs)
    
    print("Колонки Excel:", excel_cols)
    print("Атрибуты DWG:", dwg_attrs)
    print("\nРезультат авто-маппинга:")
    for col, attr in auto_map_result.items():
        print(f"  {col} -> {attr}")
        
    # Ожидаемый результат:
    # Project_Name -> PROJECTNAME (Exact match)
    # Address -> ADDRESS (Exact match)
    # Power_kW -> POWER_KW (Exact match)
    # Description -> DESC (Fuzzy match, т.к. DESC - синоним)
    # Date_of_Issue -> ISSUE_DATE (Fuzzy match)
    
    # Тест сохранения/загрузки
    temp_map_file = "test_mapping.json"
    mapper.save_map(auto_map_result, temp_map_file)
    loaded_map = mapper.load_map(temp_map_file)
    print(f"\nСохранено и загружено сопоставление. Количество элементов: {len(loaded_map)}")
    os.remove(temp_map_file)

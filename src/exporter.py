from ezdxf.document import Drawing
from typing import Optional
from .autocad_bridge import AutoCADBridge

class Exporter:
    """
    Модуль для сохранения DWG и опционального экспорта в PDF.
    """
    
    def __init__(self, io_manager, autocad_bridge: AutoCADBridge):
        self.io_manager = io_manager
        self.autocad_bridge = autocad_bridge

    def save_dwg(self, doc: Drawing, path: str, version: str = 'AC1021') -> None:
        """
        Сохраняет изменённый DWG/DXF документ, используя io_manager.
        """
        self.io_manager.save_dwg(doc, path, version)

    def export_pdf(self, dwg_path: str, pdf_path: str) -> bool:
        """
        Экспортирует DWG в PDF. Использует AutoCAD COM, если доступен.
        """
        if self.autocad_bridge.is_available:
            print("Используется AutoCAD COM для экспорта PDF.")
            return self.autocad_bridge.export_pdf(dwg_path, pdf_path)
        else:
            # Fallback: можно использовать ezdxf для экспорта в PDF, 
            # но это не будет WYSIWYG. Для MVP оставляем только AutoCAD-экспорт.
            print("AutoCAD не обнаружен. Экспорт PDF невозможен.")
            return False

# Обновление src/__init__.py
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
    f.write('from .autocad_bridge import AutoCADBridge\n')
    f.write('from .exporter import Exporter\n')

if __name__ == '__main__':
    print("Модули AutoCADBridge и Exporter созданы.")
    pass

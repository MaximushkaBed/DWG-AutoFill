import os
from typing import Optional, Any
from PIL import Image
import io

# Заглушка для win32com, так как он не работает в Linux
try:
    import win32com.client
    IS_WINDOWS = os.name == 'nt'
except ImportError:
    win32com = None
    IS_WINDOWS = False

class AutoCADBridge:
    """
    Модуль связи с AutoCAD COM API для PRO-режима (только Windows).
    """
    
    def __init__(self):
        self.acad_app: Optional[Any] = None
        self.is_available = self._check_availability()

    def _check_availability(self) -> bool:
        """
        Проверяет наличие AutoCAD COM-объекта.
        """
        if not IS_WINDOWS or win32com is None:
            return False
        
        try:
            # Попытка создать COM-объект AutoCAD
            self.acad_app = win32com.client.Dispatch("AutoCAD.Application")
            # Если объект создан, значит AutoCAD доступен
            return True
        except Exception:
            self.acad_app = None
            return False

    def open_document(self, path: str) -> Optional[Any]:
        """
        Открывает DWG-документ в AutoCAD.
        """
        if not self.is_available or not self.acad_app:
            return None
        
        try:
            self.acad_app.Visible = False # Работаем в фоновом режиме
            doc = self.acad_app.Documents.Open(path)
            return doc
        except Exception as e:
            print(f"Ошибка при открытии документа в AutoCAD: {e}")
            return None

    def export_pdf(self, dwg_path: str, pdf_path: str) -> bool:
        """
        Экспортирует DWG в PDF высокого качества через AutoCAD.
        """
        if not self.is_available:
            return False
        
        doc = None
        try:
            doc = self.open_document(dwg_path)
            if not doc:
                return False
            
            # В реальном приложении здесь будет сложная логика PlotConfiguration
            # Упрощенный пример:
            layout = doc.ActiveLayout
            layout.ConfigName = "DWG To PDF.pc3"
            layout.PlotToFile(pdf_path)
            
            return True
        except Exception as e:
            print(f"Ошибка экспорта PDF через AutoCAD: {e}")
            return False
        finally:
            if doc:
                doc.Close(False) # Закрыть без сохранения изменений

    def render_preview(self, dwg_path: str, dpi: int = 300) -> Optional[Image.Image]:
        """
        Рендерит предпросмотр DWG в PNG (в памяти) через AutoCAD и возвращает PIL Image.
        """
        if not self.is_available:
            return None
        
        doc = None
        temp_png_path = os.path.join(os.environ.get('TEMP', '/tmp'), "acad_preview.png")
        
        try:
            doc = self.open_document(dwg_path)
            if not doc:
                return None
            
            # AutoCAD COM API не имеет прямого метода "рендерить в память"
            # Приходится экспортировать во временный файл
            doc.Export(temp_png_path, "PNG", None) # None - для экспорта всего документа
            
            # Загружаем PNG во временный буфер и удаляем файл
            img = Image.open(temp_png_path)
            img_copy = img.copy()
            img.close()
            os.remove(temp_png_path)
            
            return img_copy
            
        except Exception as e:
            print(f"Ошибка рендеринга предпросмотра через AutoCAD: {e}")
            return None
        finally:
            if doc:
                doc.Close(False)
                
    def close_app(self):
        """
        Закрывает экземпляр AutoCAD.
        """
        if self.acad_app:
            try:
                self.acad_app.Quit()
            except Exception:
                pass
            self.acad_app = None

# Обновление src/__init__.py
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
    f.write('from .autocad_bridge import AutoCADBridge\n')

if __name__ == '__main__':
    # Тестирование возможно только на Windows с установленным AutoCAD
    bridge = AutoCADBridge()
    print(f"AutoCAD COM доступен: {bridge.is_available}")
    pass

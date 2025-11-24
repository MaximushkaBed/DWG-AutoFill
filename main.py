import sys
import os

# Добавляем директорию src в путь для импорта модулей
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.gui import DWGAutoFillGUI

if __name__ == '__main__':
    # Запуск основного цикла GUI
    app = DWGAutoFillGUI()
    app.run()

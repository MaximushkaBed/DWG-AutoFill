import logging
import logging.handlers
import os
from typing import Optional

class AppLogger:
    """
    Централизованный модуль логирования для приложения.
    """
    
    def __init__(self, log_file: str = 'dwg_autofill.log', max_bytes: int = 10*1024*1024, backup_count: int = 5):
        self.logger = logging.getLogger('DWGAutoFill')
        self.logger.setLevel(logging.INFO)
        
        # Формат лога
        formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        
        # Файловый хендлер (с ротацией)
        log_dir = os.path.join(os.getcwd(), 'logs')
        os.makedirs(log_dir, exist_ok=True)
        file_handler = logging.handlers.RotatingFileHandler(
            os.path.join(log_dir, log_file), 
            maxBytes=max_bytes, 
            backupCount=backup_count,
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        # Консольный хендлер (для отладки)
        # stream_handler = logging.StreamHandler()
        # stream_handler.setFormatter(formatter)
        # self.logger.addHandler(stream_handler)

    def info(self, message: str, context: Optional[dict] = None):
        self.logger.info(self._format_message(message, context))

    def warning(self, message: str, context: Optional[dict] = None):
        self.logger.warning(self._format_message(message, context))

    def error(self, message: str, context: Optional[dict] = None):
        self.logger.error(self._format_message(message, context))

    def _format_message(self, message: str, context: Optional[dict]) -> str:
        if context:
            context_str = ' | '.join([f"{k}: {v}" for k, v in context.items()])
            return f"{message} ({context_str})"
        return message

# Создаем глобальный экземпляр логгера
logger = AppLogger().logger

# Обновление src/__init__.py
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
    f.write('from .logger import logger\n')

if __name__ == '__main__':
    logger.info("Тест логирования: Приложение запущено.")
    logger.warning("Тест логирования: Не найден файл конфигурации.", context={'file': 'config.json'})
    logger.error("Тест логирования: Критическая ошибка.", context={'module': 'io_manager', 'exception': 'FileNotFound'})

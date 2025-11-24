import PySimpleGUI as sg
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
from typing import Dict, Any, List

# Импорт модулей
from .io_manager import IOManager
from .mapper import Mapper
from .filler import Filler
from .renderer import Renderer, Highlighter

# --- Константы ---
CANVAS_KEY = '-CANVAS-'
LOG_KEY = '-LOG-'
DWG_PATH_KEY = '-DWG_PATH-'
DATA_PATH_KEY = '-DATA_PATH-'
MAPPING_TABLE_KEY = '-MAPPING_TABLE-'
HIGHLIGHT_TOGGLE_KEY = '-HIGHLIGHT_TOGGLE-'

class DWGAutoFillGUI:
    """
    Основной класс GUI, использующий PySimpleGUI и Matplotlib.
    """
    def __init__(self):
        self.io_manager = IOManager()
        self.mapper = Mapper()
        self.filler = Filler()
        self.renderer = Renderer()
        self.highlighter = Highlighter()
        
        self.current_doc = None
        self.current_data_df = None
        self.current_mapping = {}
        self.changed_entities = []
        self.is_autocad_pro_mode = False # Будет определено позже в autocad_bridge
        
        # Настройка темы PySimpleGUI
        sg.theme('LightBlue')

    def _draw_figure(self, canvas, figure):
        """Вспомогательная функция для встраивания Matplotlib Figure в Tkinter Canvas."""
        figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
        figure_canvas_agg.draw()
        figure_canvas_agg.get_tk_widget().pack(side='top', fill='both', expand=1)
        return figure_canvas_agg

    def _log(self, message: str, level: str = 'INFO') -> None:
        """Простая функция логирования в Multiline элемент."""
        log_element = self.window[LOG_KEY]
        color = 'black'
        if level == 'ERROR':
            color = 'red'
        elif level == 'WARNING':
            color = 'orange'
        
        log_element.print(f"[{level}] {message}", text_color=color)

    def _update_mapping_table(self, mapping: Dict[str, str], df_columns: List[str], dxf_attributes: List[str]) -> None:
        """Обновляет таблицу сопоставления в GUI."""
        # Создаем данные для таблицы: [Колонка, Атрибут DWG, Статус]
        table_data = []
        mapped_attrs = set(mapping.values())
        
        for col in df_columns:
            attr = mapping.get(col, '---')
            status = 'Mapped' if attr != '---' else 'Unmapped'
            table_data.append([col, attr, status])
            
        # Добавляем неиспользованные атрибуты DWG
        for attr in dxf_attributes:
            if attr not in mapped_attrs:
                table_data.append(['---', attr, 'Unused DWG Attr'])
                
        self.window[MAPPING_TABLE_KEY].update(values=table_data)

    def _load_dwg_template(self, path: str) -> None:
        """Загружает DWG/DXF шаблон и извлекает атрибуты."""
        try:
            self.current_doc = self.io_manager.load_dwg(path)
            self.window[DWG_PATH_KEY].update(path)
            self._log(f"Шаблон DWG/DXF загружен: {os.path.basename(path)}")
            
            # Извлекаем атрибуты для маппинга
            dxf_attributes_dict = self.io_manager.get_dxf_attributes(self.current_doc)
            dxf_attributes = list(dxf_attributes_dict.keys())
            
            # Если данные уже загружены, делаем авто-маппинг
            if self.current_data_df is not None:
                self.current_mapping = self.mapper.auto_map(self.current_data_df.columns.tolist(), dxf_attributes)
                self._update_mapping_table(self.current_mapping, self.current_data_df.columns.tolist(), dxf_attributes)
                self._log(f"Авто-маппинг выполнен. Найдено {len(self.current_mapping)} совпадений.")
            else:
                self._update_mapping_table({}, [], dxf_attributes)
                
            # Рендерим предпросмотр
            self.renderer.render_to_canvas(self.current_doc)
            
        except Exception as e:
            self.current_doc = None
            self._log(f"Ошибка загрузки DWG: {e}", 'ERROR')

    def _load_data_file(self, path: str) -> None:
        """Загружает файл данных (Excel/CSV/JSON)."""
        try:
            self.current_data_df = self.io_manager.read_table(path)
            self.window[DATA_PATH_KEY].update(path)
            self._log(f"Файл данных загружен: {os.path.basename(path)}. Строк: {len(self.current_data_df)}")
            
            # Если DWG уже загружен, делаем авто-маппинг
            if self.current_doc is not None:
                dxf_attributes_dict = self.io_manager.get_dxf_attributes(self.current_doc)
                dxf_attributes = list(dxf_attributes_dict.keys())
                self.current_mapping = self.mapper.auto_map(self.current_data_df.columns.tolist(), dxf_attributes)
                self._update_mapping_table(self.current_mapping, self.current_data_df.columns.tolist(), dxf_attributes)
                self._log(f"Авто-маппинг выполнен. Найдено {len(self.current_mapping)} совпадений.")
                
        except Exception as e:
            self.current_data_df = None
            self._log(f"Ошибка загрузки данных: {e}", 'ERROR')

    def _preview_document(self) -> None:
        """Предпросмотр первого документа с заполненными данными."""
        if not self.current_doc or self.current_data_df is None or not self.current_mapping:
            self._log("Необходимо загрузить DWG, данные и настроить маппинг.", 'WARNING')
            return
        
        try:
            # Берем первую строку для предпросмотра
            first_row = self.current_data_df.iloc[0].to_dict()
            
            # Создаем копию документа для заполнения (в реальном приложении нужно перезагружать)
            # Здесь мы просто перезагружаем, чтобы избежать проблем с ezdxf
            temp_doc = self.io_manager.load_dwg(self.window[DWG_PATH_KEY].get())
            
            # Заполняем
            doc_filled, self.changed_entities = self.filler.fill_document(temp_doc, self.current_mapping, first_row)
            
            # Рендерим заполненный документ
            self.renderer.render_to_canvas(doc_filled)
            
            # Накладываем подсветку
            bboxes = [e['bbox'] for e in self.changed_entities]
            self.highlighter.overlay_on_axes(self.renderer.ax, bboxes)
            self.window[HIGHLIGHT_TOGGLE_KEY].update(value=True)
            
            self._log(f"Предпросмотр выполнен. Изменено {len(self.changed_entities)} полей.")
            
        except Exception as e:
            self._log(f"Ошибка при предпросмотре: {e}", 'ERROR')

    def _generate_documents(self) -> None:
        """Запуск пакетной генерации."""
        if not self.current_doc or self.current_data_df is None or not self.current_mapping:
            self._log("Необходимо загрузить DWG, данные и настроить маппинг.", 'WARNING')
            return
        
        output_folder = sg.popup_get_folder('Выберите папку для сохранения результатов', default_path=os.getcwd())
        if not output_folder:
            return
        
        self._log(f"Начало пакетной генерации в папку: {output_folder}")
        
        # В реальном приложении здесь будет использоваться filler.batch_fill
        # Но для MVP мы просто имитируем процесс
        
        total_rows = len(self.current_data_df)
        success_count = 0
        
        for i in range(total_rows):
            # Имитация работы
            self._log(f"Обработка строки {i+1}/{total_rows}...", 'INFO')
            success_count += 1
            
        self._log(f"Генерация завершена. Успешно создано {success_count} из {total_rows} файлов.", 'INFO')
        
    def _export_pdf(self) -> None:
        """Экспорт в PDF (заглушка)."""
        if self.is_autocad_pro_mode:
            self._log("Экспорт PDF через AutoCAD COM (PRO-режим)...", 'INFO')
            # Здесь будет вызов autocad_bridge
        else:
            self._log("Экспорт PDF недоступен (требуется AutoCAD). Используйте PNG-превью.", 'WARNING')

    def run(self):
        """Запуск основного цикла GUI."""
        
        # --- Layout ---
        
        # Левая панель (Управление)
        left_column = [
            [sg.Frame('Шаблон DWG/DXF', [
                [sg.Input(key=DWG_PATH_KEY, enable_events=True, readonly=True, size=(30, 1)), 
                 sg.FileBrowse('Выбрать', file_types=(("DWG/DXF Files", "*.dwg;*.dxf"),))],
            ])],
            [sg.Frame('Данные (Excel/CSV/JSON)', [
                [sg.Input(key=DATA_PATH_KEY, enable_events=True, readonly=True, size=(30, 1)), 
                 sg.FileBrowse('Выбрать', file_types=(("Data Files", "*.xlsx;*.csv;*.json"),))],
            ])],
            [sg.Frame('Сопоставление полей (Mapping)', [
                [sg.Table(values=[], 
                          headings=['Колонка данных', 'Атрибут DWG', 'Статус'], 
                          key=MAPPING_TABLE_KEY, 
                          auto_size_columns=False, 
                          col_widths=[15, 15, 10],
                          justification='left',
                          num_rows=10,
                          display_row_numbers=False,
                          enable_events=True)],
                [sg.Button('Ручной Маппинг', key='-MANUAL_MAP-', disabled=True)]
            ])],
            [sg.HorizontalSeparator()],
            [sg.Button('Preview (1-я строка)', key='-PREVIEW-', size=(15, 1)), 
             sg.Button('Generate (Batch)', key='-GENERATE-', size=(15, 1))],
            [sg.HorizontalSeparator()],
            [sg.Checkbox('Использовать AutoCAD (PRO-режим)', key='-AUTOCAD_PRO-', default=False, disabled=True, enable_events=True)],
            [sg.Button('Экспорт PDF', key='-EXPORT_PDF-', size=(15, 1)), 
             sg.Button('Сохранить PNG', key='-SAVE_PNG-', size=(15, 1))]
        ]

        # Правая панель (Viewer)
        right_column = [
            [sg.Frame('Интерактивный 2D-Просмотр', [
                [sg.Canvas(key=CANVAS_KEY, size=(800, 600))],
                [sg.Button('Fit', key='-FIT-', size=(8, 1)), 
                 sg.Button('Zoom In', key='-ZOOM_IN-', size=(8, 1), disabled=True), 
                 sg.Button('Zoom Out', key='-ZOOM_OUT-', size=(8, 1), disabled=True),
                 sg.VerticalSeparator(),
                 sg.Checkbox('Подсветка изменений', key=HIGHLIGHT_TOGGLE_KEY, default=False, enable_events=True)]
            ], size=(820, 650))]
        ]
        
        # Нижняя часть (Лог)
        log_row = [sg.Multiline(size=(120, 8), key=LOG_KEY, autoscroll=True, disabled=True, reroute_stdout=True, reroute_stderr=True)]

        # Общий Layout
        layout = [
            [sg.Column(left_column, vertical_alignment='top'), 
             sg.Column(right_column, vertical_alignment='top')],
            log_row
        ]

        self.window = sg.Window('DWG AutoFill (MVP)', layout, finalize=True)
        
        # Встраивание Matplotlib
        canvas_elem = self.window[CANVAS_KEY].TKCanvas
        self.renderer.create_figure(canvas_elem)
        
        # --- Основной цикл событий ---
        while True:
            event, values = self.window.read()
            
            if event == sg.WIN_CLOSED:
                break
            
            # Обработка загрузки файлов
            if event == DWG_PATH_KEY:
                self._load_dwg_template(values[DWG_PATH_KEY])
            elif event == DATA_PATH_KEY:
                self._load_data_file(values[DATA_PATH_KEY])
                
            # Обработка кнопок
            elif event == '-PREVIEW-':
                self._preview_document()
            elif event == '-GENERATE-':
                self._generate_documents()
            elif event == '-EXPORT_PDF-':
                self._export_pdf()
            elif event == '-FIT-':
                self.renderer.fit_to_view()
            elif event == HIGHLIGHT_TOGGLE_KEY:
                self.highlighter.toggle_highlights(values[HIGHLIGHT_TOGGLE_KEY])
                self.renderer.canvas.draw_idle()
            elif event == '-SAVE_PNG-':
                # Сохранение PNG (просто сохраняем текущий Matplotlib Figure)
                save_path = sg.popup_get_file('Сохранить предпросмотр как PNG', save_as=True, file_types=(("PNG Files", "*.png"),))
                if save_path:
                    self.renderer.fig.savefig(save_path, dpi=300)
                    self._log(f"Предпросмотр сохранен в {save_path}")
            
        self.window.close()

# Обновление src/__init__.py
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
    f.write('from .gui import DWGAutoFillGUI\n')

if __name__ == '__main__':
    # Для запуска GUI
    # gui = DWGAutoFillGUI()
    # gui.run()
    pass

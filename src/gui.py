import FreeSimpleGUI as sg
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
from typing import Dict, Any, List
import pandas as pd

# Импорт модулей
from .io_manager import IOManager
from .mapper import Mapper
from .filler import Filler
from .renderer import Renderer, Highlighter
from .exporter import Exporter
from .autocad_bridge import AutoCADBridge
from .logger import logger

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
        self.autocad_bridge = AutoCADBridge()
        self.renderer = Renderer()
        self.highlighter = Highlighter()
        self.exporter = Exporter(self.io_manager, self.autocad_bridge)
        
        self.current_doc = None
        self.current_data_df = None
        self.current_mapping = {}
        self.changed_entities = []
        self.is_autocad_pro_mode = self.autocad_bridge.is_available
        
        # Настройка темы PySimpleGUI
        sg.theme('LightBlue')

    def _draw_figure(self, canvas, figure):
        """Вспомогательная функция для встраивания Matplotlib Figure в Tkinter Canvas."""
        figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
        figure_canvas_agg.draw()
        figure_canvas_agg.get_tk_widget().pack(side='top', fill='both', expand=1)
        return figure_canvas_agg

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
            logger.info(f"Шаблон DWG/DXF загружен: {os.path.basename(path)}")
            
            # Извлекаем атрибуты для маппинга
            dxf_attributes_dict = self.io_manager.get_dxf_attributes(self.current_doc)
            dxf_attributes = list(dxf_attributes_dict.keys())
            
            # Если данные уже загружены, делаем авто-маппинг
            if self.current_data_df is not None:
                self.current_mapping = self.mapper.auto_map(self.current_data_df.columns.tolist(), dxf_attributes)
                self._update_mapping_table(self.current_mapping, self.current_data_df.columns.tolist(), dxf_attributes)
                logger.info(f"Авто-маппинг выполнен. Найдено {len(self.current_mapping)} совпадений.")
            else:
                self._update_mapping_table({}, [], dxf_attributes)
                
            # Рендерим предпросмотр
            self.renderer.render_to_canvas(self.current_doc)
            
        except Exception as e:
            self.current_doc = None
            logger.error(f"Ошибка загрузки DWG: {e}")
            sg.popup_error(f"Ошибка загрузки DWG/DXF: {e}")

    def _load_data_file(self, path: str) -> None:
        """Загружает файл данных (Excel/CSV/JSON)."""
        try:
            self.current_data_df = self.io_manager.read_table(path)
            self.window[DATA_PATH_KEY].update(path)
            logger.info(f"Файл данных загружен: {os.path.basename(path)}. Строк: {len(self.current_data_df)}")
            
            # Если DWG уже загружен, делаем авто-маппинг
            if self.current_doc is not None:
                dxf_attributes_dict = self.io_manager.get_dxf_attributes(self.current_doc)
                dxf_attributes = list(dxf_attributes_dict.keys())
                self.current_mapping = self.mapper.auto_map(self.current_data_df.columns.tolist(), dxf_attributes)
                self._update_mapping_table(self.current_mapping, self.current_data_df.columns.tolist(), dxf_attributes)
                logger.info(f"Авто-маппинг выполнен. Найдено {len(self.current_mapping)} совпадений.")
                
        except Exception as e:
            self.current_data_df = None
            logger.error(f"Ошибка загрузки данных: {e}")
            sg.popup_error(f"Ошибка загрузки данных: {e}")

    def _preview_document(self) -> None:
        """Предпросмотр первого документа с заполненными данными."""
        if not self.current_doc or self.current_data_df is None or not self.current_mapping:
            logger.warning("Невозможно выполнить предпросмотр: не загружен DWG, данные или маппинг.")
            sg.popup_quick_message("Необходимо загрузить DWG, данные и настроить маппинг.", background_color='red', text_color='white')
            return
        
        try:
            # Берем первую строку для предпросмотра
            first_row = self.current_data_df.iloc[0].to_dict()
            
            # Создаем копию документа для заполнения (перезагружаем)
            temp_doc = self.io_manager.load_dwg(self.window[DWG_PATH_KEY].get())
            
            # Заполняем
            doc_filled, self.changed_entities = self.filler.fill_document(temp_doc, self.current_mapping, first_row)
            
            # Рендерим заполненный документ
            self.renderer.render_to_canvas(doc_filled)
            
            # Накладываем подсветку
            bboxes = [e['bbox'] for e in self.changed_entities]
            self.highlighter.overlay_on_axes(self.renderer.ax, bboxes)
            self.window[HIGHLIGHT_TOGGLE_KEY].update(value=True)
            
            logger.info(f"Предпросмотр выполнен. Изменено {len(self.changed_entities)} полей.")
            
        except Exception as e:
            logger.error(f"Ошибка при предпросмотре: {e}")
            sg.popup_error(f"Ошибка при предпросмотре: {e}")

    def _generate_documents(self) -> None:
        """Запуск пакетной генерации."""
        if not self.current_doc or self.current_data_df is None or not self.current_mapping:
            logger.warning("Невозможно выполнить генерацию: не загружен DWG, данные или маппинг.")
            sg.popup_quick_message("Необходимо загрузить DWG, данные и настроить маппинг.", background_color='red', text_color='white')
            return
        
        output_folder = sg.popup_get_folder('Выберите папку для сохранения результатов', default_path=os.getcwd())
        if not output_folder:
            return
        
        logger.info(f"Начало пакетной генерации в папку: {output_folder}")
        
        # Используем filler.batch_fill с полным отчетом
        report = self.filler.batch_fill(self.window[DWG_PATH_KEY].get(), self.current_data_df, self.current_mapping, output_folder, self.io_manager)
        
        logger.info(f"Пакетная генерация завершена. Успешно создано {report['success_count']} из {report['total_rows']} файлов. Ошибок: {report['failed_count']}.")
        sg.popup_ok(f"Пакетная генерация завершена.\nУспешно: {report['success_count']}\nОшибок: {report['failed_count']}", title="Генерация завершена")

    def _export_pdf(self) -> None:
        """Экспорт в PDF с использованием Exporter."""
        if not self.current_doc:
            logger.warning("Невозможно экспортировать: не загружен DWG.")
            sg.popup_quick_message("Сначала загрузите DWG/DXF шаблон.", background_color='red', text_color='white')
            return
        
        pdf_path = sg.popup_get_file('Сохранить как PDF', save_as=True, file_types=(("PDF Files", "*.pdf"),))
        if not pdf_path:
            return
            
        dwg_path = self.window[DWG_PATH_KEY].get()
        
        try:
            if self.window['-AUTOCAD_PRO-'].get():
                logger.info("Запуск экспорта PDF через AutoCAD COM (PRO-режим)...")
                if self.exporter.export_pdf(dwg_path, pdf_path):
                    logger.info(f"PDF успешно экспортирован через AutoCAD: {pdf_path}")
                    sg.popup_ok(f"PDF успешно экспортирован: {pdf_path}")
                else:
                    logger.error("Экспорт PDF через AutoCAD COM завершился неудачей.")
                    sg.popup_error("Экспорт PDF через AutoCAD COM завершился неудачей. Проверьте логи.")
            else:
                logger.warning("Экспорт PDF недоступен (требуется AutoCAD).")
                sg.popup_ok("Экспорт PDF недоступен. Используйте PNG-превью.")
        except Exception as e:
            logger.error(f"Критическая ошибка при экспорте PDF: {e}")
            sg.popup_error(f"Критическая ошибка при экспорте PDF: {e}")

    def _show_manual_mapping_window(self):
        """Показывает окно ручного сопоставления."""
        if self.current_data_df is None or self.current_doc is None:
            sg.popup_quick_message("Сначала загрузите DWG и данные.", background_color='red', text_color='white')
            return

        df_columns = self.current_data_df.columns.tolist()
        dxf_attributes = list(self.io_manager.get_dxf_attributes(self.current_doc).keys())
        
        # Создаем список для ComboBox: [колонка_данных, атрибут_dwg]
        mapping_rows = []
        for col in df_columns:
            mapped_attr = self.current_mapping.get(col, '')
            mapping_rows.append([col, mapped_attr])

        # Создаем ComboBox с атрибутами DWG для выбора
        dxf_options = [''] + dxf_attributes

        # Layout для ручного маппинга
        manual_map_layout = [
            [sg.Text('Ручное сопоставление полей', font=('Helvetica', 14))],\
            [sg.HorizontalSeparator()],
            [sg.Table(values=mapping_rows,
                      headings=['Колонка данных', 'Сопоставленный атрибут DWG'],
                      key='-MANUAL_TABLE-', 
                      auto_size_columns=False,
                      col_widths=[20, 20],
                      justification='left',
                      num_rows=min(20, len(df_columns) + 2),
                      display_row_numbers=False,
                      enable_events=True,
                      select_mode=sg.TABLE_SELECT_MODE_BROWSE)],
            [sg.Text('Выбранная колонка:', size=(15, 1)), sg.Input(key='-SELECTED_COL-', disabled=True, size=(20, 1))],
            [sg.Text('Выбрать атрибут:', size=(15, 1)), sg.Combo(dxf_options, key='-ATTR_COMBO-', enable_events=True, size=(20, 1))],
            [sg.HorizontalSeparator()],
            [sg.Button('Сохранить', key='-SAVE_MANUAL_MAP-'), sg.Button('Отмена')]
        ]

        window = sg.Window('Ручное сопоставление', manual_map_layout, modal=True, finalize=True)
        selected_row_index = -1

        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Отмена':
                break
            
            if event == '-MANUAL_TABLE-':
                if values['-MANUAL_TABLE-']:
                    selected_row_index = values['-MANUAL_TABLE-'][0]
                    selected_col = mapping_rows[selected_row_index][0]
                    selected_attr = mapping_rows[selected_row_index][1]
                    window['-SELECTED_COL-'].update(selected_col)
                    window['-ATTR_COMBO-'].update(value=selected_attr)
                
            elif event == '-ATTR_COMBO-':
                if selected_row_index != -1:
                    new_attr = values['-ATTR_COMBO-']
                    col_to_map = mapping_rows[selected_row_index][0]
                    
                    # Обновляем локальную таблицу
                    mapping_rows[selected_row_index][1] = new_attr
                    window['-MANUAL_TABLE-'].update(values=mapping_rows)
                    
                    # Обновляем ComboBox, чтобы он показывал новое значение
                    window['-ATTR_COMBO-'].update(value=new_attr)

            elif event == '-SAVE_MANUAL_MAP-':
                # Собираем новое сопоставление
                new_mapping = {}
                for col, attr in mapping_rows:
                    if attr:
                        new_mapping[col] = attr
                
                self.current_mapping = new_mapping
                self._update_mapping_table(self.current_mapping, df_columns, dxf_attributes)
                logger.info(f"Ручное сопоставление сохранено. Всего сопоставлено: {len(new_mapping)}")
                break

        window.close()

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
                [sg.Button('Ручной Маппинг', key='-MANUAL_MAP-', disabled=False)]
            ])],
            [sg.HorizontalSeparator()],
            [sg.Button('Preview (1-я строка)', key='-PREVIEW-', size=(15, 1)), 
             sg.Button('Generate (Batch)', key='-GENERATE-', size=(15, 1))],
            [sg.HorizontalSeparator()],
            [sg.Checkbox('Использовать AutoCAD (PRO-режим)', key='-AUTOCAD_PRO-', default=self.is_autocad_pro_mode, disabled=not self.is_autocad_pro_mode, enable_events=True)],
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
                if not self.renderer.fig:
                    sg.popup_quick_message("Сначала загрузите DWG/DXF шаблон.", background_color='red', text_color='white')
                    continue
                # Сохранение PNG (просто сохраняем текущий Matplotlib Figure)
                save_path = sg.popup_get_file('Сохранить предпросмотр как PNG', save_as=True, file_types=(("PNG Files", "*.png"),))
                if save_path:
                    self.renderer.fig.savefig(save_path, dpi=300)
                    logger.info(f"Предпросмотр сохранен в {save_path}")
                    sg.popup_ok(f"PNG сохранен: {save_path}")
            elif event == '-AUTOCAD_PRO-':
                # Логика для PRO-режима (пока только логирование)
                if values['-AUTOCAD_PRO-']:
                    logger.info("PRO-режим (AutoCAD COM) активирован.")
                else:
                    logger.info("PRO-режим (AutoCAD COM) деактивирован.")
            elif event == '-MANUAL_MAP-':
                self._show_manual_mapping_window()
            
        self.window.close()
        self.autocad_bridge.close_app() # Закрываем COM-объект при выходе

# Обновление src/__init__.py
# with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
#     f.write('from .gui import DWGAutoFillGUI\n')

if __name__ == '__main__':
    # Для запуска GUI
    # gui = DWGAutoFillGUI()
    # gui.run()
    pass

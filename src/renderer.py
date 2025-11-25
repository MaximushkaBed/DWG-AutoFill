import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.axes import Axes
from matplotlib.figure import Figure
from matplotlib.patches import Rectangle
import ezdxf
from ezdxf.document import Drawing
from ezdxf.addons.drawing import matplotlib as ezdxf_mpl
from typing import List, Tuple, Callable, Dict, Any

# Тип для bounding box в мировых координатах
WorldBBox = Tuple[float, float, float, float]

class Renderer:
    """
    Модуль для отрисовки DWG/DXF в интерактивном matplotlib canvas.
    """
    
    def __init__(self):
        self.fig: Figure = None
        self.ax: Axes = None
        self.canvas: FigureCanvasTkAgg = None
        self.dxf_artist = None

    def create_figure(self, parent_tk_canvas) -> FigureCanvasTkAgg:
        """
        Создает Figure и Axes, встраивает их в Tkinter Canvas (для PySimpleGUI).
        Возвращает FigureCanvasTkAgg.
        """
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.ax.set_aspect('equal', adjustable='box')
        self.ax.axis('off') # Отключаем оси
        
        # Создаем Canvas для Tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, master=parent_tk_canvas)
        self.canvas.draw()
        
        return self.canvas

    def render_to_canvas(self, doc: Drawing, layout_name: str = None) -> None:
        """
        Отрисовывает DWG-документ в Axes.
        """
        if not self.ax:
            raise ValueError("Figure and Axes must be created first using create_figure().")
        
        # Очищаем предыдущий рендер
        self.ax.clear()
        self.ax.set_aspect('equal', adjustable='box')
        self.ax.axis('off')
        
        # Выбираем layout
        layout = doc.modelspace()
        if layout_name and layout_name in doc.layouts:
            layout = doc.layouts.get(layout_name)
        
        # Отрисовка с помощью ezdxf.addons.drawing
        try:
            # Создаем новое рисование
            ezdxf_mpl.draw_layout(layout, self.ax)
            
            # Автоматически подгоняем вид
            self.ax.autoscale()
            self.ax.margins(0.1)
            
            self.canvas.draw_idle()
            
        except Exception as e:
            print(f"Ошибка при рендеринге DWG: {e}")
            self.ax.text(0.5, 0.5, f"Ошибка рендеринга: {e}", 
                         transform=self.ax.transAxes, ha='center', va='center', color='red')
            self.canvas.draw_idle()

    def fit_to_view(self) -> None:
        """
        Подгоняет чертеж под размер окна.
        """
        if self.ax:
            self.ax.autoscale()
            self.ax.margins(0.1)
            self.canvas.draw_idle()

    def get_world_to_screen_transform(self) -> Callable[[float, float], Tuple[float, float]]:
        """
        Возвращает функцию для преобразования мировых координат в координаты пикселей на экране.
        """
        if not self.ax:
            raise ValueError("Axes not initialized.")
        
        # transData - преобразование из данных (world) в пиксели (display)
        return self.ax.transData.transform

class Highlighter:
    """
    Модуль для наложения оверлея (подсветки) на matplotlib Axes.
    """
    
    def __init__(self):
        self.patches: List[Rectangle] = []

    def overlay_on_axes(self, ax: Axes, bboxes: List[WorldBBox], style: Dict[str, Any] = None) -> None:
        """
        Отрисовывает прямоугольники-оверлеи на Axes.
        """
        # Удаляем предыдущие патчи
        for patch in self.patches:
            patch.remove()
        self.patches.clear()
        
        default_style = {
            'edgecolor': 'red',
            'facecolor': 'yellow',
            'alpha': 0.3,
            'linewidth': 2,
            'fill': True
        }
        if style:
            default_style.update(style)
            
        for bbox in bboxes:
            xmin, ymin, xmax, ymax = bbox
            width = xmax - xmin
            height = ymax - ymin
            
            # Создаем прямоугольник в мировых координатах
            rect = Rectangle((xmin, ymin), width, height, 
                             transform=ax.transData, **default_style)
            
            ax.add_patch(rect)
            self.patches.append(rect)
            
        ax.figure.canvas.draw_idle()

    def toggle_highlights(self, visible: bool) -> None:
        """
        Показывает или скрывает оверлеи.
        """
        for patch in self.patches:
            patch.set_visible(visible)
        # Перерисовка будет вызвана из GUI
        
# Обновление src/__init__.py
with open(os.path.join(os.path.dirname(__file__), '__init__.py'), 'a') as f:
    f.write('from .renderer import Renderer, Highlighter\n')

if __name__ == '__main__':
    # Тестирование требует GUI-фреймворка (Tkinter/PySimpleGUI)
    print("Модули Renderer и Highlighter созданы. Для тестирования необходим GUI.")
    pass

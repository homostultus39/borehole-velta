"""
Модуль конфигурации для работы со скважинами.
Содержит настройки по умолчанию и константы.
"""

import os
from typing import List, Dict, Any

# Настройки поиска скважин
BOREHOLE_PATTERNS = [
    r'скв[а-я]*\.?\s*(\d+)',  # скв. 123, скважина 123
    r'№\s*(\d+)',             # № 123
    r'(\d+)\s*скв',           # 123 скв
    r'скв\s*(\d+)',           # скв 123
    r'^(\d+)$',               # просто число
]

# Настройки поиска связанных объектов
MAX_DISTANCE_FOR_CIRCLE_TEXT = 50.0  # Максимальное расстояние между текстом и кругом
MAX_DISTANCE_FOR_TEXT_CIRCLE = 50.0  # Максимальное расстояние между кругом и текстом

# Настройки Excel
DEFAULT_OUTPUT_DIRECTORY = "output"
EXCEL_SHEET_NAMES = {
    'BOREHOLES': 'Скважины',
    'SUMMARY': 'Сводка',
    'REFERENCE': 'Опорная скважина',
    'ALL_BOREHOLES': 'Все скважины',
    'REFERENCE_BOREHOLES': 'Опорные скважины',
    'HEIGHT_STATS': 'Статистика высот',
    'COORDINATES': 'Координаты'
}

# Настройки логирования
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_FILE = 'borehole_analysis.log'

# Настройки AutoCAD
AUTOCAD_CONNECTION_TIMEOUT = 30  # Таймаут подключения к AutoCAD в секундах

# Настройки обработки данных
DEFAULT_REFERENCE_Z = 0.0
DEFAULT_HEIGHT_CALCULATION_METHOD = 'y_coordinate'  # Использовать Y-координату как высоту

# Поддерживаемые форматы файлов
SUPPORTED_FILE_EXTENSIONS = ['.dwg', '.dxf']

# Настройки экспорта
EXPORT_SETTINGS = {
    'include_coordinates': True,
    'include_relative_heights': True,
    'include_reference_info': True,
    'sort_by_number': True,
    'create_detailed_report': True
}

def get_default_config() -> Dict[str, Any]:
    """
    Получение конфигурации по умолчанию.
    
    Returns:
        Dict[str, Any]: Словарь с настройками по умолчанию
    """
    return {
        'borehole_patterns': BOREHOLE_PATTERNS,
        'max_distance_circle_text': MAX_DISTANCE_FOR_CIRCLE_TEXT,
        'max_distance_text_circle': MAX_DISTANCE_FOR_TEXT_CIRCLE,
        'output_directory': DEFAULT_OUTPUT_DIRECTORY,
        'excel_sheet_names': EXCEL_SHEET_NAMES,
        'log_format': LOG_FORMAT,
        'log_file': LOG_FILE,
        'autocad_timeout': AUTOCAD_CONNECTION_TIMEOUT,
        'default_reference_z': DEFAULT_REFERENCE_Z,
        'height_calculation_method': DEFAULT_HEIGHT_CALCULATION_METHOD,
        'supported_extensions': SUPPORTED_FILE_EXTENSIONS,
        'export_settings': EXPORT_SETTINGS
    }

def validate_file_extension(file_path: str) -> bool:
    """
    Проверка поддерживаемого расширения файла.
    
    Args:
        file_path: Путь к файлу
        
    Returns:
        bool: True если расширение поддерживается
    """
    _, ext = os.path.splitext(file_path.lower())
    return ext in SUPPORTED_FILE_EXTENSIONS

def get_output_directory() -> str:
    """
    Получение директории для выходных файлов.
    
    Returns:
        str: Путь к директории
    """
    output_dir = os.path.abspath(DEFAULT_OUTPUT_DIRECTORY)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return output_dir


"""
Пример использования скрипта для определения номеров скважин.
"""

import sys
import os

# Добавляем src в путь для импорта
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.main import process_dwg_file
from src.autocad_handler import AutoCADHandler
from src.borehole_processor import BoreholeProcessor
from src.console_output import ConsoleOutput


def example_basic_usage():
    """Пример базового использования."""
    print("=== Пример базового использования ===")
    
    # Путь к вашему .dwg файлу
    dwg_file = "path/to/your/file.dwg"
    
    # Проверяем существование файла
    if not os.path.exists(dwg_file):
        print(f"Файл не найден: {dwg_file}")
        print("Пожалуйста, укажите правильный путь к .dwg файлу")
        return
    
    # Обработка файла
    success = process_dwg_file(
        file_path=dwg_file,
        reference_borehole=None,  # Автоматический выбор первой скважины
        reference_z=0.0
    )
    
    if success:
        print("Обработка завершена успешно!")
    else:
        print("Ошибка при обработке файла")


def example_advanced_usage():
    """Пример расширенного использования с настройками."""
    print("=== Пример расширенного использования ===")
    
    # Инициализация компонентов
    autocad_handler = AutoCADHandler()
    borehole_processor = BoreholeProcessor()
    console_output = ConsoleOutput()
    
    try:
        # Подключение к AutoCAD
        if not autocad_handler.connect():
            print("Ошибка подключения к AutoCAD")
            return
        
        # Открытие файла
        dwg_file = "path/to/your/file.dwg"
        if not autocad_handler.open_dwg(dwg_file):
            print("Ошибка открытия файла")
            return
        
        # Извлечение данных
        text_entities = autocad_handler.find_text_entities()
        circles = autocad_handler.find_circles()
        
        print(f"Найдено текстовых объектов: {len(text_entities)}")
        print(f"Найдено кругов: {len(circles)}")
        
        # Обработка скважин
        boreholes = borehole_processor.extract_borehole_numbers(text_entities, circles)
        print(f"Найдено скважин: {len(boreholes)}")
        
        # Установка опорной скважины
        if boreholes:
            # Выбираем скважину с номером "1" как опорную
            borehole_processor.set_reference_borehole("1")
            
            # Расчет высот
            borehole_processor.calculate_relative_heights(reference_z=100.0)
            
            # Получение данных
            data = borehole_processor.get_boreholes_data()
            
            # Вывод результатов в консоль
            console_output.print_boreholes_summary(data)
            console_output.print_boreholes_table(data)
            print("Данные выведены в консоль")
    
    except Exception as e:
        print(f"Ошибка: {e}")
    
    finally:
        # Закрытие соединений
        autocad_handler.close_document()
        autocad_handler.disconnect()


def example_command_line():
    """Пример использования через командную строку."""
    print("=== Пример использования через командную строку ===")
    print()
    print("Базовое использование:")
    print("python src/main.py path/to/file.dwg")
    print()
    print("С указанием опорной скважины:")
    print("python src/main.py path/to/file.dwg -r 5")
    print()
    print("С настройкой Z-координаты опорной скважины:")
    print("python src/main.py path/to/file.dwg -r 5 -z 100.0")
    print()
    print("С отладочным логированием:")
    print("python src/main.py path/to/file.dwg -l DEBUG")


if __name__ == "__main__":
    print("Примеры использования скрипта для определения номеров скважин")
    print("=" * 60)
    print()
    
    # Показываем примеры командной строки
    example_command_line()
    
    print("\n" + "=" * 60)
    print("Для запуска обработки файла раскомментируйте нужный пример:")
    print()
    print("# example_basic_usage()")
    print("# example_advanced_usage()")


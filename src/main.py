"""
Основной скрипт для определения номеров скважин в .dwg файлах.
Автоматически устанавливает опорную скважину и рассчитывает относительные высоты.
"""

import os
import sys
import logging
import argparse
from typing import Optional

# Добавляем текущую директорию в путь для импорта модулей
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from autocad_handler import AutoCADHandler
from borehole_processor import BoreholeProcessor
from console_output import ConsoleOutput


def setup_logging(level: str = "INFO") -> None:
    """
    Настройка системы логирования.
    
    Args:
        level: Уровень логирования (DEBUG, INFO, WARNING, ERROR)
    """
    logging.basicConfig(
        level=getattr(logging, level.upper()),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('borehole_analysis.log', encoding='utf-8')
        ]
    )


def process_dwg_file(file_path: str, reference_borehole: Optional[str] = None, 
                    reference_z: float = 0.0) -> bool:
    """
    Обработка .dwg файла и извлечение данных о скважинах.
    
    Args:
        file_path: Путь к .dwg файлу
        reference_borehole: Номер опорной скважины
        reference_z: Z-координата опорной скважины
        
    Returns:
        bool: True если обработка успешна
    """
    logger = logging.getLogger(__name__)
    
    # Инициализация компонентов
    autocad_handler = AutoCADHandler()
    borehole_processor = BoreholeProcessor()
    console_output = ConsoleOutput()
    
    try:
        # Вывод заголовка
        console_output.print_header("АНАЛИЗ СКВАЖИН В .DWG ФАЙЛЕ")
        console_output.print_file_info(file_path)
        
        # Подключение к AutoCAD
        console_output.print_autocad_connection_status(False)
        if not autocad_handler.connect():
            console_output.print_error_message("Не удалось подключиться к AutoCAD")
            return False
        
        console_output.print_autocad_connection_status(True)
        
        # Открытие .dwg файла
        if not autocad_handler.open_dwg(file_path):
            console_output.print_error_message("Не удалось открыть .dwg файл")
            return False
        
        # Извлечение данных из AutoCAD - сначала ищем блоки "скважина"
        borehole_blocks = autocad_handler.find_borehole_blocks()

        # Если блоки найдены, используем их
        if borehole_blocks:
            boreholes = borehole_processor.extract_borehole_from_blocks(borehole_blocks)
        else:
            # Fallback: ищем по текстовым объектам и кругам
            logger.info("Блоки 'скважина' не найдены, ищем по текстовым объектам и кругам")
            text_entities = autocad_handler.find_text_entities()
            circles = autocad_handler.find_circles()

            if not text_entities and not circles:
                console_output.print_warning_message("В файле не найдено блоков скважин, текстовых объектов или кругов")
                return False

            boreholes = borehole_processor.extract_borehole_numbers(text_entities, circles)
        
        if not boreholes:
            console_output.print_warning_message("Не найдено скважин в файле")
            return False
        
        # Установка опорной скважины
        if not borehole_processor.set_reference_borehole(reference_borehole):
            console_output.print_error_message("Не удалось установить опорную скважину")
            return False
        
        # Расчет относительных высот
        if not borehole_processor.calculate_relative_heights(reference_z):
            console_output.print_error_message("Не удалось рассчитать относительные высоты")
            return False
        
        # Получение данных
        boreholes_data = borehole_processor.get_boreholes_data()
        reference_bh = borehole_processor.get_reference_borehole()
        
        # Преобразуем опорную скважину в словарь для вывода
        reference_bh_data = None
        if reference_bh:
            reference_bh_data = {
                'number': reference_bh.number,
                'x': reference_bh.x,
                'y': reference_bh.y,
                'z': reference_bh.z,
                'relative_height': reference_bh.relative_height
            }
        
        # Вывод результатов
        text_count = len(borehole_blocks) if borehole_blocks else 0
        circle_count = 0
        console_output.print_processing_stats(text_count, circle_count, len(boreholes))
        console_output.print_boreholes_summary(boreholes_data)
        console_output.print_reference_borehole_info(reference_bh_data)
        console_output.print_boreholes_table(boreholes_data)
        console_output.print_success_message(file_path)
        console_output.print_footer()
        
        return True
        
    except Exception as e:
        logger.error(f"Ошибка при обработке файла: {e}")
        return False
    
    finally:
        # Закрытие соединений
        try:
            autocad_handler.close_document()
            autocad_handler.disconnect()
        except Exception as e:
            logger.warning(f"Ошибка при закрытии соединений: {e}")


def main():
    """Основная функция программы."""
    parser = argparse.ArgumentParser(
        description="Определение номеров скважин в .dwg файлах и расчет относительных высот"
    )
    
    parser.add_argument(
        "dwg_file",
        help="Путь к .dwg файлу для обработки"
    )
    
    parser.add_argument(
        "-r", "--reference",
        help="Номер опорной скважины (если не указан, выбирается случайная)"
    )
    
    parser.add_argument(
        "-z", "--reference-z",
        type=float,
        default=0.0,
        help="Z-координата опорной скважины (по умолчанию: 0.0)"
    )
    
    
    parser.add_argument(
        "-l", "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Уровень логирования (по умолчанию: INFO)"
    )
    
    args = parser.parse_args()
    
    # Настройка логирования
    setup_logging(args.log_level)
    logger = logging.getLogger(__name__)
    
    # Проверка существования файла
    if not os.path.exists(args.dwg_file):
        logger.error(f"Файл не найден: {args.dwg_file}")
        sys.exit(1)
    
    # Обработка файла
    logger.info("Запуск обработки скважин...")
    success = process_dwg_file(
        file_path=args.dwg_file,
        reference_borehole=args.reference,
        reference_z=args.reference_z
    )
    
    if success:
        logger.info("Обработка завершена успешно!")
        sys.exit(0)
    else:
        logger.error("Обработка завершена с ошибками!")
        sys.exit(1)


if __name__ == "__main__":
    main()


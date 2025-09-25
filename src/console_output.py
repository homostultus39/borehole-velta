"""
Модуль для красивого вывода результатов в консоль.
"""

import logging
from typing import List, Dict, Any, Optional
from datetime import datetime

logger = logging.getLogger(__name__)


class ConsoleOutput:
    """Класс для вывода результатов в консоль."""
    
    def __init__(self):
        """Инициализация консольного вывода."""
        self.terminal_width = 80
    
    def print_header(self, title: str) -> None:
        """
        Вывод заголовка.
        
        Args:
            title: Заголовок для вывода
        """
        print("\n" + "=" * self.terminal_width)
        print(f" {title}".center(self.terminal_width))
        print("=" * self.terminal_width)
    
    def print_boreholes_summary(self, boreholes_data: List[Dict[str, Any]]) -> None:
        """
        Вывод сводки по скважинам.
        
        Args:
            boreholes_data: Данные о скважинах
        """
        if not boreholes_data:
            print("❌ Скважины не найдены")
            return
        
        # Находим опорную скважину
        reference_borehole = next((bh for bh in boreholes_data if bh.get('is_reference', False)), None)
        
        print(f"📊 Найдено скважин: {len(boreholes_data)}")
        if reference_borehole:
            print(f"🎯 Опорная скважина: №{reference_borehole['number']}")
        else:
            print("⚠️  Опорная скважина не установлена")
        
        # Статистика по высотам
        heights = [bh.get('relative_height', 0) for bh in boreholes_data if bh.get('relative_height') is not None]
        if heights:
            print(f"📏 Минимальная высота: {min(heights):.2f}")
            print(f"📏 Максимальная высота: {max(heights):.2f}")
            print(f"📏 Средняя высота: {sum(heights)/len(heights):.2f}")
    
    def print_boreholes_table(self, boreholes_data: List[Dict[str, Any]]) -> None:
        """
        Вывод таблицы со скважинами.
        
        Args:
            boreholes_data: Данные о скважинах
        """
        if not boreholes_data:
            return
        
        print("\n📋 Детальная информация о скважинах:")
        print("-" * self.terminal_width)
        
        # Заголовок таблицы
        header = f"{'№':<6} {'X':<10} {'Y':<10} {'Z':<10} {'Отн.высота':<12} {'Статус':<15}"
        print(header)
        print("-" * self.terminal_width)
        
        # Сортируем по номеру скважины
        sorted_boreholes = sorted(boreholes_data, key=lambda x: str(x['number']).zfill(10))
        
        for borehole in sorted_boreholes:
            number = borehole['number']
            x = f"{borehole['x']:.2f}" if borehole['x'] is not None else "Н/Д"
            y = f"{borehole['y']:.2f}" if borehole['y'] is not None else "Н/Д"
            z = f"{borehole['z']:.2f}" if borehole['z'] is not None else "Н/Д"
            rel_height = f"{borehole['relative_height']:.2f}" if borehole['relative_height'] is not None else "Н/Д"
            status = "🎯 Опорная" if borehole.get('is_reference', False) else "📌 Обычная"
            
            row = f"{number:<6} {x:<10} {y:<10} {z:<10} {rel_height:<12} {status:<15}"
            print(row)
        
        print("-" * self.terminal_width)
    
    def print_reference_borehole_info(self, reference_borehole: Optional[Dict[str, Any]]) -> None:
        """
        Вывод информации об опорной скважине.
        
        Args:
            reference_borehole: Данные об опорной скважине
        """
        if not reference_borehole:
            print("\n⚠️  Информация об опорной скважине недоступна")
            return
        
        print(f"\n🎯 Опорная скважина №{reference_borehole['number']}:")
        print(f"   📍 Координаты: X={reference_borehole['x']:.2f}, Y={reference_borehole['y']:.2f}")
        if reference_borehole['z'] is not None:
            print(f"   📏 Z-координата: {reference_borehole['z']:.2f}")
        print(f"   📊 Относительная высота: {reference_borehole['relative_height']:.2f}")
    
    def print_processing_stats(self, text_entities: int, circles: int, boreholes: int) -> None:
        """
        Вывод статистики обработки.
        
        Args:
            text_entities: Количество текстовых объектов
            circles: Количество кругов
            boreholes: Количество найденных скважин
        """
        print(f"\n📈 Статистика обработки:")
        print(f"   📝 Текстовых объектов: {text_entities}")
        print(f"   ⭕ Кругов: {circles}")
        print(f"   🕳️  Скважин найдено: {boreholes}")
    
    def print_success_message(self, file_path: str) -> None:
        """
        Вывод сообщения об успешной обработке.
        
        Args:
            file_path: Путь к обработанному файлу
        """
        print(f"\n✅ Файл успешно обработан: {file_path}")
        print(f"⏰ Время обработки: {datetime.now().strftime('%H:%M:%S')}")
    
    def print_error_message(self, error: str) -> None:
        """
        Вывод сообщения об ошибке.
        
        Args:
            error: Текст ошибки
        """
        print(f"\n❌ Ошибка: {error}")
    
    def print_warning_message(self, warning: str) -> None:
        """
        Вывод предупреждения.
        
        Args:
            warning: Текст предупреждения
        """
        print(f"\n⚠️  Предупреждение: {warning}")
    
    def print_help_info(self) -> None:
        """Вывод справочной информации."""
        print("\n💡 Справка:")
        print("   - Скрипт автоматически находит номера скважин в .dwg файлах")
        print("   - Первая найденная скважина становится опорной (относительная высота = 0)")
        print("   - Остальные скважины рассчитываются относительно опорной")
        print("   - Поддерживаемые форматы: 'скв. 123', '№ 123', '123 скв', 'скв 123', '123'")
    
    def print_file_info(self, file_path: str) -> None:
        """
        Вывод информации о файле.
        
        Args:
            file_path: Путь к файлу
        """
        print(f"\n📁 Обрабатываемый файл: {file_path}")
    
    def print_autocad_connection_status(self, connected: bool) -> None:
        """
        Вывод статуса подключения к AutoCAD.
        
        Args:
            connected: Статус подключения
        """
        if connected:
            print("🔗 Подключение к AutoCAD: ✅ Успешно")
        else:
            print("🔗 Подключение к AutoCAD: ❌ Ошибка")
    
    def print_footer(self) -> None:
        """Вывод подвала."""
        print("\n" + "=" * self.terminal_width)
        print(" Обработка завершена ".center(self.terminal_width))
        print("=" * self.terminal_width)

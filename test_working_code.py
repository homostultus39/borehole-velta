"""
Простой тест рабочего кода из второго коммита.
"""

import sys
import os

# Добавляем src в путь для импорта
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

def test_working_code():
    """Тест рабочего кода."""
    print("🔧 ТЕСТ РАБОЧЕГО КОДА ИЗ ВТОРОГО КОММИТА")
    print("=" * 60)
    
    # Тест 1: Проверка импорта
    print("\n1. Проверка импорта...")
    try:
        from src.autocad_handler import AutoCADHandler
        from src.borehole_processor import BoreholeProcessor
        from src.console_output import ConsoleOutput
        print("✅ Все модули импортированы успешно")
    except ImportError as e:
        print(f"❌ Ошибка импорта: {e}")
        return False
    
    # Тест 2: Создание объектов
    print("\n2. Создание объектов...")
    try:
        handler = AutoCADHandler()
        processor = BoreholeProcessor()
        console = ConsoleOutput()
        print("✅ Объекты созданы успешно")
    except Exception as e:
        print(f"❌ Ошибка создания объектов: {e}")
        return False
    
    # Тест 3: Попытка подключения к AutoCAD
    print("\n3. Попытка подключения к AutoCAD...")
    try:
        if handler.connect():
            print("✅ Подключение к AutoCAD успешно")
            
            # Проверяем доступ к объектам
            if handler.acad:
                print(f"✅ AutoCAD объект создан")
                if handler.doc:
                    print(f"✅ Активный документ: {handler.doc.Name}")
                else:
                    print("⚠️  Нет активного документа")
            
            # Отключаемся
            handler.disconnect()
            print("✅ Отключение успешно")
            return True
        else:
            print("❌ Не удалось подключиться к AutoCAD")
            return False
            
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        return False


def test_with_file():
    """Тест с реальным файлом."""
    print("\n📁 ТЕСТ С РЕАЛЬНЫМ ФАЙЛОМ")
    print("=" * 40)
    
    file_path = input("Введите путь к .dwg файлу (или нажмите Enter для пропуска): ").strip()
    
    if not file_path:
        print("⏭️  Тест с файлом пропущен")
        return True
    
    if not os.path.exists(file_path):
        print(f"❌ Файл не найден: {file_path}")
        return False
    
    try:
        from src.autocad_handler import AutoCADHandler
        
        handler = AutoCADHandler()
        
        if handler.connect():
            print("✅ AutoCAD подключен")
            
            if handler.open_dwg(file_path):
                print(f"✅ Файл открыт: {file_path}")
                
                # Получаем объекты
                entities = handler.get_all_entities()
                print(f"📊 Найдено объектов: {len(entities)}")
                
                # Ищем текстовые объекты
                text_entities = handler.find_text_entities()
                print(f"📝 Найдено текстовых объектов: {len(text_entities)}")
                
                # Ищем круги
                circles = handler.find_circles()
                print(f"⭕ Найдено кругов: {len(circles)}")
                
                # Закрываем файл
                handler.close_document()
                handler.disconnect()
                print("✅ Файл закрыт")
                
                return True
            else:
                print("❌ Не удалось открыть файл")
                return False
        else:
            print("❌ Не удалось подключиться к AutoCAD")
            return False
            
    except Exception as e:
        print(f"❌ Ошибка тестирования файла: {e}")
        return False


def main():
    """Основная функция тестирования."""
    print("🔧 ТЕСТИРОВАНИЕ РАБОЧЕГО КОДА")
    print("=" * 60)
    
    # Базовый тест
    if test_working_code():
        print("\n✅ Базовый тест прошел успешно!")
        
        # Тест с файлом
        if test_with_file():
            print("\n✅ Все тесты прошли успешно!")
            print("\n🎉 Рабочий код функционирует!")
        else:
            print("\n⚠️  Проблемы с файловыми операциями")
    else:
        print("\n❌ Базовый тест не прошел")
    
    print("\n" + "=" * 60)
    print("Тестирование завершено!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Тестирование прервано пользователем")
    except Exception as e:
        print(f"\n❌ Ошибка во время тестирования: {e}")
    
    input("\nНажмите Enter для выхода...")

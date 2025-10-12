"""
Тестовый скрипт для проверки улучшенного подключения к AutoCAD.
"""

import sys
import os
import logging

# Добавляем путь к src
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from autocad_handler import AutoCADHandler

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_improved_connection(file_path: str):
    """Тестирование улучшенного подключения к AutoCAD."""
    print("\n" + "="*60)
    print("🔧 ТЕСТИРОВАНИЕ УЛУЧШЕННОГО ПОДКЛЮЧЕНИЯ К AUTOCAD")
    print("="*60)
    
    handler = AutoCADHandler()
    
    # Тест 1: Подключение
    print("\n1️⃣ ТЕСТ ПОДКЛЮЧЕНИЯ")
    print("-" * 30)
    print("🔍 Попытка подключения с использованием:")
    print("   - DirectAutoCADConnector (AutoCAD.Application.25)")
    print("   - PyAutoCADConnector")
    print("   - Win32COMConnector")
    print("   - ComTypesConnector")
    
    if handler.connect():
        print("✅ Подключение успешно!")
        
        # Получаем информацию о подключении
        conn_info = handler.get_connection_info()
        print(f"   Метод подключения: {conn_info['method']}")
        print(f"   Подключен: {conn_info['connected']}")
        print(f"   Есть приложение: {conn_info['has_application']}")
        print(f"   Есть документ: {conn_info['has_document']}")
        
        # Тест 2: Открытие файла
        print("\n2️⃣ ТЕСТ ОТКРЫТИЯ ФАЙЛА")
        print("-" * 30)
        
        if handler.open_dwg(file_path):
            print(f"✅ Файл успешно открыт: {file_path}")
            
            # Тест 3: Поиск текстовых объектов
            print("\n3️⃣ ТЕСТ ПОИСКА ТЕКСТОВЫХ ОБЪЕКТОВ")
            print("-" * 30)
            
            text_entities = handler.find_text_entities()
            if text_entities:
                print(f"✅ Найдено {len(text_entities)} текстовых объектов")
                print("   Первые 3 объекта:")
                for i, entity in enumerate(text_entities[:3]):
                    print(f"   {i+1}. Текст: '{entity['text']}'")
                    print(f"      Позиция: {entity['position']}")
                    print(f"      Слой: {entity['layer']}")
            else:
                print("⚠️  Текстовые объекты не найдены")
            
            # Тест 4: Поиск кругов
            print("\n4️⃣ ТЕСТ ПОИСКА КРУГОВ")
            print("-" * 30)
            
            circles = handler.find_circles()
            if circles:
                print(f"✅ Найдено {len(circles)} кругов")
                print("   Первые 3 круга:")
                for i, circle in enumerate(circles[:3]):
                    print(f"   {i+1}. Центр: {circle['center']}")
                    print(f"      Радиус: {circle['radius']}")
                    print(f"      Слой: {circle['layer']}")
            else:
                print("⚠️  Круги не найдены")
            
            # Тест 5: Закрытие документа
            print("\n5️⃣ ТЕСТ ЗАКРЫТИЯ ДОКУМЕНТА")
            print("-" * 30)
            
            if handler.close_document():
                print("✅ Документ успешно закрыт")
            else:
                print("❌ Ошибка закрытия документа")
        
        else:
            print(f"❌ Не удалось открыть файл: {file_path}")
        
        # Тест 6: Отключение
        print("\n6️⃣ ТЕСТ ОТКЛЮЧЕНИЯ")
        print("-" * 30)
        
        if handler.disconnect():
            print("✅ Отключение успешно")
        else:
            print("❌ Ошибка отключения")
    
    else:
        print("❌ Не удалось подключиться к AutoCAD")
        print("\n💡 РЕКОМЕНДАЦИИ:")
        print("   1. Запустите диагностический скрипт: python diagnose_system.py")
        print("   2. Убедитесь, что AutoCAD запущен")
        print("   3. Проверьте, что установлена полная версия AutoCAD (не LT)")
        print("   4. Убедитесь, что Python и AutoCAD одинаковой разрядности")
        print("   5. Попробуйте запустить AutoCAD от имени администратора")
    
    print("\n" + "="*60)
    print("🏁 ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
    print("="*60)

def main():
    """Основная функция."""
    print("🔧 ТЕСТИРОВАНИЕ УЛУЧШЕННОГО ПОДКЛЮЧЕНИЯ К AUTOCAD")
    print("="*60)
    
    # Запрашиваем путь к файлу
    file_path = input("\nВведите полный путь к .dwg файлу для тестирования: ").strip()
    
    if not file_path:
        print("❌ Путь к файлу не указан. Тестирование отменено.")
        return
    
    if not os.path.exists(file_path):
        print(f"❌ Файл не найден по пути: {file_path}")
        return
    
    # Запускаем тестирование
    test_improved_connection(file_path)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Тестирование прервано пользователем")
    except Exception as e:
        print(f"\n❌ Ошибка во время тестирования: {e}")
        logger.exception("Детали ошибки:")
    
    input("\nНажмите Enter для выхода...")

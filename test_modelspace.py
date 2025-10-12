"""
Тестовый скрипт для проверки ModelSpace в AutoCAD.
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

def test_modelspace(file_path: str):
    """Тестирование ModelSpace в AutoCAD."""
    print("\n" + "="*60)
    print("🔧 ТЕСТИРОВАНИЕ MODELSPACE В AUTOCAD")
    print("="*60)
    
    handler = AutoCADHandler()
    
    # Подключение
    print("\n1️⃣ ПОДКЛЮЧЕНИЕ")
    print("-" * 30)
    
    if not handler.connect():
        print("❌ Не удалось подключиться к AutoCAD")
        return
    
    print("✅ Подключение успешно!")
    
    # Открытие файла
    print("\n2️⃣ ОТКРЫТИЕ ФАЙЛА")
    print("-" * 30)
    
    if not handler.open_dwg(file_path):
        print("❌ Не удалось открыть файл")
        handler.disconnect()
        return
    
    print("✅ Файл открыт!")
    
    # Тестирование ModelSpace
    print("\n3️⃣ ТЕСТИРОВАНИЕ MODELSPACE")
    print("-" * 30)
    
    try:
        # Получаем документ
        doc = handler.doc
        print(f"📄 Документ: {type(doc)}")
        
        # Показываем информацию о документе
        try:
            doc_name = getattr(doc, 'Name', 'Unknown')
            print(f"📄 Имя документа: {doc_name}")
        except Exception as e:
            print(f"⚠️ Не удалось получить имя документа: {e}")
        
        # Получаем ModelSpace разными способами
        model_space = None
        
        # Способ 1: Прямое обращение
        try:
            model_space = doc.ModelSpace
            print(f"📋 ModelSpace (способ 1): {type(model_space)}")
        except Exception as e:
            print(f"❌ Способ 1 не сработал: {e}")
            
            # Способ 2: Через ActiveDocument
            try:
                model_space = handler.acad.ActiveDocument.ModelSpace
                print(f"📋 ModelSpace (способ 2): {type(model_space)}")
            except Exception as e:
                print(f"❌ Способ 2 не сработал: {e}")
                
                # Способ 3: Через Documents коллекцию
                try:
                    model_space = handler.acad.Documents.Item(0).ModelSpace
                    print(f"📋 ModelSpace (способ 3): {type(model_space)}")
                except Exception as e:
                    print(f"❌ Способ 3 не сработал: {e}")
                    print("❌ Все способы получения ModelSpace не сработали")
                    return
        
        if model_space is None:
            print("❌ ModelSpace не получен ни одним способом")
            return
        
        # Пытаемся получить количество объектов
        try:
            count = model_space.Count
            print(f"📊 Количество объектов в ModelSpace: {count}")
        except Exception as e:
            print(f"⚠️ Не удалось получить количество объектов: {e}")
        
        # Пытаемся итерироваться по объектам
        print("\n4️⃣ ИТЕРАЦИЯ ПО ОБЪЕКТАМ")
        print("-" * 30)
        
        entity_count = 0
        for entity in model_space:
            entity_count += 1
            if entity_count <= 5:  # Показываем первые 5
                print(f"🔸 Объект {entity_count}: {type(entity)}")
                try:
                    entity_name = getattr(entity, 'EntityName', 'Unknown')
                    print(f"   Имя: {entity_name}")
                except Exception as e:
                    print(f"   Ошибка получения имени: {e}")
            
            if entity_count >= 10:  # Ограничиваем для теста
                print(f"📊 Обработано {entity_count} объектов (ограничено для теста)")
                break
        
        print(f"✅ Итерация завершена! Всего обработано: {entity_count}")
        
    except Exception as e:
        print(f"❌ Ошибка при работе с ModelSpace: {e}")
        logger.exception("Детали ошибки:")
    
    # Закрытие
    print("\n5️⃣ ЗАКРЫТИЕ")
    print("-" * 30)
    
    handler.close_document()
    handler.disconnect()
    print("✅ Закрытие завершено")
    
    print("\n" + "="*60)
    print("🏁 ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
    print("="*60)

def main():
    """Основная функция."""
    print("🔧 ТЕСТИРОВАНИЕ MODELSPACE В AUTOCAD")
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
    test_modelspace(file_path)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Тестирование прервано пользователем")
    except Exception as e:
        print(f"\n❌ Ошибка во время тестирования: {e}")
        logger.exception("Детали ошибки:")
    
    input("\nНажмите Enter для выхода...")

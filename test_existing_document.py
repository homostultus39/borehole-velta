"""
Тестовый скрипт для работы с уже открытым документом в AutoCAD.
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

def test_existing_document():
    """Тестирование работы с уже открытым документом."""
    print("\n" + "="*60)
    print("🔧 ТЕСТИРОВАНИЕ РАБОТЫ С УЖЕ ОТКРЫТЫМ ДОКУМЕНТОМ")
    print("="*60)
    
    handler = AutoCADHandler()
    
    # Подключение
    print("\n1️⃣ ПОДКЛЮЧЕНИЕ")
    print("-" * 30)
    
    if not handler.connect():
        print("❌ Не удалось подключиться к AutoCAD")
        return
    
    print("✅ Подключение успешно!")
    
    # Проверяем текущий документ
    print("\n2️⃣ ПРОВЕРКА ТЕКУЩЕГО ДОКУМЕНТА")
    print("-" * 30)
    
    try:
        # Получаем текущий документ
        current_doc = handler.acad.ActiveDocument
        if current_doc:
            doc_name = getattr(current_doc, 'Name', 'Unknown')
            print(f"📄 Текущий документ: {doc_name}")
            print(f"📄 Тип документа: {type(current_doc)}")
            
            # Устанавливаем документ в обработчик
            handler.doc = current_doc
            print("✅ Документ установлен в обработчик")
        else:
            print("❌ Нет активного документа")
            return
    except Exception as e:
        print(f"❌ Ошибка получения текущего документа: {e}")
        return
    
    # Тестирование ModelSpace
    print("\n3️⃣ ТЕСТИРОВАНИЕ MODELSPACE")
    print("-" * 30)
    
    try:
        # Пытаемся получить ModelSpace
        model_space = current_doc.ModelSpace
        print(f"📋 ModelSpace: {type(model_space)}")
        
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
        text_count = 0
        circle_count = 0
        
        for entity in model_space:
            entity_count += 1
            
            # Проверяем тип объекта
            try:
                entity_name = getattr(entity, 'EntityName', 'Unknown')
                if entity_name in ['AcDbText', 'AcDbMText']:
                    text_count += 1
                    if text_count <= 3:  # Показываем первые 3 текста
                        text_content = getattr(entity, 'TextString', '')
                        print(f"📝 Текст {text_count}: '{text_content}'")
                elif entity_name == 'AcDbCircle':
                    circle_count += 1
                    if circle_count <= 3:  # Показываем первые 3 круга
                        center = getattr(entity, 'Center', (0, 0, 0))
                        radius = getattr(entity, 'Radius', 0)
                        print(f"⭕ Круг {circle_count}: центр {center}, радиус {radius}")
            except Exception as e:
                # Пропускаем проблемные объекты
                continue
            
            if entity_count >= 50:  # Ограничиваем для теста
                print(f"📊 Обработано {entity_count} объектов (ограничено для теста)")
                break
        
        print(f"✅ Итерация завершена!")
        print(f"📊 Всего объектов: {entity_count}")
        print(f"📝 Текстовых объектов: {text_count}")
        print(f"⭕ Кругов: {circle_count}")
        
    except Exception as e:
        print(f"❌ Ошибка при работе с ModelSpace: {e}")
        logger.exception("Детали ошибки:")
    
    # Закрытие
    print("\n5️⃣ ЗАКРЫТИЕ")
    print("-" * 30)
    
    handler.disconnect()
    print("✅ Отключение завершено")
    
    print("\n" + "="*60)
    print("🏁 ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
    print("="*60)

def main():
    """Основная функция."""
    print("🔧 ТЕСТИРОВАНИЕ РАБОТЫ С УЖЕ ОТКРЫТЫМ ДОКУМЕНТОМ")
    print("="*60)
    print("💡 Убедитесь, что в AutoCAD уже открыт нужный .dwg файл!")
    
    input("\nНажмите Enter, когда будете готовы...")
    
    # Запускаем тестирование
    test_existing_document()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Тестирование прервано пользователем")
    except Exception as e:
        print(f"\n❌ Ошибка во время тестирования: {e}")
        logger.exception("Детали ошибки:")
    
    input("\nНажмите Enter для выхода...")

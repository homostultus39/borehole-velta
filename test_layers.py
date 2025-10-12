"""
Тестовый скрипт для проверки слоев в AutoCAD.
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

def test_layers():
    """Тестирование слоев в AutoCAD."""
    print("\n" + "="*60)
    print("🔧 ТЕСТИРОВАНИЕ СЛОЕВ В AUTOCAD")
    print("="*60)
    
    handler = AutoCADHandler()
    
    # Подключение
    print("\n1️⃣ ПОДКЛЮЧЕНИЕ")
    print("-" * 30)
    
    if not handler.connect():
        print("❌ Не удалось подключиться к AutoCAD")
        return
    
    print("✅ Подключение успешно!")
    
    # Получение информации о слоях
    print("\n2️⃣ ПОЛУЧЕНИЕ ИНФОРМАЦИИ О СЛОЯХ")
    print("-" * 30)
    
    layers_info = handler.get_layers_info()
    
    if not layers_info:
        print("❌ Не удалось получить информацию о слоях")
        handler.disconnect()
        return
    
    print(f"✅ Получена информация о {len(layers_info)} слоях")
    
    # Поиск слоев со скважинами
    print("\n3️⃣ ПОИСК СЛОЕВ СО СКВАЖИНАМИ")
    print("-" * 30)
    
    borehole_layers = []
    for layer in layers_info:
        layer_name = layer['name'].upper()
        if 'СКВ' in layer_name:
            borehole_layers.append(layer)
    
    if borehole_layers:
        print(f"✅ Найдено {len(borehole_layers)} слоев со скважинами:")
        for layer in borehole_layers:
            print(f"   📋 {layer['name']} (видимый: {layer['visible']}, заблокирован: {layer['locked']})")
    else:
        print("⚠️  Слои со скважинами не найдены")
    
    # Показываем все слои (первые 20)
    print("\n4️⃣ ВСЕ СЛОИ (первые 20)")
    print("-" * 30)
    
    for i, layer in enumerate(layers_info[:20]):
        status = "✅" if layer['visible'] else "❌"
        locked = "🔒" if layer['locked'] else "🔓"
        print(f"   {status} {locked} {layer['name']}")
    
    if len(layers_info) > 20:
        print(f"   ... и еще {len(layers_info) - 20} слоев")
    
    # Тестирование поиска по слоям
    print("\n5️⃣ ТЕСТИРОВАНИЕ ПОИСКА ПО СЛОЯМ")
    print("-" * 30)
    
    # Поиск текста на слоях со скважинами
    text_entities = handler.find_text_entities("СКВ")
    print(f"📝 Найдено {len(text_entities)} текстовых объектов на слоях со скважинами")
    
    # Поиск кругов на слоях со скважинами
    circles = handler.find_circles("СКВ")
    print(f"⭕ Найдено {len(circles)} кругов на слоях со скважинами")
    
    # Показываем первые найденные объекты
    if text_entities:
        print("\n📝 Первые 5 текстовых объектов:")
        for i, text in enumerate(text_entities[:5]):
            print(f"   {i+1}. Слой '{text['layer']}': '{text['text']}'")
    
    if circles:
        print("\n⭕ Первые 5 кругов:")
        for i, circle in enumerate(circles[:5]):
            print(f"   {i+1}. Слой '{circle['layer']}': центр {circle['center']}, радиус {circle['radius']}")
    
    # Закрытие
    print("\n6️⃣ ЗАКРЫТИЕ")
    print("-" * 30)
    
    handler.disconnect()
    print("✅ Отключение завершено")
    
    print("\n" + "="*60)
    print("🏁 ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
    print("="*60)

def main():
    """Основная функция."""
    print("🔧 ТЕСТИРОВАНИЕ СЛОЕВ В AUTOCAD")
    print("="*60)
    print("💡 Убедитесь, что в AutoCAD уже открыт нужный .dwg файл!")
    
    input("\nНажмите Enter, когда будете готовы...")
    
    # Запускаем тестирование
    test_layers()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Тестирование прервано пользователем")
    except Exception as e:
        print(f"\n❌ Ошибка во время тестирования: {e}")
        logger.exception("Детали ошибки:")
    
    input("\nНажмите Enter для выхода...")

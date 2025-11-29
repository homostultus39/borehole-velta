"""
Диагностический скрипт для анализа блоков в AutoCAD.
Помогает понять структуру блоков "скважина" в вашем проекте.
"""

import win32com.client
import sys

def diagnose_autocad_blocks(dwg_path=None):
    """
    Диагностика блоков в AutoCAD документе.

    Args:
        dwg_path: Путь к .dwg файлу (опционально)
    """
    try:
        # Подключение к AutoCAD
        print("=" * 80)
        print("ДИАГНОСТИКА БЛОКОВ AUTOCAD")
        print("=" * 80)

        acad = win32com.client.Dispatch("AutoCAD.Application")
        print(f"✅ Подключено к AutoCAD версии: {acad.Version}")

        # Пробуем получить активный документ
        try:
            doc = acad.ActiveDocument
            print(f"✅ Используется активный документ: {doc.Name}")

            if dwg_path:
                print(f"⚠️ Файл указан ({dwg_path}), но используется уже открытый документ")
                print(f"   Пожалуйста, откройте нужный файл в AutoCAD вручную перед запуском скрипта")
        except Exception as e:
            print(f"❌ Не удалось получить активный документ: {e}")
            print("   Откройте нужный .dwg файл в AutoCAD и запустите скрипт снова")
            return

        print("\n" + "=" * 80)
        print("АНАЛИЗ БЛОКОВ В MODELSPACE")
        print("=" * 80)

        # Статистика
        total_entities = 0
        block_references = 0
        blocks_by_name = {}
        blocks_by_layer = {}
        blocks_with_attributes = 0
        sample_blocks = []

        # Проходим по всем объектам
        for entity in doc.ModelSpace:
            total_entities += 1

            if entity.EntityName == 'AcDbBlockReference':
                block_references += 1

                # Получаем информацию о блоке
                block_name = getattr(entity, 'Name', 'Unknown')
                effective_name = getattr(entity, 'EffectiveName', block_name)
                layer = getattr(entity, 'Layer', 'Unknown')
                has_attributes = getattr(entity, 'HasAttributes', False)

                # Статистика по именам
                if effective_name not in blocks_by_name:
                    blocks_by_name[effective_name] = 0
                blocks_by_name[effective_name] += 1

                # Статистика по слоям
                if layer not in blocks_by_layer:
                    blocks_by_layer[layer] = []
                blocks_by_layer[layer].append(effective_name)

                # Считаем блоки с атрибутами
                if has_attributes:
                    blocks_with_attributes += 1

                # Собираем примеры блоков "скважина"
                if 'скважина' in effective_name.lower() and len(sample_blocks) < 5:
                    insertion_point = entity.InsertionPoint

                    attributes_info = {}
                    if has_attributes:
                        try:
                            attrs = entity.GetAttributes()
                            for attr in attrs:
                                tag = getattr(attr, 'TagString', '')
                                value = getattr(attr, 'TextString', '')
                                attributes_info[tag] = value
                        except Exception as e:
                            attributes_info = {"error": str(e)}

                    sample_blocks.append({
                        'name': effective_name,
                        'layer': layer,
                        'position': (insertion_point[0], insertion_point[1], insertion_point[2]),
                        'has_attributes': has_attributes,
                        'attributes': attributes_info
                    })

            # Прогресс каждые 10000 объектов
            if total_entities % 10000 == 0:
                print(f"Обработано {total_entities} объектов...")

        # Выводим результаты
        print(f"\n📊 ОБЩАЯ СТАТИСТИКА:")
        print(f"   Всего объектов: {total_entities}")
        print(f"   Блоков (AcDbBlockReference): {block_references}")
        print(f"   Блоков с атрибутами: {blocks_with_attributes}")

        print(f"\n📦 ТИПЫ БЛОКОВ (топ-20):")
        sorted_blocks = sorted(blocks_by_name.items(), key=lambda x: x[1], reverse=True)
        for name, count in sorted_blocks[:20]:
            marker = "⭐" if 'скважина' in name.lower() else "  "
            print(f"   {marker} '{name}': {count} вставок")

        print(f"\n🗂️ БЛОКИ ПО СЛОЯМ (только слои с 'СКВ'):")
        for layer, block_names in sorted(blocks_by_layer.items()):
            if 'СКВ' in layer.upper():
                unique_blocks = set(block_names)
                print(f"   Слой '{layer}':")
                for block_name in unique_blocks:
                    count = block_names.count(block_name)
                    print(f"      - '{block_name}': {count} вставок")

        print(f"\n🔍 ПРИМЕРЫ БЛОКОВ 'СКВАЖИНА':")
        if sample_blocks:
            for i, block in enumerate(sample_blocks, 1):
                print(f"\n   Пример #{i}:")
                print(f"      Имя: {block['name']}")
                print(f"      Слой: {block['layer']}")
                print(f"      Позиция: ({block['position'][0]:.2f}, {block['position'][1]:.2f}, {block['position'][2]:.2f})")
                print(f"      Есть атрибуты: {block['has_attributes']}")
                if block['attributes']:
                    print(f"      Атрибуты:")
                    for tag, value in block['attributes'].items():
                        print(f"         {tag}: {value}")
        else:
            print("   ⚠️ Не найдено блоков с именем 'скважина'")

        print("\n" + "=" * 80)
        print("РЕКОМЕНДАЦИИ:")
        print("=" * 80)

        # Анализ и рекомендации
        skvazhina_blocks = {name: count for name, count in blocks_by_name.items()
                           if 'скважина' in name.lower()}

        if not skvazhina_blocks:
            print("❌ Не найдено блоков с именем 'скважина'")
            print("   Проверьте:")
            print("   1. Правильное ли имя блока в вашем проекте?")
            print("   2. Может быть используются другие имена (например, 'well', 'borehole', 'СКВ')?")
            print("\n   Посмотрите на список 'ТИПЫ БЛОКОВ' выше и найдите правильное имя")
        else:
            print(f"✅ Найдено блоков 'скважина': {sum(skvazhina_blocks.values())} вставок")
            print(f"   Варианты имен: {list(skvazhina_blocks.keys())}")

            if sample_blocks and sample_blocks[0]['has_attributes']:
                print(f"\n✅ Блоки имеют атрибуты:")
                if sample_blocks[0]['attributes']:
                    print(f"   Теги атрибутов: {list(sample_blocks[0]['attributes'].keys())}")
                    print("   Используйте эти теги для извлечения номеров скважин")
            else:
                print(f"\n⚠️ Блоки НЕ имеют атрибутов")
                print("   Возможно, номера хранятся в других объектах (текст рядом с блоком)")

        print("\n" + "=" * 80)

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    dwg_path = sys.argv[1] if len(sys.argv) > 1 else None
    diagnose_autocad_blocks(dwg_path)

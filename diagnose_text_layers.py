"""
Поиск текстовых объектов на слоях со "СКВ" и рядом с блоками.
"""

import win32com.client
import re

def diagnose_text_and_layers():
    """Анализ текстов на слоях со СКВ."""
    print("=" * 80)
    print("АНАЛИЗ ТЕКСТОВЫХ ОБЪЕКТОВ И СЛОЕВ СО 'СКВ'")
    print("=" * 80)

    acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument

    print(f"Документ: {doc.Name}\n")

    # Собираем все слои со "СКВ"
    print("ШАГ 1: ПОИСК СЛОЕВ СО 'СКВ'")
    print("-" * 80)

    skv_layers = []
    layers = doc.Layers
    for i in range(layers.Count):
        layer = layers.Item(i)
        layer_name = layer.Name
        if 'СКВ' in layer_name.upper():
            skv_layers.append(layer_name)
            print(f"   ✓ {layer_name}")

    print(f"\nНайдено {len(skv_layers)} слоев со 'СКВ'\n")

    # Ищем текстовые объекты на этих слоях
    print("ШАГ 2: ПОИСК ТЕКСТОВ НА СЛОЯХ СО 'СКВ'")
    print("-" * 80)

    text_objects = []
    blocks_on_skv_layers = []

    model_space = doc.ModelSpace
    processed = 0

    for entity in model_space:
        processed += 1
        if processed % 10000 == 0:
            print(f"Обработано {processed} объектов...")

        try:
            entity_layer = getattr(entity, 'Layer', None)
            if not entity_layer:
                continue

            # Только объекты на слоях со СКВ
            if not any(skv in entity_layer.upper() for skv in ['СКВ']):
                continue

            entity_type = entity.EntityName

            # Текстовые объекты
            if hasattr(entity, 'TextString'):
                text = entity.TextString
                pos = entity.InsertionPoint

                text_objects.append({
                    'text': text,
                    'layer': entity_layer,
                    'position': (pos[0], pos[1], pos[2]),
                    'type': entity_type
                })

            # Блоки
            elif entity_type == 'AcDbBlockReference':
                name = getattr(entity, 'EffectiveName', getattr(entity, 'Name', 'Unknown'))
                pos = entity.InsertionPoint
                has_attrs = getattr(entity, 'HasAttributes', False)

                attrs = {}
                if has_attrs:
                    try:
                        for attr in entity.GetAttributes():
                            attrs[attr.TagString] = attr.TextString
                    except:
                        pass

                blocks_on_skv_layers.append({
                    'name': name,
                    'layer': entity_layer,
                    'position': (pos[0], pos[1], pos[2]),
                    'has_attrs': has_attrs,
                    'attrs': attrs
                })

        except:
            continue

    print(f"\n✅ Найдено {len(text_objects)} текстовых объектов")
    print(f"✅ Найдено {len(blocks_on_skv_layers)} блоков")

    # Анализируем текстовые объекты
    print("\n" + "=" * 80)
    print("ШАГ 3: АНАЛИЗ ТЕКСТОВ (ПОИСК НОМЕРОВ СКВАЖИН)")
    print("=" * 80)

    patterns = [
        r'скв[а-я]*\.?\s*(\d+)',
        r'№\s*(\d+)',
        r'(\d+)\s*скв',
        r'скв\s*(\d+)',
        r'^(\d+)$',
    ]

    potential_boreholes = []

    for text_obj in text_objects[:100]:  # Первые 100 для примера
        text = text_obj['text'].strip().lower()

        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                number = match.group(1)
                potential_boreholes.append({
                    'number': number,
                    'text': text_obj['text'],
                    'layer': text_obj['layer'],
                    'position': text_obj['position']
                })
                break

    print(f"Найдено {len(potential_boreholes)} потенциальных номеров скважин в текстах\n")

    if potential_boreholes:
        print("Примеры (первые 10):")
        for i, bh in enumerate(potential_boreholes[:10], 1):
            print(f"\n   #{i}: Номер '{bh['number']}'")
            print(f"        Текст: '{bh['text']}'")
            print(f"        Слой: {bh['layer']}")
            print(f"        Позиция: ({bh['position'][0]:.2f}, {bh['position'][1]:.2f}, {bh['position'][2]:.2f})")

    # Анализируем блоки на слоях СКВ
    print("\n" + "=" * 80)
    print("ШАГ 4: АНАЛИЗ БЛОКОВ НА СЛОЯХ СКВ")
    print("=" * 80)

    blocks_by_name = {}
    for block in blocks_on_skv_layers:
        name = block['name']
        if name not in blocks_by_name:
            blocks_by_name[name] = 0
        blocks_by_name[name] += 1

    print(f"Типы блоков (топ-20):")
    for name, count in sorted(blocks_by_name.items(), key=lambda x: x[1], reverse=True)[:20]:
        print(f"   '{name}': {count} вставок")

    # Показываем примеры блоков с атрибутами
    blocks_with_attrs = [b for b in blocks_on_skv_layers if b['has_attrs'] and b['attrs']]

    if blocks_with_attrs:
        print(f"\n📌 Блоки с атрибутами (первые 5):")
        for i, block in enumerate(blocks_with_attrs[:5], 1):
            print(f"\n   #{i}: Блок '{block['name']}'")
            print(f"        Слой: {block['layer']}")
            print(f"        Позиция: ({block['position'][0]:.2f}, {block['position'][1]:.2f}, {block['position'][2]:.2f})")
            print(f"        Атрибуты:")
            for tag, val in block['attrs'].items():
                print(f"           {tag}: {val}")

    # Рекомендации
    print("\n" + "=" * 80)
    print("ВЫВОДЫ И РЕКОМЕНДАЦИИ")
    print("=" * 80)

    if potential_boreholes:
        print(f"✅ Найдено {len(potential_boreholes)} номеров скважин в ТЕКСТОВЫХ объектах")
        print(f"   Скважины представлены текстами, а не блоками 'скважина'")
        print(f"\n💡 РЕШЕНИЕ:")
        print(f"   Нужно использовать ТЕКСТОВЫЕ объекты для определения скважин,")
        print(f"   а не блоки. Возможно, рядом с текстами есть круги или другие маркеры.")

    if blocks_on_skv_layers:
        print(f"\n✅ Найдено {len(blocks_on_skv_layers)} блоков на слоях со 'СКВ'")
        most_common_block = max(blocks_by_name.items(), key=lambda x: x[1])[0]
        print(f"   Самый частый блок: '{most_common_block}' ({blocks_by_name[most_common_block]} вставок)")
        print(f"\n💡 Возможно, скважины обозначены блоком '{most_common_block}'")

    print("\n" + "=" * 80)


if __name__ == "__main__":
    diagnose_text_and_layers()

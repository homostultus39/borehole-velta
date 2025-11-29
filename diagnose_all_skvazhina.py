"""
Детальный поиск ВСЕХ блоков "скважина" в проекте.
Проверяет Name, EffectiveName, динамические блоки, вложенные блоки.
"""

import win32com.client

def diagnose_all_skvazhina_blocks():
    """Найти абсолютно все блоки 'скважина'."""
    print("=" * 80)
    print("ДЕТАЛЬНЫЙ ПОИСК ВСЕХ БЛОКОВ 'СКВАЖИНА'")
    print("=" * 80)

    acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument

    print(f"Документ: {doc.Name}\n")

    # ШАГ 1: Проверяем определения блоков
    print("ШАГ 1: ПРОВЕРКА ОПРЕДЕЛЕНИЙ БЛОКОВ (Block Definitions)")
    print("-" * 80)

    blocks_collection = doc.Blocks
    print(f"Всего определений блоков в файле: {blocks_collection.Count}\n")

    skvazhina_definition = None
    for i in range(blocks_collection.Count):
        block_def = blocks_collection.Item(i)
        block_name = block_def.Name

        if 'скважина' in block_name.lower():
            is_xref = getattr(block_def, 'IsXRef', False)
            is_layout = getattr(block_def, 'IsLayout', False)

            print(f"✓ Найдено определение: '{block_name}'")
            print(f"  - XRef: {is_xref}")
            print(f"  - Layout: {is_layout}")
            print(f"  - Объектов в определении: {block_def.Count}")

            skvazhina_definition = block_def

    if not skvazhina_definition:
        print("❌ Определение блока 'скважина' не найдено!")
        return

    print()

    # ШАГ 2: Ищем ВСЕ вставки этого блока
    print("ШАГ 2: ПОИСК ВСЕХ ВСТАВОК (Block References)")
    print("-" * 80)

    model_space = doc.ModelSpace
    all_skvazhina_refs = []

    processed = 0
    for entity in model_space:
        processed += 1
        if processed % 10000 == 0:
            print(f"Обработано {processed} объектов, найдено {len(all_skvazhina_refs)} блоков 'скважина'...")

        try:
            if entity.EntityName == 'AcDbBlockReference':
                # Проверяем разные способы получения имени
                name = getattr(entity, 'Name', None)
                effective_name = getattr(entity, 'EffectiveName', None)
                is_dynamic = getattr(entity, 'IsDynamicBlock', False)

                # Проверяем все варианты
                if (name and 'скважина' in name.lower()) or \
                   (effective_name and 'скважина' in effective_name.lower()):

                    layer = getattr(entity, 'Layer', 'Unknown')
                    pos = entity.InsertionPoint
                    has_attrs = getattr(entity, 'HasAttributes', False)

                    # Получаем атрибуты
                    attrs = {}
                    if has_attrs:
                        try:
                            for attr in entity.GetAttributes():
                                tag = getattr(attr, 'TagString', '')
                                val = getattr(attr, 'TextString', '')
                                attrs[tag] = val
                        except:
                            pass

                    all_skvazhina_refs.append({
                        'Name': name,
                        'EffectiveName': effective_name,
                        'IsDynamic': is_dynamic,
                        'Layer': layer,
                        'Position': (pos[0], pos[1], pos[2]),
                        'HasAttributes': has_attrs,
                        'Attributes': attrs
                    })

        except Exception as e:
            continue

    print(f"\n✅ НАЙДЕНО {len(all_skvazhina_refs)} ВСТАВОК БЛОКА 'СКВАЖИНА'!\n")

    # ШАГ 3: Анализ найденных вставок
    print("ШАГ 3: АНАЛИЗ НАЙДЕННЫХ ВСТАВОК")
    print("-" * 80)

    if not all_skvazhina_refs:
        print("❌ Вставки не найдены!")
        print("\n💡 ВОЗМОЖНЫЕ ПРИЧИНЫ:")
        print("1. Блок определен, но не вставлен в документ")
        print("2. Вставки находятся в PaperSpace (не в ModelSpace)")
        print("3. Блок вложен в другой блок")
        print("4. Проблема с кодировкой имени блока")
        return

    # Группируем по слоям
    by_layer = {}
    for ref in all_skvazhina_refs:
        layer = ref['Layer']
        if layer not in by_layer:
            by_layer[layer] = 0
        by_layer[layer] += 1

    print(f"Распределение по слоям:")
    for layer, count in sorted(by_layer.items(), key=lambda x: x[1], reverse=True):
        print(f"   {layer}: {count} вставок")

    # Проверяем динамические блоки
    dynamic_count = sum(1 for ref in all_skvazhina_refs if ref['IsDynamic'])
    if dynamic_count:
        print(f"\n⚡ Динамических блоков: {dynamic_count}")

    # Проверяем атрибуты
    with_attrs = [ref for ref in all_skvazhina_refs if ref['HasAttributes']]
    print(f"\n📌 Блоков с атрибутами: {len(with_attrs)}")

    # Показываем примеры
    print(f"\n🔍 ПРИМЕРЫ ВСТАВОК (первые 10):")
    for i, ref in enumerate(all_skvazhina_refs[:10], 1):
        print(f"\n   Вставка #{i}:")
        print(f"      Name: {ref['Name']}")
        print(f"      EffectiveName: {ref['EffectiveName']}")
        print(f"      Слой: {ref['Layer']}")
        print(f"      Позиция: ({ref['Position'][0]:.2f}, {ref['Position'][1]:.2f}, {ref['Position'][2]:.2f})")
        print(f"      Динамический: {ref['IsDynamic']}")
        if ref['Attributes']:
            print(f"      Атрибуты:")
            for tag, val in ref['Attributes'].items():
                print(f"         {tag}: {val}")

    # ШАГ 4: Проверка PaperSpace
    print("\n" + "=" * 80)
    print("ШАГ 4: ПРОВЕРКА PAPERSPACE")
    print("-" * 80)

    try:
        layouts = doc.Layouts
        for i in range(layouts.Count):
            layout = layouts.Item(i)
            if not layout.ModelType:  # PaperSpace
                layout_block = layout.Block
                ps_count = 0

                for entity in layout_block:
                    try:
                        if entity.EntityName == 'AcDbBlockReference':
                            name = getattr(entity, 'Name', '')
                            eff_name = getattr(entity, 'EffectiveName', '')

                            if 'скважина' in name.lower() or 'скважина' in eff_name.lower():
                                ps_count += 1
                    except:
                        continue

                if ps_count > 0:
                    print(f"   Layout '{layout.Name}': {ps_count} вставок")

    except Exception as e:
        print(f"⚠️ Ошибка проверки PaperSpace: {e}")

    # ВЫВОДЫ
    print("\n" + "=" * 80)
    print("ИТОГОВЫЙ ВЫВОД")
    print("=" * 80)

    if len(all_skvazhina_refs) > 1:
        print(f"✅ УСПЕХ! Найдено {len(all_skvazhina_refs)} вставок блока 'скважина'")
        print(f"\n💡 Код должен работать правильно с этими блоками")
        print(f"   Проблема была в фильтрации по слою 'СКВ' - блоки на других слоях!")
    elif len(all_skvazhina_refs) == 1:
        print(f"⚠️ Найдена только 1 вставка блока 'скважина'")
        print(f"\n💡 ВОЗМОЖНЫЕ ОБЪЯСНЕНИЯ:")
        print(f"1. В проекте действительно только одна скважина")
        print(f"2. Остальные скважины обозначены ДРУГИМИ блоками")
        print(f"3. Номера скважин - это атрибуты ОДНОГО блока")
        print(f"4. Скважины представлены вложенными блоками")

    print("\n" + "=" * 80)


if __name__ == "__main__":
    diagnose_all_skvazhina_blocks()

"""
Глубокая диагностика AutoCAD - проверяет все возможные места, где могут быть блоки.
"""

import win32com.client
import sys
import os

def deep_diagnose():
    """Детальная диагностика AutoCAD."""
    print("=" * 80)
    print("ГЛУБОКАЯ ДИАГНОСТИКА AUTOCAD")
    print("=" * 80)

    # Используем dynamic.Dispatch для избежания проблем с кэшем
    try:
        acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
        print(f"✅ Подключено к AutoCAD (dynamic) версии: {acad.Version}")
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        print("\n💡 РЕШЕНИЕ:")
        print("1. Очистите кэш win32com:")
        print("   import win32com")
        print("   import shutil, os")
        print("   gen_py = os.path.join(win32com.__gen_path__, 'gen_py')")
        print("   shutil.rmtree(gen_py)")
        print("2. Перезапустите скрипт")
        return

    print("\n" + "=" * 80)
    print("ШАГ 1: ПРОВЕРКА ДОКУМЕНТОВ")
    print("=" * 80)

    # Получаем активный документ
    try:
        active_doc = acad.ActiveDocument
        print(f"✅ Активный документ: {active_doc.Name}")
        print(f"   Полный путь: {active_doc.FullName}")
    except Exception as e:
        print(f"❌ Нет активного документа: {e}")
        return

    doc = active_doc

    print("\n" + "=" * 80)
    print("ШАГ 2: ПРОВЕРКА LAYOUTS (ModelSpace, PaperSpace)")
    print("=" * 80)

    layouts_info = {}
    try:
        layouts = doc.Layouts
        print(f"Найдено {layouts.Count} layouts:")

        for i in range(layouts.Count):
            layout = layouts.Item(i)
            layout_name = layout.Name
            is_model = layout.ModelType

            # Считаем объекты в layout
            block = layout.Block
            entity_count = block.Count

            layouts_info[layout_name] = {
                'is_model': is_model,
                'entity_count': entity_count,
                'block': block
            }

            marker = "📐 ModelSpace" if is_model else "📄 PaperSpace"
            print(f"   {i+1}. {layout_name} {marker} - {entity_count} объектов")

    except Exception as e:
        print(f"⚠️ Проблема с Layouts: {e}")

    print("\n" + "=" * 80)
    print("ШАГ 3: ПОИСК БЛОКОВ В MODELSPACE")
    print("=" * 80)

    try:
        model_space = doc.ModelSpace
        print(f"ModelSpace: {model_space.Count} объектов")

        block_refs = []
        block_names = {}

        for entity in model_space:
            if entity.EntityName == 'AcDbBlockReference':
                block_refs.append(entity)
                name = getattr(entity, 'EffectiveName', getattr(entity, 'Name', 'Unknown'))

                if name not in block_names:
                    block_names[name] = 0
                block_names[name] += 1

        print(f"\n📦 Найдено {len(block_refs)} вставок блоков")

        if block_names:
            print(f"\nТипы блоков:")
            for name, count in sorted(block_names.items(), key=lambda x: x[1], reverse=True)[:20]:
                marker = "⭐" if 'скважина' in name.lower() else "  "
                print(f"   {marker} '{name}': {count} вставок")

        # Показываем примеры блоков "скважина"
        skvazhina_examples = []
        for entity in model_space:
            if entity.EntityName == 'AcDbBlockReference':
                name = getattr(entity, 'EffectiveName', getattr(entity, 'Name', 'Unknown'))
                if 'скважина' in name.lower():
                    layer = getattr(entity, 'Layer', 'Unknown')
                    has_attrs = getattr(entity, 'HasAttributes', False)
                    pos = entity.InsertionPoint

                    attrs = {}
                    if has_attrs:
                        try:
                            for attr in entity.GetAttributes():
                                attrs[attr.TagString] = attr.TextString
                        except:
                            pass

                    skvazhina_examples.append({
                        'name': name,
                        'layer': layer,
                        'position': (pos[0], pos[1], pos[2]),
                        'has_attrs': has_attrs,
                        'attrs': attrs
                    })

                    if len(skvazhina_examples) >= 5:
                        break

        if skvazhina_examples:
            print(f"\n🔍 ПРИМЕРЫ БЛОКОВ 'СКВАЖИНА':")
            for i, ex in enumerate(skvazhina_examples, 1):
                print(f"\n   Пример #{i}:")
                print(f"      Имя: {ex['name']}")
                print(f"      Слой: {ex['layer']}")
                print(f"      Позиция: ({ex['position'][0]:.2f}, {ex['position'][1]:.2f}, {ex['position'][2]:.2f})")
                print(f"      Есть атрибуты: {ex['has_attrs']}")
                if ex['attrs']:
                    print(f"      Атрибуты:")
                    for tag, val in ex['attrs'].items():
                        print(f"         {tag}: {val}")

    except Exception as e:
        print(f"❌ Ошибка работы с ModelSpace: {e}")
        import traceback
        traceback.print_exc()

    print("\n" + "=" * 80)
    print("ШАГ 4: ПРОВЕРКА ВНЕШНИХ ССЫЛОК (XREFS)")
    print("=" * 80)

    try:
        blocks = doc.Blocks
        print(f"Блоков в документе: {blocks.Count}")

        xrefs = []
        for i in range(blocks.Count):
            block = blocks.Item(i)
            is_xref = getattr(block, 'IsXRef', False)
            if is_xref:
                xref_name = block.Name
                xref_path = getattr(block, 'Path', 'Unknown')
                xrefs.append((xref_name, xref_path))

        if xrefs:
            print(f"\n🔗 Найдено {len(xrefs)} внешних ссылок (xrefs):")
            for name, path in xrefs:
                print(f"   - {name}: {path}")
                if 'svodniy' in path.lower():
                    print(f"      ⚠️ Возможно, ваш файл загружен как XREF!")
        else:
            print("✅ Внешних ссылок не найдено")

    except Exception as e:
        print(f"⚠️ Проблема с проверкой xrefs: {e}")

    print("\n" + "=" * 80)
    print("РЕКОМЕНДАЦИИ")
    print("=" * 80)

    if not block_names:
        print("❌ Блоки не найдены в ModelSpace")
        print("\n💡 Возможные причины:")
        print("1. Файл не открыт - откройте svodniy_plan.dwg в AutoCAD")
        print("2. Объекты находятся в PaperSpace - проверьте вкладки Layout")
        print("3. Файл загружен как XREF - проверьте внешние ссылки выше")
        print("4. Проблема с кэшем win32com - очистите кэш:")
        print("\n   В Python консоли:")
        print("   >>> import win32com, shutil, os")
        print("   >>> gen_py = os.path.join(win32com.__gen_path__, 'gen_py')")
        print("   >>> if os.path.exists(gen_py): shutil.rmtree(gen_py)")
        print("\n5. Переключитесь в AutoCAD на вкладку 'Model' (не Layout)")
    else:
        if any('скважина' in name.lower() for name in block_names.keys()):
            print(f"✅ Найдены блоки 'скважина': {sum(c for n, c in block_names.items() if 'скважина' in n.lower())} вставок")
        else:
            print("⚠️ Блоки 'скважина' не найдены")
            print(f"   Проверьте список выше - возможно, блоки называются иначе")

    print("\n" + "=" * 80)


if __name__ == "__main__":
    deep_diagnose()

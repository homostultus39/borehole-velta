"""
Комплексная диагностика системы для решения проблем с AutoCAD.
"""

import sys
import os
import subprocess
import winreg
from typing import Dict, List, Any

def check_python_info() -> Dict[str, Any]:
    """Проверка информации о Python."""
    print("🐍 ИНФОРМАЦИЯ О PYTHON")
    print("=" * 40)
    
    info = {
        'version': sys.version,
        'executable': sys.executable,
        'platform': sys.platform,
        'architecture': '64-bit' if sys.maxsize > 2**32 else '32-bit'
    }
    
    print(f"Версия: {info['version']}")
    print(f"Исполняемый файл: {info['executable']}")
    print(f"Платформа: {info['platform']}")
    print(f"Архитектура: {info['architecture']}")
    
    return info

def check_autocad_installation() -> Dict[str, Any]:
    """Проверка установки AutoCAD."""
    print("\n🏗️ ИНФОРМАЦИЯ О AUTOCAD")
    print("=" * 40)
    
    info = {
        'installed': False,
        'versions': [],
        'com_registered': False,
        'paths': []
    }
    
    # Проверяем реестр на наличие AutoCAD
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Autodesk\AutoCAD")
        print("✅ AutoCAD найден в реестре")
        info['installed'] = True
        
        # Получаем версии
        i = 0
        while True:
            try:
                version = winreg.EnumKey(key, i)
                info['versions'].append(version)
                print(f"  Версия: {version}")
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
        
    except FileNotFoundError:
        print("❌ AutoCAD не найден в реестре")
    
    # Проверяем COM-регистрацию
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "AutoCAD.Application")
        print("✅ AutoCAD.Application зарегистрирован в COM")
        info['com_registered'] = True
        winreg.CloseKey(key)
    except FileNotFoundError:
        print("❌ AutoCAD.Application не зарегистрирован в COM")
    
    # Проверяем пути установки
    common_paths = [
        r"C:\Program Files\Autodesk\AutoCAD 2024",
        r"C:\Program Files\Autodesk\AutoCAD 2025", 
        r"C:\Program Files\Autodesk\AutoCAD 2026",
        r"C:\Program Files (x86)\Autodesk\AutoCAD 2024",
        r"C:\Program Files (x86)\Autodesk\AutoCAD 2025",
        r"C:\Program Files (x86)\Autodesk\AutoCAD 2026"
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            info['paths'].append(path)
            print(f"✅ Найден путь: {path}")
    
    return info

def check_python_packages() -> Dict[str, Any]:
    """Проверка установленных Python пакетов."""
    print("\n📦 PYTHON ПАКЕТЫ")
    print("=" * 40)
    
    packages = {
        'pyautocad': False,
        'pywin32': False,
        'comtypes': False,
        'versions': {}
    }
    
    # Проверяем pyautocad
    try:
        import pyautocad
        packages['pyautocad'] = True
        packages['versions']['pyautocad'] = getattr(pyautocad, '__version__', 'unknown')
        print(f"✅ pyautocad: {packages['versions']['pyautocad']}")
    except ImportError:
        print("❌ pyautocad не установлен")
    
    # Проверяем pywin32
    try:
        import win32com.client
        packages['pywin32'] = True
        packages['versions']['pywin32'] = getattr(win32com, '__version__', 'unknown')
        print(f"✅ pywin32: {packages['versions']['pywin32']}")
    except ImportError:
        print("❌ pywin32 не установлен")
    
    # Проверяем comtypes
    try:
        import comtypes
        packages['comtypes'] = True
        packages['versions']['comtypes'] = getattr(comtypes, '__version__', 'unknown')
        print(f"✅ comtypes: {packages['versions']['comtypes']}")
    except ImportError:
        print("❌ comtypes не установлен")
    
    return packages

def test_com_connection() -> Dict[str, Any]:
    """Тест COM-подключения к AutoCAD."""
    print("\n🔌 ТЕСТ COM-ПОДКЛЮЧЕНИЯ")
    print("=" * 40)
    
    results = {
        'win32com': False,
        'comtypes': False,
        'pyautocad': False,
        'errors': {}
    }
    
    # Тест 1: win32com
    try:
        import win32com.client
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("✅ win32com: Подключение к существующему AutoCAD успешно")
        results['win32com'] = True
    except Exception as e:
        print(f"❌ win32com: {e}")
        results['errors']['win32com'] = str(e)
    
    # Тест 2: comtypes
    try:
        import comtypes.client
        acad = comtypes.client.GetActiveObject("AutoCAD.Application")
        print("✅ comtypes: Подключение к существующему AutoCAD успешно")
        results['comtypes'] = True
    except Exception as e:
        print(f"❌ comtypes: {e}")
        results['errors']['comtypes'] = str(e)
    
    # Тест 3: pyautocad
    try:
        from pyautocad import Autocad
        acad = Autocad(create_if_not_exists=False)
        print("✅ pyautocad: Подключение к существующему AutoCAD успешно")
        results['pyautocad'] = True
    except Exception as e:
        print(f"❌ pyautocad: {e}")
        results['errors']['pyautocad'] = str(e)
    
    return results

def test_autocad_versions() -> List[str]:
    """Тест различных версий AutoCAD."""
    print("\n🔍 ТЕСТ ВЕРСИЙ AUTOCAD")
    print("=" * 40)
    
    working_versions = []
    versions_to_test = [
        "AutoCAD.Application",
        "AutoCAD.Application.24",  # 2024
        "AutoCAD.Application.25",  # 2025
        "AutoCAD.Application.26",  # 2026
    ]
    
    for version in versions_to_test:
        try:
            import win32com.client
            acad = win32com.client.GetActiveObject(version)
            print(f"✅ {version}: Работает")
            working_versions.append(version)
        except Exception as e:
            print(f"❌ {version}: {e}")
    
    return working_versions

def generate_recommendations(python_info: Dict, autocad_info: Dict, 
                           packages: Dict, com_results: Dict, 
                           working_versions: List[str]) -> None:
    """Генерация рекомендаций по исправлению проблем."""
    print("\n💡 РЕКОМЕНДАЦИИ")
    print("=" * 40)
    
    # Проверяем архитектуру
    if python_info['architecture'] == '64-bit':
        print("✅ Python 64-bit - хорошо")
    else:
        print("⚠️  Python 32-bit - убедитесь, что AutoCAD тоже 32-bit")
    
    # Проверяем установку AutoCAD
    if not autocad_info['installed']:
        print("❌ КРИТИЧНО: AutoCAD не установлен или не найден в реестре")
        print("   Решение: Переустановите AutoCAD с правами администратора")
        return
    
    if not autocad_info['com_registered']:
        print("❌ КРИТИЧНО: AutoCAD не зарегистрирован в COM")
        print("   Решение: Запустите AutoCAD от имени администратора")
        print("   Или выполните: regsvr32 \"путь_к_autocad\\acad.exe\"")
    
    # Проверяем пакеты
    if not packages['pywin32']:
        print("❌ КРИТИЧНО: pywin32 не установлен")
        print("   Решение: pip install pywin32")
    
    if not packages['pyautocad']:
        print("❌ pyautocad не установлен")
        print("   Решение: pip install pyautocad")
    
    # Проверяем COM-подключение
    if not com_results['win32com'] and not com_results['comtypes']:
        print("❌ КРИТИЧНО: Ни один COM-метод не работает")
        print("   Решение: Перезапустите AutoCAD и попробуйте снова")
    
    if working_versions:
        print(f"✅ Рабочие версии AutoCAD: {', '.join(working_versions)}")
    else:
        print("❌ КРИТИЧНО: Ни одна версия AutoCAD не отвечает")
        print("   Решение: Убедитесь, что AutoCAD запущен")

def main():
    """Основная функция диагностики."""
    print("🔧 КОМПЛЕКСНАЯ ДИАГНОСТИКА СИСТЕМЫ")
    print("=" * 60)
    
    # Собираем информацию
    python_info = check_python_info()
    autocad_info = check_autocad_installation()
    packages = check_python_packages()
    com_results = test_com_connection()
    working_versions = test_autocad_versions()
    
    # Генерируем рекомендации
    generate_recommendations(python_info, autocad_info, packages, com_results, working_versions)
    
    print("\n" + "=" * 60)
    print("Диагностика завершена!")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Диагностика прервана пользователем")
    except Exception as e:
        print(f"\n❌ Ошибка во время диагностики: {e}")
    
    input("\nНажмите Enter для выхода...")

"""
Альтернативные способы подключения к AutoCAD.
Предоставляет несколько fallback-механизмов для надежного подключения.
"""

import os
import time
import logging
from typing import Optional, Dict, Any, List
from abc import ABC, abstractmethod

logger = logging.getLogger(__name__)


class AutoCADConnector(ABC):
    """Абстрактный базовый класс для подключения к AutoCAD."""
    
    @abstractmethod
    def connect(self) -> bool:
        """Подключение к AutoCAD."""
        pass
    
    @abstractmethod
    def get_application(self):
        """Получение объекта приложения AutoCAD."""
        pass
    
    @abstractmethod
    def get_active_document(self):
        """Получение активного документа."""
        pass
    
    @abstractmethod
    def open_document(self, file_path: str) -> bool:
        """Открытие документа."""
        pass


class PyAutoCADConnector(AutoCADConnector):
    """Подключение через pyautocad."""
    
    def __init__(self):
        self.acad = None
        self.doc = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """Подключение через pyautocad."""
        try:
            from pyautocad import Autocad
            
            # Сначала пытаемся подключиться к существующему AutoCAD
            try:
                self.acad = Autocad(create_if_not_exists=False)
                logger.info("✅ Подключение к существующему AutoCAD через pyautocad")
            except:
                # Если не удалось, создаем новый экземпляр
                self.acad = Autocad(create_if_not_exists=True)
                logger.info("✅ Создание нового экземпляра AutoCAD через pyautocad")
            
            self.doc = self.acad.ActiveDocument
            self.is_connected = True
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка подключения через pyautocad: {e}")
            return False
    
    def get_application(self):
        return self.acad
    
    def get_active_document(self):
        return self.doc
    
    def open_document(self, file_path: str) -> bool:
        try:
            if self.acad:
                self.acad.ActiveDocument = self.acad.Documents.Open(file_path)
                self.doc = self.acad.ActiveDocument
                return True
        except Exception as e:
            logger.error(f"Ошибка открытия документа через pyautocad: {e}")
        return False


class Win32COMConnector(AutoCADConnector):
    """Подключение через win32com."""
    
    def __init__(self):
        self.acad = None
        self.doc = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """Подключение через win32com."""
        try:
            import win32com.client
            
            # Список версий AutoCAD для попытки подключения
            autocad_versions = [
                "AutoCAD.Application.25",  # 2025 (работает по диагностике)
                "AutoCAD.Application.24",  # 2024
                "AutoCAD.Application.26",  # 2026
                "AutoCAD.Application"      # Общая версия
            ]
            
            for version in autocad_versions:
                try:
                    # Пытаемся подключиться к существующему AutoCAD
                    self.acad = win32com.client.GetActiveObject(version)
                    logger.info(f"✅ Подключение к существующему AutoCAD {version} через win32com")
                    break
                except:
                    try:
                        # Создаем новый экземпляр
                        self.acad = win32com.client.Dispatch(version)
                        logger.info(f"✅ Создание нового экземпляра AutoCAD {version} через win32com")
                        break
                    except:
                        continue
            else:
                raise Exception("Не удалось подключиться ни к одной версии AutoCAD")
            
            # Проверяем, есть ли активный документ
            try:
                self.doc = self.acad.ActiveDocument
                if self.doc is None:
                    # Создаем новый документ, если нет активного
                    logger.info("📄 Создание нового документа в AutoCAD...")
                    self.acad.Documents.Add()
                    self.doc = self.acad.ActiveDocument
                    logger.info("✅ Новый документ создан")
            except Exception as doc_error:
                logger.warning(f"⚠️ Проблема с документом: {doc_error}")
                # Пытаемся создать новый документ
                try:
                    self.acad.Documents.Add()
                    self.doc = self.acad.ActiveDocument
                    logger.info("✅ Новый документ создан после ошибки")
                except Exception as create_error:
                    logger.error(f"❌ Не удалось создать документ: {create_error}")
                    return False
            
            self.is_connected = True
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка подключения через win32com: {e}")
            return False
    
    def get_application(self):
        return self.acad
    
    def get_active_document(self):
        return self.doc
    
    def open_document(self, file_path: str) -> bool:
        try:
            if self.acad:
                self.acad.ActiveDocument = self.acad.Documents.Open(file_path)
                self.doc = self.acad.ActiveDocument
                return True
        except Exception as e:
            logger.error(f"Ошибка открытия документа через win32com: {e}")
        return False


class ComTypesConnector(AutoCADConnector):
    """Подключение через comtypes."""
    
    def __init__(self):
        self.acad = None
        self.doc = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """Подключение через comtypes."""
        try:
            import comtypes.client
            
            # Список версий AutoCAD для попытки подключения
            autocad_versions = [
                "AutoCAD.Application.25",  # 2025 (работает по диагностике)
                "AutoCAD.Application.24",  # 2024
                "AutoCAD.Application.26",  # 2026
                "AutoCAD.Application"      # Общая версия
            ]
            
            for version in autocad_versions:
                try:
                    # Пытаемся подключиться к существующему AutoCAD
                    self.acad = comtypes.client.GetActiveObject(version)
                    logger.info(f"✅ Подключение к существующему AutoCAD {version} через comtypes")
                    break
                except:
                    try:
                        # Создаем новый экземпляр
                        self.acad = comtypes.client.CreateObject(version)
                        logger.info(f"✅ Создание нового экземпляра AutoCAD {version} через comtypes")
                        break
                    except:
                        continue
            else:
                raise Exception("Не удалось подключиться ни к одной версии AutoCAD")
            
            # Проверяем, есть ли активный документ
            try:
                self.doc = self.acad.ActiveDocument
                if self.doc is None:
                    # Создаем новый документ, если нет активного
                    logger.info("📄 Создание нового документа в AutoCAD...")
                    self.acad.Documents.Add()
                    self.doc = self.acad.ActiveDocument
                    logger.info("✅ Новый документ создан")
            except Exception as doc_error:
                logger.warning(f"⚠️ Проблема с документом: {doc_error}")
                # Пытаемся создать новый документ
                try:
                    self.acad.Documents.Add()
                    self.doc = self.acad.ActiveDocument
                    logger.info("✅ Новый документ создан после ошибки")
                except Exception as create_error:
                    logger.error(f"❌ Не удалось создать документ: {create_error}")
                    return False
            
            self.is_connected = True
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка подключения через comtypes: {e}")
            return False
    
    def get_application(self):
        return self.acad
    
    def get_active_document(self):
        return self.doc
    
    def open_document(self, file_path: str) -> bool:
        try:
            if self.acad:
                self.acad.ActiveDocument = self.acad.Documents.Open(file_path)
                self.doc = self.acad.ActiveDocument
                return True
        except Exception as e:
            logger.error(f"Ошибка открытия документа через comtypes: {e}")
        return False


class DirectAutoCADConnector(AutoCADConnector):
    """Прямое подключение к AutoCAD.Application.25 (рабочая версия)."""
    
    def __init__(self):
        self.acad = None
        self.doc = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """Прямое подключение к AutoCAD.Application.25."""
        try:
            import win32com.client
            
            # Используем только рабочую версию из диагностики
            self.acad = win32com.client.GetActiveObject("AutoCAD.Application.25")
            
            # Проверяем, есть ли активный документ
            try:
                self.doc = self.acad.ActiveDocument
                if self.doc is None:
                    # НЕ создаем новый документ автоматически
                    logger.warning("⚠️ Нет активного документа. Документ должен быть открыт вручную.")
                    return False
                else:
                    doc_name = getattr(self.doc, 'Name', 'Unknown')
                    logger.info(f"📄 Активный документ: {doc_name}")
            except Exception as doc_error:
                logger.warning(f"⚠️ Проблема с документом: {doc_error}")
                # НЕ создаем новый документ автоматически
                logger.warning("⚠️ Не удалось получить активный документ. Документ должен быть открыт вручную.")
                return False
            
            self.is_connected = True
            logger.info("✅ Прямое подключение к AutoCAD.Application.25 успешно")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка прямого подключения к AutoCAD.Application.25: {e}")
            return False
    
    def get_application(self):
        return self.acad
    
    def get_active_document(self):
        return self.doc
    
    def open_document(self, file_path: str) -> bool:
        """
        Поиск нужного документа среди открытых в AutoCAD.
        """
        try:
            if self.acad:
                expected_name = os.path.basename(file_path)

                # Пытаемся получить список всех открытых документов
                try:
                    doc_count = self.acad.Documents.Count
                    logger.info(f"📂 Найдено {doc_count} открытых документов")

                    for i in range(doc_count):
                        doc = self.acad.Documents.Item(i)
                        doc_name = getattr(doc, 'Name', 'Unknown')
                        logger.info(f"   {i+1}. {doc_name}")

                        if doc_name.lower() == expected_name.lower():
                            logger.info(f"✅ Найден нужный документ: {doc_name}")
                            self.doc = doc
                            self.acad.ActiveDocument = doc  # Делаем его активным
                            return True

                    logger.warning(f"⚠️ Документ '{expected_name}' не найден среди открытых")
                    logger.warning("Откройте нужный файл в AutoCAD и запустите скрипт снова")
                    return False

                except Exception as docs_error:
                    # Если Documents не работает, пытаемся через ActiveDocument
                    logger.warning(f"⚠️ Documents не работает: {docs_error}")
                    logger.info("Проверяем активный документ...")

                    current_doc = self.acad.ActiveDocument
                    current_name = getattr(current_doc, 'Name', 'Unknown')
                    logger.info(f"📄 Активный документ: {current_name}")

                    if current_name.lower() == expected_name.lower():
                        logger.info("✅ Активный документ соответствует ожидаемому")
                        self.doc = current_doc
                        return True
                    else:
                        logger.warning(f"⚠️ Ожидается: {expected_name}, активен: {current_name}")
                        logger.warning("Переключитесь на нужный документ в AutoCAD")
                        return False

        except Exception as e:
            logger.error(f"Ошибка открытия документа: {e}")
        return False


class AutoCADConnectionManager:
    """Менеджер подключений к AutoCAD с fallback-механизмами."""
    
    def __init__(self):
        # Начинаем с прямого подключения к рабочей версии
        self.connectors = [
            DirectAutoCADConnector(),  # Прямое подключение к .25
            PyAutoCADConnector(),
            Win32COMConnector(),
            ComTypesConnector()
        ]
        self.active_connector = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """Попытка подключения через все доступные методы."""
        logger.info("🔌 Попытка подключения к AutoCAD...")
        
        for i, connector in enumerate(self.connectors):
            logger.info(f"Попытка {i+1}: {connector.__class__.__name__}")
            
            if connector.connect():
                self.active_connector = connector
                self.is_connected = True
                logger.info(f"✅ Успешное подключение через {connector.__class__.__name__}")
                return True
            
            # Небольшая задержка между попытками
            time.sleep(1)
        
        logger.error("❌ Не удалось подключиться ни одним из методов")
        return False
    
    def get_application(self):
        """Получение объекта приложения AutoCAD."""
        if self.active_connector:
            return self.active_connector.get_application()
        return None
    
    def get_active_document(self):
        """Получение активного документа."""
        if self.active_connector:
            return self.active_connector.get_active_document()
        return None
    
    def open_document(self, file_path: str) -> bool:
        """Открытие документа."""
        if self.active_connector:
            return self.active_connector.open_document(file_path)
        return False
    
    def get_connection_info(self) -> Dict[str, Any]:
        """Получение информации о текущем подключении."""
        if self.active_connector:
            return {
                'method': self.active_connector.__class__.__name__,
                'connected': self.is_connected,
                'has_application': self.get_application() is not None,
                'has_document': self.get_active_document() is not None
            }
        return {
            'method': 'None',
            'connected': False,
            'has_application': False,
            'has_document': False
        }
    
    def disconnect(self) -> bool:
        """Отключение от AutoCAD."""
        try:
            if self.active_connector:
                self.active_connector = None
            self.is_connected = False
            logger.info("✅ Отключение от AutoCAD")
            return True
        except Exception as e:
            logger.error(f"Ошибка отключения: {e}")
            return False

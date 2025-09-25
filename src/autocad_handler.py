"""
Модуль для работы с AutoCAD файлами.
Обеспечивает подключение к AutoCAD и извлечение данных из .dwg файлов.
"""

import os
from typing import List, Dict, Any, Optional
from pyautocad import Autocad, APoint
import logging

logger = logging.getLogger(__name__)


class AutoCADHandler:
    """Класс для работы с AutoCAD файлами."""
    
    def __init__(self):
        """Инициализация обработчика AutoCAD."""
        self.acad = None
        self.doc = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """
        Подключение к AutoCAD.
        
        Returns:
            bool: True если подключение успешно, False в противном случае
        """
        try:
            self.acad = Autocad(create_if_not_exists=True)
            self.doc = self.acad.ActiveDocument
            self.is_connected = True
            logger.info("Успешное подключение к AutoCAD")
            return True
        except Exception as e:
            logger.error(f"Ошибка подключения к AutoCAD: {e}")
            self.is_connected = False
            return False
    
    def open_dwg(self, file_path: str) -> bool:
        """
        Открытие .dwg файла.
        
        Args:
            file_path (str): Путь к .dwg файлу
            
        Returns:
            bool: True если файл успешно открыт, False в противном случае
        """
        if not self.is_connected:
            logger.error("Нет подключения к AutoCAD")
            return False
        
        if not os.path.exists(file_path):
            logger.error(f"Файл не найден: {file_path}")
            return False
        
        try:
            self.acad.ActiveDocument = self.acad.Documents.Open(file_path)
            self.doc = self.acad.ActiveDocument
            logger.info(f"Файл успешно открыт: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Ошибка открытия файла {file_path}: {e}")
            return False
    
    def get_all_entities(self) -> List[Any]:
        """
        Получение всех объектов из текущего документа.
        
        Returns:
            List[Any]: Список всех объектов в документе
        """
        if not self.is_connected or not self.doc:
            logger.error("Нет активного документа")
            return []
        
        try:
            entities = []
            for entity in self.doc.ModelSpace:
                entities.append(entity)
            logger.info(f"Найдено {len(entities)} объектов в документе")
            return entities
        except Exception as e:
            logger.error(f"Ошибка получения объектов: {e}")
            return []
    
    def find_text_entities(self) -> List[Dict[str, Any]]:
        """
        Поиск текстовых объектов в документе.
        
        Returns:
            List[Dict[str, Any]]: Список словарей с информацией о текстовых объектах
        """
        entities = self.get_all_entities()
        text_entities = []
        
        for entity in entities:
            try:
                # Проверяем, является ли объект текстом
                if hasattr(entity, 'TextString'):
                    text_data = {
                        'text': entity.TextString,
                        'position': (entity.InsertionPoint[0], entity.InsertionPoint[1]),
                        'layer': getattr(entity, 'Layer', 'Unknown'),
                        'entity_type': entity.EntityName
                    }
                    text_entities.append(text_data)
            except Exception as e:
                logger.warning(f"Ошибка обработки объекта: {e}")
                continue
        
        logger.info(f"Найдено {len(text_entities)} текстовых объектов")
        return text_entities
    
    def find_circles(self) -> List[Dict[str, Any]]:
        """
        Поиск кругов в документе.
        
        Returns:
            List[Dict[str, Any]]: Список словарей с информацией о кругах
        """
        entities = self.get_all_entities()
        circles = []
        
        for entity in entities:
            try:
                if entity.EntityName == 'AcDbCircle':
                    circle_data = {
                        'center': (entity.Center[0], entity.Center[1]),
                        'radius': entity.Radius,
                        'layer': getattr(entity, 'Layer', 'Unknown')
                    }
                    circles.append(circle_data)
            except Exception as e:
                logger.warning(f"Ошибка обработки круга: {e}")
                continue
        
        logger.info(f"Найдено {len(circles)} кругов")
        return circles
    
    def close_document(self) -> bool:
        """
        Закрытие текущего документа.
        
        Returns:
            bool: True если документ успешно закрыт
        """
        try:
            if self.doc:
                self.doc.Close()
                self.doc = None
            logger.info("Документ закрыт")
            return True
        except Exception as e:
            logger.error(f"Ошибка закрытия документа: {e}")
            return False
    
    def disconnect(self) -> bool:
        """
        Отключение от AutoCAD.
        
        Returns:
            bool: True если отключение успешно
        """
        try:
            if self.acad:
                self.acad = None
            self.is_connected = False
            logger.info("Отключение от AutoCAD")
            return True
        except Exception as e:
            logger.error(f"Ошибка отключения от AutoCAD: {e}")
            return False


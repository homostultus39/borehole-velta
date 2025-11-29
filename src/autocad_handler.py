"""
Модуль для работы с AutoCAD файлами.
Обеспечивает подключение к AutoCAD и извлечение данных из .dwg файлов.
Использует улучшенный менеджер подключений с fallback-механизмами.
"""

import os
import time
from typing import List, Dict, Any, Optional
import logging

from autocad_connector import AutoCADConnectionManager

logger = logging.getLogger(__name__)


class AutoCADHandler:
    """Класс для работы с AutoCAD файлами с улучшенным подключением."""
    
    def __init__(self):
        """Инициализация обработчика AutoCAD."""
        self.connection_manager = AutoCADConnectionManager()
        self.acad = None
        self.doc = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """
        Подключение к AutoCAD с использованием fallback-механизмов.
        
        Returns:
            bool: True если подключение успешно, False в противном случае
        """
        logger.info("🔌 Попытка подключения к AutoCAD...")
        
        if self.connection_manager.connect():
            self.acad = self.connection_manager.get_application()
            self.doc = self.connection_manager.get_active_document()
            self.is_connected = True
            
            # Получаем информацию о подключении
            conn_info = self.connection_manager.get_connection_info()
            logger.info(f"✅ Подключение успешно через: {conn_info['method']}")
            return True
        else:
            logger.error("❌ Не удалось подключиться к AutoCAD ни одним из методов")
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
            if self.connection_manager.open_document(file_path):
                self.doc = self.connection_manager.get_active_document()
                logger.info(f"✅ Файл успешно открыт: {file_path}")
                return True
            else:
                logger.error(f"❌ Не удалось открыть файл: {file_path}")
                return False
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
    
    def find_borehole_blocks(self, block_name: str = "скважина") -> List[Dict[str, Any]]:
        """
        Поиск блоков с именем "скважина" в документе.

        Args:
            block_name: Имя блока для поиска (по умолчанию "скважина")

        Returns:
            List[Dict[str, Any]]: Список словарей с информацией о блоках скважин
        """
        if not self.is_connected or not self.doc:
            logger.error("Нет активного документа")
            return []

        boreholes = []
        processed_count = 0

        try:
            logger.info(f"🔍 Начинаем поиск блоков с именем '{block_name}'...")
            start_time = time.time()
            max_search_time = 30

            model_space = None
            try:
                model_space = self.doc.ModelSpace
                logger.info(f"📋 ModelSpace получен, тип: {type(model_space)}")
            except:
                try:
                    model_space = self.acad.ActiveDocument.ModelSpace
                    logger.info(f"📋 ModelSpace получен (способ 2), тип: {type(model_space)}")
                except:
                    try:
                        model_space = self.acad.Documents.Item(0).ModelSpace
                        logger.info(f"📋 ModelSpace получен (способ 3), тип: {type(model_space)}")
                    except Exception as e:
                        logger.error(f"❌ Не удалось получить ModelSpace: {e}")
                        return []

            if model_space is None:
                logger.error("❌ ModelSpace не получен")
                return []

            for entity in model_space:
                processed_count += 1

                if time.time() - start_time > max_search_time:
                    logger.warning(f"⏰ Поиск прерван по времени ({max_search_time} сек). Обработано {processed_count} объектов")
                    break

                if processed_count % 100 == 0:
                    elapsed = time.time() - start_time
                    logger.info(f"📊 Обработано {processed_count} объектов за {elapsed:.1f} сек...")

                try:
                    # Проверяем, является ли объект блоком
                    if entity.EntityName == 'AcDbBlockReference':
                        entity_name = getattr(entity, 'Name', '').lower()

                        # Проверяем имя блока
                        if block_name.lower() in entity_name:
                            entity_layer = getattr(entity, 'Layer', 'Unknown')
                            insertion_point = entity.InsertionPoint

                            # Получаем атрибуты блока (номер скважины может быть там)
                            attributes = {}
                            try:
                                if hasattr(entity, 'GetAttributes'):
                                    attrs = entity.GetAttributes()
                                    for attr in attrs:
                                        tag = getattr(attr, 'TagString', '')
                                        value = getattr(attr, 'TextString', '')
                                        attributes[tag] = value
                            except:
                                pass

                            borehole_data = {
                                'name': entity_name,
                                'position': (insertion_point[0], insertion_point[1], insertion_point[2]),
                                'layer': entity_layer,
                                'entity_type': entity.EntityName,
                                'attributes': attributes
                            }
                            boreholes.append(borehole_data)

                            if len(boreholes) <= 5:
                                logger.info(f"🕳️ Найден блок '{entity_name}' на слое '{entity_layer}' в позиции {borehole_data['position']}")

                except Exception as e:
                    continue

            logger.info(f"✅ Поиск завершен! Найдено {len(boreholes)} блоков '{block_name}' из {processed_count} обработанных")
            return boreholes

        except Exception as e:
            logger.error(f"Ошибка поиска блоков скважин: {e}")
            return []

    def find_text_entities(self, layer_filter: str = "СКВ") -> List[Dict[str, Any]]:
        """
        Поиск текстовых объектов в документе с оптимизацией по слоям.

        Args:
            layer_filter: Фильтр для поиска слоев (по умолчанию "СКВ")

        Returns:
            List[Dict[str, Any]]: Список словарей с информацией о текстовых объектах
        """
        if not self.is_connected or not self.doc:
            logger.error("Нет активного документа")
            return []

        text_entities = []
        processed_count = 0

        try:
            logger.info(f"🔍 Начинаем поиск текстовых объектов на слоях, содержащих '{layer_filter}'...")
            start_time = time.time()
            max_search_time = 30

            model_space = None
            try:
                model_space = self.doc.ModelSpace
                logger.info(f"📋 ModelSpace получен, тип: {type(model_space)}")
            except:
                try:
                    model_space = self.acad.ActiveDocument.ModelSpace
                    logger.info(f"📋 ModelSpace получен (способ 2), тип: {type(model_space)}")
                except:
                    try:
                        model_space = self.acad.Documents.Item(0).ModelSpace
                        logger.info(f"📋 ModelSpace получен (способ 3), тип: {type(model_space)}")
                    except Exception as e:
                        logger.error(f"❌ Не удалось получить ModelSpace: {e}")
                        return []

            if model_space is None:
                logger.error("❌ ModelSpace не получен")
                return []

            for entity in model_space:
                processed_count += 1

                if time.time() - start_time > max_search_time:
                    logger.warning(f"⏰ Поиск прерван по времени ({max_search_time} сек). Обработано {processed_count} объектов")
                    break

                if processed_count % 100 == 0:
                    elapsed = time.time() - start_time
                    logger.info(f"📊 Обработано {processed_count} объектов за {elapsed:.1f} сек...")

                try:
                    if hasattr(entity, 'TextString'):
                        entity_layer = getattr(entity, 'Layer', 'Unknown')

                        if layer_filter and layer_filter.upper() not in entity_layer.upper():
                            continue

                        text_data = {
                            'text': entity.TextString,
                            'position': (entity.InsertionPoint[0], entity.InsertionPoint[1], entity.InsertionPoint[2]),
                            'layer': entity_layer,
                            'entity_type': entity.EntityName
                        }
                        text_entities.append(text_data)

                        if len(text_entities) <= 5:
                            logger.info(f"📝 Найден текст на слое '{entity_layer}': '{text_data['text']}'")

                except Exception as e:
                    continue

            logger.info(f"✅ Поиск завершен! Найдено {len(text_entities)} текстовых объектов на слоях с '{layer_filter}' из {processed_count} обработанных")
            return text_entities

        except Exception as e:
            logger.error(f"Ошибка поиска текстовых объектов: {e}")
            return []
    
    def find_circles(self, layer_filter: str = "СКВ") -> List[Dict[str, Any]]:
        """
        Поиск кругов в документе с оптимизацией по слоям.
        
        Args:
            layer_filter: Фильтр для поиска слоев (по умолчанию "СКВ")
        
        Returns:
            List[Dict[str, Any]]: Список словарей с информацией о кругах
        """
        if not self.is_connected or not self.doc:
            logger.error("Нет активного документа")
            return []
        
        circles = []
        processed_count = 0
        
        try:
            logger.info(f"🔍 Начинаем поиск кругов на слоях, содержащих '{layer_filter}'...")
            start_time = time.time()
            max_search_time = 30  # Максимум 30 секунд на поиск
            
            # Получаем ModelSpace правильно через COM-интерфейс
            try:
                # Пытаемся получить ModelSpace разными способами
                model_space = None
                
                # Способ 1: Прямое обращение
                try:
                    model_space = self.doc.ModelSpace
                    logger.info(f"📋 ModelSpace получен (способ 1), тип: {type(model_space)}")
                except:
                    # Способ 2: Через ActiveDocument
                    try:
                        model_space = self.acad.ActiveDocument.ModelSpace
                        logger.info(f"📋 ModelSpace получен (способ 2), тип: {type(model_space)}")
                    except:
                        # Способ 3: Через Documents коллекцию
                        try:
                            model_space = self.acad.Documents.Item(0).ModelSpace
                            logger.info(f"📋 ModelSpace получен (способ 3), тип: {type(model_space)}")
                        except Exception as e:
                            logger.error(f"❌ Все способы получения ModelSpace не сработали: {e}")
                            return []
                
                if model_space is None:
                    logger.error("❌ ModelSpace не получен ни одним способом")
                    return []
                    
            except Exception as e:
                logger.error(f"❌ Ошибка получения ModelSpace: {e}")
                return []
            
            for entity in model_space:
                processed_count += 1
                
                # Проверяем время выполнения
                if time.time() - start_time > max_search_time:
                    logger.warning(f"⏰ Поиск прерван по времени ({max_search_time} сек). Обработано {processed_count} объектов")
                    break
                
                # Показываем прогресс каждые 100 объектов
                if processed_count % 100 == 0:
                    elapsed = time.time() - start_time
                    logger.info(f"📊 Обработано {processed_count} объектов за {elapsed:.1f} сек...")
                
                try:
                    if entity.EntityName == 'AcDbCircle':
                        # Получаем слой объекта
                        entity_layer = getattr(entity, 'Layer', 'Unknown')
                        
                        # Фильтруем по слою (если указан фильтр)
                        if layer_filter and layer_filter.upper() not in entity_layer.upper():
                            continue
                        
                        circle_data = {
                            'center': (entity.Center[0], entity.Center[1], entity.Center[2]),
                            'radius': entity.Radius,
                            'layer': entity_layer
                        }
                        circles.append(circle_data)
                        
                        # Показываем найденный круг
                        if len(circles) <= 5:  # Показываем первые 5
                            logger.info(f"⭕ Найден круг на слое '{entity_layer}': центр {circle_data['center']}, радиус {circle_data['radius']}")
                
                except Exception as e:
                    # Пропускаем проблемные объекты без остановки
                    continue
            
            logger.info(f"✅ Поиск завершен! Найдено {len(circles)} кругов на слоях с '{layer_filter}' из {processed_count} обработанных")
            return circles
            
        except Exception as e:
            logger.error(f"Ошибка поиска кругов: {e}")
            return []
    
    def close_document(self) -> bool:
        """
        Закрытие текущего документа.
        
        Returns:
            bool: True если документ успешно закрыт
        """
        try:
            if self.doc:
                # Проверяем, можем ли мы закрыть документ
                try:
                    doc_name = getattr(self.doc, 'Name', 'Unknown')
                    logger.info(f"📄 Закрытие документа: {doc_name}")
                    self.doc.Close()
                    logger.info("✅ Документ закрыт")
                except Exception as close_error:
                    logger.warning(f"⚠️ Не удалось закрыть документ: {close_error}")
                    # Не критично, продолжаем
                finally:
                    self.doc = None
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
            if self.connection_manager:
                self.connection_manager.disconnect()
            self.acad = None
            self.doc = None
            self.is_connected = False
            logger.info("✅ Отключение от AutoCAD")
            return True
        except Exception as e:
            logger.error(f"Ошибка отключения от AutoCAD: {e}")
            return False
    
    def get_connection_info(self) -> Dict[str, Any]:
        """
        Получение информации о текущем подключении.
        
        Returns:
            Dict[str, Any]: Информация о подключении
        """
        if self.connection_manager:
            return self.connection_manager.get_connection_info()
        return {
            'method': 'None',
            'connected': False,
            'has_application': False,
            'has_document': False
        }
    
    def get_layers_info(self) -> List[Dict[str, Any]]:
        """
        Получение информации о слоях в документе.
        
        Returns:
            List[Dict[str, Any]]: Список слоев с информацией
        """
        if not self.is_connected or not self.doc:
            logger.error("Нет активного документа")
            return []
        
        layers_info = []
        
        try:
            logger.info("🔍 Получение информации о слоях...")
            
            # Получаем коллекцию слоев
            layers = self.doc.Layers
            
            for i in range(layers.Count):
                try:
                    layer = layers.Item(i)
                    layer_info = {
                        'name': layer.Name,
                        'color': getattr(layer, 'Color', 'Unknown'),
                        'visible': getattr(layer, 'LayerOn', True),
                        'locked': getattr(layer, 'Lock', False)
                    }
                    layers_info.append(layer_info)
                except Exception as e:
                    logger.warning(f"Ошибка получения информации о слое {i}: {e}")
                    continue
            
            logger.info(f"✅ Получена информация о {len(layers_info)} слоях")
            return layers_info
            
        except Exception as e:
            logger.error(f"Ошибка получения информации о слоях: {e}")
            return []


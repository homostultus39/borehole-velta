"""
Модуль для обработки данных о скважинах.
Обеспечивает определение номеров скважин и расчет относительных высот.
"""

import re
import random
import logging
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class Borehole:
    """Класс для представления скважины."""
    number: str
    x: float
    y: float
    z: Optional[float] = None
    relative_height: Optional[float] = None
    text_entity: Optional[Dict[str, Any]] = None
    circle_entity: Optional[Dict[str, Any]] = None


class BoreholeProcessor:
    """Класс для обработки данных о скважинах."""
    
    def __init__(self):
        """Инициализация процессора скважин."""
        self.boreholes: List[Borehole] = []
        self.reference_borehole: Optional[Borehole] = None
        self.borehole_patterns = [
            r'скв[а-я]*\.?\s*(\d+)',  # скв. 123, скважина 123
            r'№\s*(\d+)',             # № 123
            r'(\d+)\s*скв',           # 123 скв
            r'скв\s*(\d+)',           # скв 123
            r'^(\d+)$',               # просто число
        ]
    
    def extract_borehole_from_blocks(self, borehole_blocks: List[Dict[str, Any]]) -> List[Borehole]:
        """
        Извлечение скважин из блоков AutoCAD.

        Args:
            borehole_blocks: Список блоков "скважина" из AutoCAD

        Returns:
            List[Borehole]: Список найденных скважин
        """
        self.boreholes = []

        for idx, block in enumerate(borehole_blocks):
            position = block['position']
            attributes = block.get('attributes', {})

            # Пытаемся найти номер скважины в атрибутах
            borehole_number = None
            for tag, value in attributes.items():
                if value and self._extract_borehole_number(value):
                    borehole_number = self._extract_borehole_number(value)
                    break

            # Если номер не найден в атрибутах, используем порядковый номер
            if not borehole_number:
                borehole_number = str(idx + 1)
                logger.warning(f"Блок скважины без номера, присвоен номер {borehole_number}")

            borehole = Borehole(
                number=borehole_number,
                x=position[0],
                y=position[1],
                z=position[2] if len(position) > 2 else 0.0
            )

            self.boreholes.append(borehole)
            logger.info(f"Найдена скважина №{borehole_number} в позиции ({borehole.x:.2f}, {borehole.y:.2f}, {borehole.z:.2f})")

        logger.info(f"Всего найдено {len(self.boreholes)} скважин")
        return self.boreholes

    def extract_borehole_numbers(self, text_entities: List[Dict[str, Any]],
                                circles: List[Dict[str, Any]]) -> List[Borehole]:
        """
        Извлечение номеров скважин из текстовых объектов и кругов.

        Args:
            text_entities: Список текстовых объектов из AutoCAD
            circles: Список кругов из AutoCAD

        Returns:
            List[Borehole]: Список найденных скважин
        """
        self.boreholes = []

        # Обрабатываем текстовые объекты
        for text_entity in text_entities:
            text = text_entity['text'].strip()
            position = text_entity['position']

            # Ищем номер скважины в тексте
            borehole_number = self._extract_borehole_number(text)
            if borehole_number:
                # Ищем ближайший круг к тексту
                closest_circle = self._find_closest_circle(position, circles)

                borehole = Borehole(
                    number=borehole_number,
                    x=position[0],
                    y=position[1],
                    z=position[2] if len(position) > 2 else 0.0,
                    text_entity=text_entity,
                    circle_entity=closest_circle
                )

                # Если найден круг, используем его координаты
                if closest_circle:
                    borehole.x = closest_circle['center'][0]
                    borehole.y = closest_circle['center'][1]
                    borehole.z = closest_circle['center'][2] if len(closest_circle['center']) > 2 else borehole.z

                self.boreholes.append(borehole)
                logger.info(f"Найдена скважина №{borehole_number} в позиции ({borehole.x:.2f}, {borehole.y:.2f})")

        # Обрабатываем круги без текста (возможно, номера скважин в атрибутах)
        for circle in circles:
            # Проверяем, не связан ли уже этот круг со скважиной
            if not any(bh.circle_entity == circle for bh in self.boreholes):
                # Пытаемся найти текст рядом с кругом
                nearby_text = self._find_nearby_text(circle['center'], text_entities)
                if nearby_text:
                    borehole_number = self._extract_borehole_number(nearby_text['text'])
                    if borehole_number:
                        borehole = Borehole(
                            number=borehole_number,
                            x=circle['center'][0],
                            y=circle['center'][1],
                            z=circle['center'][2] if len(circle['center']) > 2 else 0.0,
                            circle_entity=circle,
                            text_entity=nearby_text
                        )
                        self.boreholes.append(borehole)
                        logger.info(f"Найдена скважина №{borehole_number} по кругу в позиции ({borehole.x:.2f}, {borehole.y:.2f})")

        logger.info(f"Всего найдено {len(self.boreholes)} скважин")
        return self.boreholes
    
    def _extract_borehole_number(self, text: str) -> Optional[str]:
        """
        Извлечение номера скважины из текста.
        
        Args:
            text: Текст для анализа
            
        Returns:
            Optional[str]: Номер скважины или None
        """
        text_lower = text.lower()
        
        for pattern in self.borehole_patterns:
            match = re.search(pattern, text_lower)
            if match:
                return match.group(1)
        
        return None
    
    def _find_closest_circle(self, text_position: Tuple[float, float], 
                           circles: List[Dict[str, Any]], 
                           max_distance: float = 50.0) -> Optional[Dict[str, Any]]:
        """
        Поиск ближайшего круга к тексту.
        
        Args:
            text_position: Позиция текста (x, y)
            circles: Список кругов
            max_distance: Максимальное расстояние для поиска
            
        Returns:
            Optional[Dict[str, Any]]: Ближайший круг или None
        """
        if not circles:
            return None
        
        closest_circle = None
        min_distance = float('inf')
        
        for circle in circles:
            circle_center = circle['center']
            distance = ((text_position[0] - circle_center[0]) ** 2 + 
                       (text_position[1] - circle_center[1]) ** 2) ** 0.5
            
            if distance < min_distance and distance <= max_distance:
                min_distance = distance
                closest_circle = circle
        
        return closest_circle
    
    def _find_nearby_text(self, circle_center: Tuple[float, float], 
                         text_entities: List[Dict[str, Any]], 
                         max_distance: float = 50.0) -> Optional[Dict[str, Any]]:
        """
        Поиск текста рядом с кругом.
        
        Args:
            circle_center: Центр круга (x, y)
            text_entities: Список текстовых объектов
            max_distance: Максимальное расстояние для поиска
            
        Returns:
            Optional[Dict[str, Any]]: Ближайший текст или None
        """
        closest_text = None
        min_distance = float('inf')
        
        for text_entity in text_entities:
            text_position = text_entity['position']
            distance = ((circle_center[0] - text_position[0]) ** 2 + 
                       (circle_center[1] - text_position[1]) ** 2) ** 0.5
            
            if distance < min_distance and distance <= max_distance:
                min_distance = distance
                closest_text = text_entity
        
        return closest_text
    
    def set_reference_borehole(self, borehole_number: Optional[str] = None) -> bool:
        """
        Установка опорной скважины с относительной высотой 0.

        Args:
            borehole_number: Номер опорной скважины. Если None, выбирается случайная.

        Returns:
            bool: True если опорная скважина установлена
        """
        if not self.boreholes:
            logger.error("Нет скважин для установки опорной")
            return False

        if borehole_number:
            # Ищем скважину по номеру
            for borehole in self.boreholes:
                if borehole.number == borehole_number:
                    self.reference_borehole = borehole
                    borehole.relative_height = 0.0
                    logger.info(f"Установлена опорная скважина №{borehole_number}")
                    return True
            logger.error(f"Скважина №{borehole_number} не найдена")
            return False
        else:
            # Выбираем случайную скважину
            self.reference_borehole = random.choice(self.boreholes)
            self.reference_borehole.relative_height = 0.0
            logger.info(f"Случайно выбрана опорная скважина №{self.reference_borehole.number}")
            return True
    
    def calculate_relative_heights(self, reference_z: float = 0.0) -> bool:
        """
        Расчет относительных высот скважин относительно опорной скважины.

        Args:
            reference_z: Z-координата опорной скважины (по умолчанию 0.0)

        Returns:
            bool: True если расчет выполнен успешно
        """
        if not self.reference_borehole:
            logger.error("Не установлена опорная скважина")
            return False

        # Если опорная скважина имеет Z-координату из AutoCAD, используем её
        # Иначе устанавливаем reference_z
        if self.reference_borehole.z is None or self.reference_borehole.z == 0.0:
            self.reference_borehole.z = reference_z

        logger.info(f"Опорная скважина №{self.reference_borehole.number}: Z = {self.reference_borehole.z:.2f}")

        for borehole in self.boreholes:
            if borehole == self.reference_borehole:
                continue

            # Если у скважины нет Z-координаты, используем 0.0 как значение по умолчанию
            if borehole.z is None:
                borehole.z = 0.0
                logger.warning(f"Скважина №{borehole.number}: Z-координата отсутствует, используется 0.0")

            # Рассчитываем относительную высоту как смещение от опорной скважины
            borehole.relative_height = borehole.z - self.reference_borehole.z
            logger.debug(f"Скважина №{borehole.number}: Z = {borehole.z:.2f}, относительная высота = {borehole.relative_height:.2f}")

        logger.info("Расчет относительных высот завершен")
        return True
    
    def get_boreholes_data(self) -> List[Dict[str, Any]]:
        """
        Получение данных о скважинах в виде списка словарей.
        
        Returns:
            List[Dict[str, Any]]: Данные о скважинах
        """
        data = []
        for borehole in self.boreholes:
            data.append({
                'number': borehole.number,
                'x': borehole.x,
                'y': borehole.y,
                'z': borehole.z,
                'relative_height': borehole.relative_height,
                'is_reference': borehole == self.reference_borehole
            })
        return data
    
    def get_reference_borehole(self) -> Optional[Borehole]:
        """
        Получение опорной скважины.
        
        Returns:
            Optional[Borehole]: Опорная скважина или None
        """
        return self.reference_borehole


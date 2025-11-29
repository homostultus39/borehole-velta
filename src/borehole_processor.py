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
        Извлечение скважин из вставок блоков AutoCAD.

        Args:
            borehole_blocks: Список вставок блоков "скважина" из AutoCAD

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
                    logger.debug(f"Найден номер в атрибуте '{tag}': {borehole_number}")
                    break

            # Если номер не найден в атрибутах, используем порядковый номер
            if not borehole_number:
                borehole_number = str(idx + 1)
                logger.warning(f"Вставка блока без номера в атрибутах, присвоен номер {borehole_number}")

            borehole = Borehole(
                number=borehole_number,
                x=position[0],
                y=position[1],
                z=position[2] if len(position) > 2 else 0.0
            )

            self.boreholes.append(borehole)
            logger.info(f"Скважина №{borehole_number}: позиция ({borehole.x:.2f}, {borehole.y:.2f}, {borehole.z:.2f})")

        logger.info(f"Всего обработано {len(self.boreholes)} скважин")
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


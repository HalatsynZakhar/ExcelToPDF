"""
Утилиты для работы с изображениями
"""
import os
import re
import io
import logging
import math
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any, Union, Set
import sys
import tempfile

from PIL import Image as PILImage

logger = logging.getLogger(__name__)

# Глобальный кэш для хранения оптимального качества сжатия
cached_quality = None

def normalize_article(article: Any, for_excel: bool = False) -> str:
    """
    Нормализует артикул для поиска.
    В зависимости от режима нормализует по-разному:
    
    - for_excel=True: Входные данные из Excel - заменяет все спецсимволы, кроме пробелов, на дефисы.
    - for_excel=False: Имена файлов изображений - заменяет все спецсимволы, кроме пробелов и нижнего подчеркивания, на дефисы.
    
    Args:
        article (Any): Артикул в любом формате
        for_excel (bool): Флаг, указывающий что это данные из Excel (True) или имя файла изображения (False)
        
    Returns:
        str: Нормализованный артикул
    """
    if article is None:
        return ""
        
    # Преобразуем в строку и удаляем пробелы в начале и конце
    article_str = str(article).strip()
    
    # Если строка пустая, возвращаем пустую строку
    if not article_str:
        return ""
    
    if for_excel:
        # Для данных из Excel: заменяем все спецсимволы (кроме пробелов) на дефисы
        # Сохраняем буквы, цифры и пробелы, остальное заменяем на дефисы
        normalized = ''
        for char in article_str:
            if char.isalnum() or char == ' ':
                normalized += char
            else:
                normalized += '-'
        # Приводим к нижнему регистру
        normalized = normalized.lower()
    else:
        # Для имен файлов: заменяем все спецсимволы (кроме пробелов и нижнего подчеркивания) на дефисы
        # Сохраняем буквы, цифры, пробелы и нижнее подчеркивание
        normalized = ''
        for char in article_str:
            if char.isalnum() or char == ' ' or char == '_':
                normalized += char
            else:
                normalized += '-'
        # Приводим к нижнему регистру
        normalized = normalized.lower()
    
    return normalized

def optimize_image_for_excel(image_path: str, target_size_kb: int = 100, 
                          quality: int = 90, min_quality: int = 1,
                          output_folder: Optional[str] = None) -> io.BytesIO:
    """
    Оптимизирует изображение до заданного размера в КБ для вставки в Excel с кешированием качества.
    Первый файл оптимизируется с подбором качества, последующие - с использованием кешированного качества.
    
    Args:
        image_path (str): Путь к изображению
        target_size_kb (int): Целевой размер файла в КБ
        quality (int): Начальное качество JPEG (1-100)
        min_quality (int): Минимальное качество JPEG
        output_folder (Optional[str]): Папка для сохранения
        
    Returns:
        io.BytesIO: Буфер с оптимизированным изображением
    """
    global cached_quality
    print(f"  [optimize_excel] Оптимизация изображения: {image_path}", file=sys.stderr)
    
    if cached_quality is not None:
        print(f"  [optimize_excel] Используем кешированное качество: {cached_quality}%", file=sys.stderr)
        img = PILImage.open(image_path)
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        buffer = io.BytesIO()
        img.save(buffer, format='JPEG', quality=cached_quality)
        buffer.seek(0)
        return buffer

    print(f"  [optimize_excel] Подбор качества для первого изображения", file=sys.stderr)
    buffer = io.BytesIO()
    current_quality = quality
    best_buffer = None
    
    while current_quality >= min_quality:
        buffer.seek(0)
        buffer.truncate(0)
        
        try:
            img = PILImage.open(image_path)
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img.save(buffer, format='JPEG', quality=current_quality)
            size_kb = buffer.tell() / 1024
            print(f"    Качество {current_quality}%: {size_kb:.1f} КБ", file=sys.stderr)
            
            if size_kb <= target_size_kb:
                print(f"  [optimize_excel] Найдено качество: {current_quality}%", file=sys.stderr)
                cached_quality = current_quality
                buffer.seek(0)
                return buffer
                
            current_quality -= 5
        except Exception as e:
            print(f"    Ошибка: {e}", file=sys.stderr)
            current_quality -= 5
    
    print(f"  [optimize_excel] Используем минимальное качество", file=sys.stderr)
    cached_quality = min_quality
    buffer.seek(0)
    buffer.truncate(0)
    img = PILImage.open(image_path)
    if img.mode == 'RGBA':
        img = img.convert('RGB')
    img.save(buffer, format='JPEG', quality=min_quality)
    buffer.seek(0)
    return buffer

# Остальные функции остаются без изменений
# ...

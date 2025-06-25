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
    Оптимизирует изображение до заданного размера в КБ для вставки в Excel.
    Для первого изображения использует двухэтапную оптимизацию качества:
      1. От 100% до 5% с шагом 5%
      2. От 4% до 1% с шагом 1%
    Для последующих изображений использует кешированное качество.
    Всегда сохраняет изображение с качеством 1%, даже если размер превышает лимит.
    
    Args:
        image_path (str): Путь к изображению
        target_size_kb (int): Целевой размер файла в КБ
        quality (int): Не используется (оставлен для совместимости)
        min_quality (int): Не используется (оставлен для совместимости)
        output_folder (Optional[str]): Папка для сохранения
        
    Returns:
        io.BytesIO: Буфер с оптимизированным изображением
    """
    global cached_quality
    
    # Если качество кешировано - используем его
    if cached_quality is not None:
        print(f"  [optimize_excel] Используем кешированное качество: {cached_quality}%", file=sys.stderr)
        img = PILImage.open(image_path)
        if img.mode == 'RGBA' or 'transparency' in img.info:
            background = PILImage.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')
            
        buffer = io.BytesIO()
        img.save(buffer, format='JPEG', quality=cached_quality)
        buffer.seek(0)
        return buffer

    print(f"  [optimize_excel] Оптимизация первого изображения: {image_path}", file=sys.stderr)
    print(f"  [optimize_excel] Цель: < {target_size_kb} КБ", file=sys.stderr)

    img = PILImage.open(image_path)
    if img.mode == 'RGBA' or 'transparency' in img.info:
        print("  [optimize_excel] Обнаружена прозрачность, заменяем на белый фон.", file=sys.stderr)
        background = PILImage.new('RGB', img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[3])
        img = background
    elif img.mode != 'RGB':
        print(f"  [optimize_excel] Конвертируем изображение из {img.mode} в RGB.", file=sys.stderr)
        img = img.convert('RGB')

    best_buffer = None
    best_quality = None

    # Этап 1: от 100 до 5 с шагом 5
    for q in range(100, 4, -5):
        buffer = io.BytesIO()
        img.save(buffer, format='JPEG', quality=q)
        size_kb = buffer.tell() / 1024
        print(f"    Этап 1: качество {q}% - размер {size_kb:.1f} КБ", file=sys.stderr)
        
        if size_kb <= target_size_kb:
            best_buffer = buffer
            best_quality = q
            print(f"  [optimize_excel] Найдено подходящее качество: {q}%", file=sys.stderr)
            break

    # Этап 2: от 4 до 1 с шагом 1 (если не нашли в этапе 1)
    if best_buffer is None:
        for q in range(4, 0, -1):
            buffer = io.BytesIO()
            img.save(buffer, format='JPEG', quality=q)
            size_kb = buffer.tell() / 1024
            print(f"    Этап 2: качество {q}% - размер {size_kb:.1f} КБ", file=sys.stderr)
            
            if size_kb <= target_size_kb:
                best_buffer = buffer
                best_quality = q
                print(f"  [optimize_excel] Найдено подходящее качество: {q}%", file=sys.stderr)
                break

    # Если не нашли подходящего качества, используем 1%
    if best_buffer is None:
        print(f"  [optimize_excel] Используем минимальное качество (1%)", file=sys.stderr)
        best_buffer = io.BytesIO()
        img.save(best_buffer, format='JPEG', quality=1)
        best_quality = 1

    # Кешируем найденное качество для следующих изображений
    cached_quality = best_quality
    print(f"  [optimize_excel] Итоговое качество: {best_quality}% (кешировано)", file=sys.stderr)
    best_buffer.seek(0)
    return best_buffer

# Остальные функции остаются без изменений
# ...

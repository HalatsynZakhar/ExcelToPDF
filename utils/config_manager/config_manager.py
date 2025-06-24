"""
Модуль для управления конфигурацией приложения
"""
import os
import json
import logging
from typing import Dict, Any, List, Optional
from pathlib import Path

logger = logging.getLogger(__name__)

class ConfigManager:
    """
    Класс для управления настройками приложения
    """
    
    def __init__(self, presets_folder: str):
        """
        Инициализирует менеджер конфигурации
        
        Args:
            presets_folder: Путь к папке с настройками
        """
        self.presets_folder = presets_folder
        self.current_settings = {}
        
        # Создаем папку для пресетов, если она не существует
        os.makedirs(self.presets_folder, exist_ok=True)
        
        # Инициализируем настройки по умолчанию
        self.reset_settings()
        
        logger.info(f"ConfigManager инициализирован с папкой настроек: {presets_folder}")
    
    def reset_settings(self):
        """
        Сбрасывает настройки к значениям по умолчанию
        """
        self.current_settings = {
            "excel_settings": {
                "article_column": "C",
                "image_column": "A",
                "start_row": 2,
                "adjust_dimensions": True
            },
            "image_settings": {
                "max_size_kb": 100,
                "quality": 90,
                "target_width": 300,
                "target_height": 300,
                "supported_extensions": [".jpg", ".jpeg", ".png", ".gif", ".bmp"]
            },
            "ui_settings": {
                "show_preview": True,
                "show_stats": True,
                "theme": "light"
            }
        }
        
        logger.info("Настройки сброшены к значениям по умолчанию")
    
    def get_setting(self, path: str, default=None) -> Any:
        """
        Получает значение настройки по указанному пути
        
        Args:
            path: Путь к настройке в формате dot notation (например, "paths.input_folder")
            default: Значение по умолчанию, если настройка не найдена
            
        Returns:
            Значение настройки или default, если настройка не найдена
        """
        parts = path.split('.')
        current = self.current_settings
        
        for part in parts:
            if part not in current:
                return default
            current = current[part]
        
        return current
    
    def set_setting(self, path: str, value: Any):
        """
        Устанавливает значение настройки по указанному пути
        
        Args:
            path: Путь к настройке в формате dot notation (например, "paths.input_folder")
            value: Новое значение настройки
        """
        parts = path.split('.')
        current = self.current_settings
        
        # Проходим по всем частям пути, кроме последней
        for i in range(len(parts) - 1):
            part = parts[i]
            
            # Если такого ключа нет, создаем его как словарь
            if part not in current:
                current[part] = {}
            
            current = current[part]
        
        # Устанавливаем значение для последней части пути
        current[parts[-1]] = value
        
        logger.debug(f"Установлена настройка {path} = {value}")
    
    def save_settings(self, preset_name: str = None) -> bool:
        """
        Сохраняет текущие настройки в файл
        
        Returns:
            True, если настройки успешно сохранены, иначе False
        """
        preset_path = os.path.join(self.presets_folder, "settings.json")
        
        try:
            with open(preset_path, 'w', encoding='utf-8') as f:
                json.dump(self.current_settings, f, indent=4, ensure_ascii=False)
            
            logger.info("Настройки успешно сохранены")
            return True
        except Exception as e:
            logger.error(f"Ошибка при сохранении настроек: {e}")
            return False
    
    def load_settings(self, preset_name: str = None) -> bool:
        """
        Загружает настройки из файла
        
        Returns:
            True, если настройки успешно загружены, иначе False
        """
        preset_path = os.path.join(self.presets_folder, "settings.json")
        
        if not os.path.exists(preset_path):
            logger.warning("Файл настроек не найден, используются настройки по умолчанию")
            return False
        
        try:
            with open(preset_path, 'r', encoding='utf-8') as f:
                loaded_settings = json.load(f)
            
            # Обновляем только те настройки, которые есть в загруженном файле
            self._update_settings_recursive(self.current_settings, loaded_settings)
            
            logger.info("Настройки успешно загружены")
            return True
        except Exception as e:
            logger.error(f"Ошибка при загрузке настроек: {e}")
            return False
    
    def _update_settings_recursive(self, target: dict, source: dict):
        """
        Рекурсивно обновляет словарь настроек
        
        Args:
            target: Целевой словарь для обновления
            source: Исходный словарь с новыми значениями
        """
        for key, value in source.items():
            if key in target and isinstance(target[key], dict) and isinstance(value, dict):
                # Если ключ существует в обоих словарях и они оба словари, обновляем рекурсивно
                self._update_settings_recursive(target[key], value)
            else:
                # Иначе просто заменяем значение
                target[key] = value 
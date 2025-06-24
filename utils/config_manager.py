import json
import os
import logging
from typing import Dict, Any, Optional, List
import copy
from pathlib import Path

# Setup logging
logger = logging.getLogger(__name__)

def get_downloads_folder():
    """Возвращает путь к папке загрузок пользователя"""
    if os.name == 'nt':  # Windows
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            try:
                return winreg.QueryValueEx(key, downloads_guid)[0]
            except:
                return os.path.join(os.path.expanduser('~'), 'Downloads')
    else:  # Linux, macOS и другие
        return os.path.join(os.path.expanduser('~'), 'Downloads')

class ConfigManager:
    """
    Класс для управления конфигурацией приложения.
    Позволяет сохранять и загружать настройки из JSON файлов.
    """
    
    def __init__(self, presets_folder: str):
        """
        Инициализирует менеджер конфигурации.
        
        Args:
            presets_folder: Путь к папке с пресетами настроек
        """
        self.presets_folder = presets_folder
        self.default_settings = self._get_default_settings()
        self.current_settings = copy.deepcopy(self.default_settings)
        self.current_preset_name = "Default"
        
        # Создаем папку для пресетов, если она не существует
        os.makedirs(self.presets_folder, exist_ok=True)
    
    def _get_default_settings(self) -> Dict[str, Any]:
        """
        Возвращает настройки по умолчанию.
        
        Returns:
            Словарь с настройками по умолчанию
        """
        return {
            "paths": {
                "input_file_path": "",
                "output_folder_path": "",
                "images_folder_path": "",
                "secondary_images_folder_path": "",  # Путь к папке с запасными изображениями (второй приоритет)
                "tertiary_images_folder_path": ""    # Путь к папке с дополнительными запасными изображениями (третий приоритет)
            },
            "excel_settings": {
                "article_column": "C",  # Столбец с артикулами
                "image_column": "A",    # Столбец для вставки изображений
                "start_row": 2,         # Начальная строка (обычно заголовки в 1-й строке)
                "adjust_dimensions": True  # Настраивать размеры строк и столбцов
            },
            "image_settings": {
                "max_size_kb": 100,     # Максимальный размер изображения в КБ
                "quality": 90,          # Начальное качество JPEG
                "min_quality": 5,       # Минимальное качество JPEG (снижено для большего сжатия)
                "target_width": 300,    # Целевая ширина изображения
                "target_height": 300,   # Целевая высота изображения
                "supported_extensions": [".jpg", ".jpeg", ".png", ".gif", ".bmp"],
                "resize_enabled": True
            },
            "ui_settings": {
                "theme": "light",
                "language": "ru",
                "show_preview": True,
                "show_stats": True
            }
        }
    
    def get_setting(self, path: str, default: Any = None) -> Any:
        """
        Получает значение настройки по указанному пути.
        
        Args:
            path: Путь к настройке в формате "section.subsection.key"
            default: Значение по умолчанию, если настройка не найдена
            
        Returns:
            Значение настройки или default, если настройка не найдена
        """
        keys = path.split('.')
        current = self.current_settings
        
        for key in keys:
            if isinstance(current, dict) and key in current:
                current = current[key]
            else:
                return default
        
        return current
    
    def set_setting(self, path: str, value: Any) -> None:
        """
        Устанавливает значение настройки по указанному пути.
        
        Args:
            path: Путь к настройке в формате "section.subsection.key"
            value: Новое значение настройки
        """
        keys = path.split('.')
        current = self.current_settings
        
        # Проходим по всем ключам кроме последнего
        for key in keys[:-1]:
            if key not in current:
                current[key] = {}
            current = current[key]
        
        # Устанавливаем значение для последнего ключа
        current[keys[-1]] = value
    
    def save_settings(self, preset_name: str) -> bool:
        """
        Сохраняет текущие настройки в файл пресета.
        
        Args:
            preset_name: Имя пресета
            
        Returns:
            True, если настройки успешно сохранены, иначе False
        """
        try:
            # Проверяем, что имя пресета не пустое
            if not preset_name:
                logger.error("Имя пресета не может быть пустым")
                return False
            
            # Формируем путь к файлу пресета
            safe_name = preset_name.replace('/', '_').replace('\\', '_')
            preset_path = os.path.join(self.presets_folder, f"{safe_name}.json")
            
            # Сохраняем настройки в файл
            with open(preset_path, 'w', encoding='utf-8') as f:
                json.dump(self.current_settings, f, ensure_ascii=False, indent=2)
            
            self.current_preset_name = preset_name
            logger.info(f"Настройки сохранены в пресет: '{preset_name}' ({preset_path})")
            return True
        except Exception as e:
            logger.error(f"Ошибка при сохранении настроек: {e}")
            return False
    
    def load_settings(self, preset_name: str) -> bool:
        """
        Загружает настройки из файла пресета.
        
        Args:
            preset_name: Имя пресета
            
        Returns:
            True, если настройки успешно загружены, иначе False
        """
        try:
            # Проверяем, что имя пресета не пустое
            if not preset_name:
                logger.error("Имя пресета не может быть пустым")
                return False
            
            # Формируем путь к файлу пресета
            safe_name = preset_name.replace('/', '_').replace('\\', '_')
            preset_path = os.path.join(self.presets_folder, f"{safe_name}.json")
            
            # Проверяем, что файл существует
            if not os.path.isfile(preset_path):
                logger.error(f"Файл пресета не найден: {preset_path}")
                return False
            
            # Загружаем настройки из файла
            with open(preset_path, 'r', encoding='utf-8') as f:
                loaded_settings = json.load(f)
            
            # Объединяем загруженные настройки с настройками по умолчанию
            # (чтобы добавить новые настройки, которых могло не быть в сохраненном файле)
            merged_settings = copy.deepcopy(self.default_settings)
            self._merge_dict(merged_settings, loaded_settings)
            
            self.current_settings = merged_settings
            self.current_preset_name = preset_name
            logger.info(f"Настройки загружены из пресета: '{preset_name}' ({preset_path})")
            return True
        except Exception as e:
            logger.error(f"Ошибка при загрузке настроек: {e}")
            return False
    
    def _merge_dict(self, target: Dict, source: Dict) -> None:
        """
        Рекурсивно объединяет два словаря.
        
        Args:
            target: Целевой словарь (будет изменен)
            source: Исходный словарь (останется неизменным)
        """
        for key, value in source.items():
            if key in target and isinstance(target[key], dict) and isinstance(value, dict):
                self._merge_dict(target[key], value)
            else:
                target[key] = value
    
    def reset_settings(self) -> None:
        """
        Сбрасывает настройки до значений по умолчанию.
        """
        self.current_settings = copy.deepcopy(self.default_settings)
        self.current_preset_name = "Default"
        logger.info("Настройки сброшены до значений по умолчанию")
    
    def get_presets_list(self) -> List[str]:
        """
        Возвращает список доступных пресетов.
        
        Returns:
            Список имен пресетов
        """
        try:
            presets = []
            for filename in os.listdir(self.presets_folder):
                if filename.endswith('.json'):
                    preset_name = os.path.splitext(filename)[0]
                    presets.append(preset_name)
            return sorted(presets)
        except Exception as e:
            logger.error(f"Ошибка при получении списка пресетов: {e}")
            return []
    
    def delete_preset(self, preset_name: str) -> bool:
        """
        Удаляет пресет.
        
        Args:
            preset_name: Имя пресета
            
        Returns:
            True, если пресет успешно удален, иначе False
        """
        try:
            # Проверяем, что имя пресета не пустое
            if not preset_name:
                logger.error("Имя пресета не может быть пустым")
                return False
            
            # Формируем путь к файлу пресета
            safe_name = preset_name.replace('/', '_').replace('\\', '_')
            preset_path = os.path.join(self.presets_folder, f"{safe_name}.json")
            
            # Проверяем, что файл существует
            if not os.path.isfile(preset_path):
                logger.error(f"Файл пресета не найден: {preset_path}")
                return False
            
            # Удаляем файл
            os.remove(preset_path)
            
            # Если это был текущий пресет, сбрасываем настройки
            if self.current_preset_name == preset_name:
                self.reset_settings()
            
            logger.info(f"Пресет удален: '{preset_name}' ({preset_path})")
            return True
        except Exception as e:
            logger.error(f"Ошибка при удалении пресета: {e}")
            return False
    
    def export_settings(self, file_path: str) -> bool:
        """
        Экспортирует текущие настройки в файл.
        
        Args:
            file_path: Путь к файлу для экспорта
            
        Returns:
            True, если настройки успешно экспортированы, иначе False
        """
        try:
            # Создаем директорию, если она не существует
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            # Сохраняем настройки в файл
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.current_settings, f, ensure_ascii=False, indent=2)
            
            logger.info(f"Настройки экспортированы в файл: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при экспорте настроек: {e}")
            return False
    
    def import_settings(self, file_path: str) -> bool:
        """
        Импортирует настройки из файла.
        
        Args:
            file_path: Путь к файлу для импорта
            
        Returns:
            True, если настройки успешно импортированы, иначе False
        """
        try:
            # Проверяем, что файл существует
            if not os.path.isfile(file_path):
                logger.error(f"Файл не найден: {file_path}")
                return False
            
            # Загружаем настройки из файла
            with open(file_path, 'r', encoding='utf-8') as f:
                loaded_settings = json.load(f)
            
            # Объединяем загруженные настройки с настройками по умолчанию
            merged_settings = copy.deepcopy(self.default_settings)
            self._merge_dict(merged_settings, loaded_settings)
            
            self.current_settings = merged_settings
            self.current_preset_name = "Imported"
            logger.info(f"Настройки импортированы из файла: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при импорте настроек: {e}")
            return False


# Глобальный экземпляр ConfigManager
_config_manager = None

def init_config_manager(presets_folder: str) -> None:
    """
    Инициализирует глобальный экземпляр ConfigManager.
    
    Args:
        presets_folder: Путь к папке с пресетами настроек
    """
    global _config_manager
    _config_manager = ConfigManager(presets_folder)

def get_config_manager() -> ConfigManager:
    """
    Возвращает глобальный экземпляр ConfigManager.
    
    Returns:
        Экземпляр ConfigManager
    """
    global _config_manager
    if _config_manager is None:
        raise RuntimeError("ConfigManager не инициализирован. Вызовите init_config_manager() перед использованием.")
    return _config_manager

def get_setting(path: str, default: Any = None) -> Any:
    """
    Получает значение настройки по указанному пути.
    
    Args:
        path: Путь к настройке в формате "section.subsection.key"
        default: Значение по умолчанию, если настройка не найдена
        
    Returns:
        Значение настройки или default, если настройка не найдена
    """
    return get_config_manager().get_setting(path, default)

def set_setting(path: str, value: Any) -> None:
    """
    Устанавливает значение настройки по указанному пути.
    
    Args:
        path: Путь к настройке в формате "section.subsection.key"
        value: Новое значение настройки
    """
    get_config_manager().set_setting(path, value)

def save_settings(preset_name: str) -> bool:
    """
    Сохраняет текущие настройки в файл пресета.
    
    Args:
        preset_name: Имя пресета
        
    Returns:
        True, если настройки успешно сохранены, иначе False
    """
    return get_config_manager().save_settings(preset_name)

def load_settings(preset_name: str) -> bool:
    """
    Загружает настройки из файла пресета.
    
    Args:
        preset_name: Имя пресета
        
    Returns:
        True, если настройки успешно загружены, иначе False
    """
    return get_config_manager().load_settings(preset_name)

def reset_settings() -> None:
    """
    Сбрасывает настройки до значений по умолчанию.
    """
    get_config_manager().reset_settings()

def get_presets_list() -> List[str]:
    """
    Возвращает список доступных пресетов.
    
    Returns:
        Список имен пресетов
    """
    return get_config_manager().get_presets_list()

def delete_preset(preset_name: str) -> bool:
    """
    Удаляет пресет.
    
    Args:
        preset_name: Имя пресета
        
    Returns:
        True, если пресет успешно удален, иначе False
    """
    return get_config_manager().delete_preset(preset_name)

def export_settings(file_path: str) -> bool:
    """
    Экспортирует текущие настройки в файл.
    
    Args:
        file_path: Путь к файлу для экспорта
        
    Returns:
        True, если настройки успешно экспортированы, иначе False
    """
    return get_config_manager().export_settings(file_path)

def import_settings(file_path: str) -> bool:
    """
    Импортирует настройки из файла.
    
    Args:
        file_path: Путь к файлу для импорта
        
    Returns:
        True, если настройки успешно импортированы, иначе False
    """
    return get_config_manager().import_settings(file_path) 
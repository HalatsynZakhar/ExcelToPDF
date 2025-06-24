import os
import logging
from typing import Any
from .config_manager import ConfigManager

# Настройка логгера
logger = logging.getLogger(__name__)

def get_downloads_folder():
    """Возвращает путь к папке с изображениями по умолчанию"""
    # Возвращаем сетевой путь вместо папки загрузок
    return r"\\10.10.100.2\Foto"
    
    # Закомментированный код ниже - оригинальная функция для получения папки загрузок
    # if os.name == 'nt':  # Windows
    #     import winreg
    #     sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    #     downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
    #     with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
    #         try:
    #             return winreg.QueryValueEx(key, downloads_guid)[0]
    #         except:
    #             return os.path.join(os.path.expanduser('~'), 'Downloads')
    # else:  # Linux, macOS и другие
    #     return os.path.join(os.path.expanduser('~'), 'Downloads')

# Глобальный экземпляр конфиг-менеджера
_config_manager = None

def init_config_manager(presets_folder: str) -> None:
    """
    Инициализирует экземпляр ConfigManager
    
    Args:
        presets_folder: Путь к папке с настройками
    """
    global _config_manager
    _config_manager = ConfigManager(presets_folder)
    
    # Пытаемся загрузить сохраненные настройки
    _config_manager.load_settings()

def get_config_manager() -> ConfigManager:
    """
    Возвращает текущий экземпляр ConfigManager
    
    Returns:
        Экземпляр ConfigManager
    """
    global _config_manager
    if _config_manager is None:
        raise RuntimeError("ConfigManager не инициализирован. Вызовите init_config_manager() перед использованием.")
    
    return _config_manager

def get_setting(path: str, default=None) -> Any:
    """
    Получает значение настройки по указанному пути
    
    Args:
        path: Путь к настройке в точечной нотации (например, "excel_settings.article_column")
        default: Значение по умолчанию, если настройка не найдена
        
    Returns:
        Значение настройки или default, если настройка не найдена
    """
    return get_config_manager().get_setting(path, default)

def set_setting(path: str, value: Any) -> None:
    """
    Устанавливает значение настройки по указанному пути
    
    Args:
        path: Путь к настройке в точечной нотации
        value: Устанавливаемое значение
    """
    get_config_manager().set_setting(path, value)

def save_settings() -> bool:
    """
    Сохраняет текущие настройки
    
    Returns:
        True, если настройки успешно сохранены, иначе False
    """
    return get_config_manager().save_settings()

def load_settings() -> bool:
    """
    Загружает настройки
    
    Returns:
        True, если настройки успешно загружены, иначе False
    """
    return get_config_manager().load_settings()

def reset_settings() -> None:
    """
    Сбрасывает настройки к значениям по умолчанию
    """
    get_config_manager().reset_settings()

def get_presets_list() -> list:
    """
    Возвращает список доступных пресетов настроек
    
    Returns:
        Список имен пресетов
    """
    return get_config_manager().get_presets_list()

def delete_preset(preset_name: str) -> bool:
    """
    Удаляет пресет с указанным именем
    
    Args:
        preset_name: Имя пресета для удаления
        
    Returns:
        True, если пресет успешно удален, иначе False
    """
    return get_config_manager().delete_preset(preset_name) 
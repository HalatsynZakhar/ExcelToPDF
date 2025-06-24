"""
ExcelWithImages - модуль для работы с Excel и добавления изображений в таблицы
"""
import logging
import importlib.util
import subprocess
import sys
import os
from pathlib import Path

# Функция для проверки и установки необходимых библиотек
def check_and_install_dependencies():
    """
    Проверяет наличие необходимых библиотек и устанавливает их при необходимости
    """
    # Переменная среды, указывающая, что мы уже установили зависимости
    if os.environ.get('EXCELWITHIMAGES_DEPS_INSTALLED') == '1':
        return
    
    # Список необходимых библиотек
    required_packages = [
        "openpyxl>=3.0.10",
        "Pillow>=9.0.0",
        "streamlit>=1.18.0",
        "numpy>=1.21.0",
        "pandas>=1.3.5",
        "watchdog>=2.1.0",
        "python-dotenv>=0.20.0"
    ]
    
    packages_to_install = []
    
    # Проверяем наличие каждой библиотеки
    for package in required_packages:
        package_name = package.split(">=")[0]
        
        try:
            # Пытаемся импортировать пакет напрямую
            if package_name == "Pillow":
                __import__("PIL")
            else:
                __import__(package_name)
        except ImportError:
            packages_to_install.append(package)
    
    # Если есть библиотеки для установки
    if packages_to_install:
        try:
            print("Обнаружены отсутствующие библиотеки. Начинаю установку...")
            
            for package in packages_to_install:
                package_name = package.split(">=")[0]
                print(f"Устанавливаю {package}...")
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                    print(f"Библиотека {package_name} успешно установлена")
                except subprocess.CalledProcessError as e:
                    print(f"Ошибка при установке библиотеки {package_name}: {e}")
            
            print("Все необходимые библиотеки установлены")
            
            # Устанавливаем переменную среды, чтобы избежать повторной установки
            os.environ['EXCELWITHIMAGES_DEPS_INSTALLED'] = '1'
        except Exception as e:
            print(f"Ошибка при установке библиотек: {e}")

# Настройка логирования
logs_dir = os.path.join(os.path.dirname(__file__), 'logs')
os.makedirs(logs_dir, exist_ok=True)

# Инициализация основного логгера
logger = logging.getLogger('ExcelWithImages')
logger.setLevel(logging.DEBUG)

# Версия пакета
__version__ = '0.1.0'

# Проверяем и устанавливаем зависимости при импорте
check_and_install_dependencies() 
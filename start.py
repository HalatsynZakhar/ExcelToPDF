#!/usr/bin/env python
"""
Скрипт запуска ExcelWithImages с выбором режима работы
"""
import os
import sys
import subprocess
import platform
import argparse
import shutil
from pathlib import Path

# Добавляем текущую директорию в PYTHONPATH
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Функция для проверки и создания структуры проекта
def ensure_project_structure():
    """
    Проверяет и создает необходимую структуру проекта
    """
    # Корневая директория проекта
    root_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Обязательные директории
    required_dirs = [
        "app",
        "examples",
        "examples/sample_data",
        "examples/output",
        "logs",
        "settings_presets",
        "utils",
        "utils/config_manager",
        "temp"
    ]
    
    # Создаем все необходимые директории
    for d in required_dirs:
        dir_path = os.path.join(root_dir, d)
        os.makedirs(dir_path, exist_ok=True)
        print(f"Директория {d} проверена")
    
    # Проверяем наличие важных файлов
    if not os.path.exists(os.path.join(root_dir, "utils/config_manager/__init__.py")):
        print("ВНИМАНИЕ: Отсутствует файл utils/config_manager/__init__.py")
    
    if not os.path.exists(os.path.join(root_dir, "utils/config_manager/config_manager.py")):
        print("ВНИМАНИЕ: Отсутствует файл utils/config_manager/config_manager.py")
    
    if not os.path.exists(os.path.join(root_dir, "app/app.py")):
        print("ВНИМАНИЕ: Отсутствует файл app/app.py")

# Функция для очистки временных файлов
def clean_temp_directory():
    """
    Удаляет все файлы в папке temp при запуске приложения
    """
    # Путь к директории temp
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
    
    # Проверяем существование директории
    if os.path.exists(temp_dir) and os.path.isdir(temp_dir):
        print("Очистка временных файлов...")
        # Перебираем все файлы в директории temp
        for filename in os.listdir(temp_dir):
            file_path = os.path.join(temp_dir, filename)
            try:
                # Если это файл, удаляем его
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                    print(f"Удален временный файл: {filename}")
                # Если это папка, удаляем её со всем содержимым
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    print(f"Удалена временная папка: {filename}")
            except Exception as e:
                print(f"Ошибка при удалении {file_path}: {e}")
        print("Очистка временных файлов завершена")
    else:
        print("Директория temp не существует или не является директорией")

# Функция для очистки консоли
def clear_screen():
    """Очищает экран консоли в зависимости от операционной системы"""
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')

# Функция для запуска веб-интерфейса
def start_web_app():
    """Запускает веб-интерфейс на Streamlit"""
    # Очищаем временные файлы перед запуском
    clean_temp_directory()
    
    # Проверяем и устанавливаем зависимости перед запуском приложения
    print("Проверка и установка необходимых зависимостей...")
    requirements_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requirements.txt")
    if os.path.exists(requirements_file):
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", requirements_file])
    else:
        print("Файл requirements.txt не найден, устанавливаем основные зависимости...")
        subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl>=3.0.10", "Pillow>=9.0.0", 
                        "streamlit>=1.18.0", "numpy>=1.21.0", "pandas>=1.3.5", "fpdf2>=2.7.4"])
    
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app", "app.py")
    
    if not os.path.exists(app_path):
        print(f"Ошибка: Файл приложения не найден: {app_path}")
        input("Нажмите Enter для продолжения...")
        return
        
    print("Запуск веб-интерфейса ExcelToPDF...")
    try:
        subprocess.run(["streamlit", "run", app_path])
    except Exception as e:
        print(f"Ошибка при запуске веб-интерфейса: {e}")
        input("Нажмите Enter для продолжения...")

def main():
    """Главная функция запуска"""
    ensure_project_structure()
    start_web_app()

if __name__ == "__main__":
    main()
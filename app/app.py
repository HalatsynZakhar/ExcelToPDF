import streamlit as st
import os
import sys
import logging
import io
import time
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from PIL import Image as PILImage
import json
import traceback
from fpdf import FPDF

# Добавляем корневую папку проекта в PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Используем относительные импорты вместо абсолютных
from utils import config_manager
from utils import excel_utils
from utils import image_utils
from utils.config_manager import get_downloads_folder, ConfigManager
# <<< ДОБАВЛЯЕМ ГЛОБАЛЬНЫЙ ИМПОРТ >>>
from core.processor import process_excel_file, create_pdf_cards

# Настройка логирования
log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)

# Ограничиваем количество файлов логов до 5 последних
log_files = sorted([f for f in os.listdir(log_dir) if f.startswith('app_')])
if len(log_files) > 5:
    for old_log in log_files[:-5]:
        try:
            os.remove(os.path.join(log_dir, old_log))
        except:
            pass

# Переименовываем текущий лог-файл, если он существует и создаем новый с правильной кодировкой
log_file = os.path.join(log_dir, 'app_latest.log')
# Всегда создаем новый лог-файл при запуске приложения
try:
    # Создаем новый файл с правильной кодировкой
    with open(log_file, 'w', encoding='utf-8') as f:
        f.write(f'{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - INFO - app - New log file created with UTF-8 encoding\n')
except Exception as e:
    print(f"Error creating log file: {e}")

log_stream = io.StringIO()
log_handler = logging.StreamHandler(log_stream)
log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
log_handler.setLevel(logging.INFO)

# Используем один файл лога для всего приложения
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
file_handler.setLevel(logging.DEBUG)

root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
# Удаляем существующие обработчики, если они есть
for handler in root_logger.handlers[:]:
    root_logger.removeHandler(handler)
root_logger.addHandler(log_handler)
root_logger.addHandler(file_handler)

# Устанавливаем кодировку для логирования
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

log = logging.getLogger(__name__)

# Определяем настройки по умолчанию
default_settings = {
    "paths": {
        "product_images_folder_path_1": get_downloads_folder(),
        "product_images_folder_path_2": "",
        "product_images_folder_path_3": "",
        "package_images_folder_path_1": "",
        "package_images_folder_path_2": "",
        "package_images_folder_path_3": ""
    },
    "excel_settings": {
        "article_column": "A"
    }
}

# Инициализация менеджера конфигурации с созданием настроек по умолчанию
def init_config_manager():
    """Инициализировать менеджер конфигурации и установить значения по умолчанию"""
    if 'config_manager' not in st.session_state:
        # Определяем путь к папке с пресетами
        presets_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
        
        # Инициализируем config manager с указанием папки пресетов
        config_manager_instance = config_manager.ConfigManager(presets_folder)
        
        # Устанавливаем значения по умолчанию, если они отсутствуют
        for i in range(1, 4):
            prod_key = f'paths.product_images_folder_path_{i}'
            pack_key = f'paths.package_images_folder_path_{i}'
            if not config_manager_instance.get_setting(prod_key):
                config_manager_instance.set_setting(prod_key, default_settings['paths'][f'product_images_folder_path_{i}'])
            if not config_manager_instance.get_setting(pack_key):
                config_manager_instance.set_setting(pack_key, default_settings['paths'][f'package_images_folder_path_{i}'])
        
        if not config_manager_instance.get_setting('excel_settings.article_column'):
            config_manager_instance.set_setting('excel_settings.article_column', default_settings['excel_settings']['article_column'])
        
        # Сохраняем конфигурацию
        config_manager_instance.save_settings("Default")
        
        # Сохраняем менеджер в session_state
        st.session_state.config_manager = config_manager_instance
        
        log.info("Менеджер конфигурации инициализирован с настройками по умолчанию")
    
    return st.session_state.config_manager

def get_downloads_folder():
    """Получить путь к папке с изображениями по умолчанию"""
    # Возвращаем сетевой путь вместо папки загрузок
    return r"\\10.10.100.2\Foto"
    
    # Закомментированный код ниже - оригинальная функция для получения папки загрузок
    # if platform.system() == "Windows":
    #     import winreg
    #     sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    #     downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
    #     with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
    #         downloads_folder = winreg.QueryValueEx(key, downloads_guid)[0]
    #         return downloads_folder
    # elif platform.system() == "Darwin":  # macOS
    #     return os.path.join(os.path.expanduser('~'), 'Downloads')
    # else:  # Linux и другие системы
    #     return os.path.join(os.path.expanduser('~'), 'Downloads')

# Обновляем код инициализации для использования нашей функции
config_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
# Инициализируем глобальный config_manager в модуле config_manager перед инициализацией нашего
config_manager.init_config_manager(config_folder)
init_config_manager()

# Настройка параметров приложения
st.set_page_config(
    page_title="Excel to PDF Card Generator",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Функция для создания временных директорий
def ensure_temp_dir(prefix: str = "") -> str:
    """
    Создает и возвращает путь к временной директории.
    
    Args:
        prefix (str): Префикс для имени временной директории
    
    Returns:
        Путь к временной директории
    """
    # Создаем временную директорию в папке проекта для лучшего доступа
    project_dir = os.path.dirname(os.path.dirname(__file__))
    temp_dir = os.path.join(project_dir, "temp")
    
    # Создаем директорию, если она не существует
    try:
        os.makedirs(temp_dir, exist_ok=True)
        log.info(f"Создана/проверена временная директория: {temp_dir}")
    except Exception as e:
        log.error(f"Ошибка при создании временной директории {temp_dir}: {e}")
        # Если не удалось создать в проекте, используем системную временную директорию
        temp_dir = os.path.join(tempfile.gettempdir(), f"{prefix}excelwithimages")
        try:
            os.makedirs(temp_dir, exist_ok=True)
            log.info(f"Использована системная временная директория: {temp_dir}")
        except Exception as e:
            log.error(f"Ошибка при создании системной временной директории {temp_dir}: {e}")
            # Если и системная не удалась, выбрасываем исключение
            raise RuntimeError("Не удалось создать временную директорию") from e
    
    return temp_dir

# Функция для очистки временных файлов
def cleanup_temp_files():
    """
    Очищает временные файлы, сохраняя только файлы текущей сессии.
    """
    try:
        # Определяем путь к временной директории
        temp_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'temp')
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir, exist_ok=True)
            log.info(f"Создана временная директория: {temp_dir}")
            return
        
        # Получаем время начала текущей сессии (приложение запущено)
        session_start_time = datetime.now()
        
        # Максимальный возраст файлов, которые мы хотим сохранить (в минутах)
        # Сохраняем только файлы, созданные в течение последнего часа
        max_age_minutes = 60
        
        # Файлы для сохранения (используемые в текущей сессии)
        files_to_keep = [
            st.session_state.get('temp_file_path', ''),
            st.session_state.get('output_file_path', '')
        ]
        
        # Получаем список всех файлов в временной директории
        all_files = os.listdir(temp_dir)
        log.info(f"Найдено {len(all_files)} файлов в директории {temp_dir}")
        
        # Удаляем старые файлы, которые не используются в текущей сессии
        removed_count = 0
        for filename in all_files:
            file_path = os.path.join(temp_dir, filename)
            
            # Пропускаем, если это не файл
            if not os.path.isfile(file_path):
                continue
                
            # Проверяем, используется ли файл в текущей сессии
            if file_path in files_to_keep:
                log.info(f"Сохраняем файл текущей сессии: {file_path}")
                continue
                
            # Получаем время последней модификации файла
            try:
                file_mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                file_age = session_start_time - file_mod_time
                
                # Если файл старше максимального возраста или не из текущей сессии
                if file_age.total_seconds() > (max_age_minutes * 60):
                    try:
                        os.remove(file_path)
                        removed_count += 1
                        log.info(f"Удален старый временный файл: {file_path} (возраст: {file_age})")
                    except Exception as e:
                        log.error(f"Ошибка при удалении файла {file_path}: {e}")
            except Exception as e:
                log.error(f"Ошибка при проверке времени файла {file_path}: {e}")
                    
        log.info(f"Очистка временных файлов завершена. Удалено {removed_count} файлов.")
    except Exception as e:
        log.error(f"Ошибка при очистке временных файлов: {e}")

# Вызываем очистку временных файлов при запуске приложения
cleanup_temp_files()

# Функция для добавления сообщения в лог сессии
def add_log_message(message, level="INFO"):
    """
    Добавляет сообщение в лог сессии с временной меткой.
    
    Args:
        message (str): Сообщение для добавления
        level (str): Уровень сообщения (INFO, WARNING, ERROR, SUCCESS)
    """
    if 'log_messages' not in st.session_state:
        st.session_state.log_messages = []
    
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.log_messages.append(f"[{timestamp}] [{level}] {message}")
    
    # Ограничиваем размер лога
    if len(st.session_state.log_messages) > 100:
        st.session_state.log_messages = st.session_state.log_messages[-100:]
    
    # Также добавляем в обычный лог
    if level == "ERROR":
        log.error(message)
    elif level == "WARNING":
        log.warning(message)
    else:
        log.info(message)

# Функция для отображения настроек
def show_settings():
    """Отображает блок с пользовательскими настройками в сайдбаре"""
    with st.sidebar:
        st.header("⚙️ Настройки")
        cm = st.session_state.config_manager

        with st.expander("Папки с изображениями товаров", expanded=True):
            st.markdown("Укажите до 3-х папок. Поиск будет идти по порядку.")
            for i in range(1, 4):
                key = f'paths.product_images_folder_path_{i}'
                current_path = cm.get_setting(key, "")
                new_path = st.text_input(f"Папка товаров {i}", value=current_path, key=f"product_folder_{i}")
                if new_path != current_path:
                    cm.set_setting(key, new_path)
                    cm.save_settings("Default")
                    st.rerun()

        with st.expander("Папки с изображениями упаковок", expanded=True):
            st.markdown("Укажите до 3-х папок с изображениями упаковок.")
            for i in range(1, 4):
                key = f'paths.package_images_folder_path_{i}'
                current_path = cm.get_setting(key, "")
                new_path = st.text_input(f"Папка упаковок {i}", value=current_path, key=f"package_folder_{i}")
                if new_path != current_path:
                    cm.set_setting(key, new_path)
                    cm.save_settings("Default")
                    st.rerun()

# Функция для загрузки Excel файла
def load_excel_file(uploaded_file_arg=None):
    # Используем файл из session_state, если аргумент не передан (для on_change)
    uploaded_file = uploaded_file_arg if uploaded_file_arg else st.session_state.get('file_uploader')
    if not uploaded_file:
        # Если файл удален из загрузчика
        log.warning("Файл был удален из загрузчика.")
        st.session_state.available_sheets = []
        st.session_state.selected_sheet = None
        st.session_state.df = None
        st.session_state.temp_file_path = None
        st.session_state.processing_error = None
        return

    # Используем временный путь из session_state
    temp_file_path = st.session_state.get('temp_file_path')
    if not temp_file_path or not os.path.exists(temp_file_path):
        log.error("Временный путь к файлу отсутствует или файл не найден.")
        st.session_state.processing_error = "Ошибка: временный файл не найден. Попробуйте загрузить файл заново."
        st.session_state.available_sheets = []
        st.session_state.selected_sheet = None
        st.session_state.df = None
        return
        
    try:
        log.info(f"Загрузка листов из файла: {temp_file_path}")
        excel_file = pd.ExcelFile(temp_file_path, engine='openpyxl')
        all_sheets = excel_file.sheet_names
        
        # Фильтруем листы, исключая листы с макросами
        filtered_sheets = [sheet for sheet in all_sheets if not sheet.startswith('xl/macrosheets/')]
        st.session_state.available_sheets = filtered_sheets
        log.info(f"Все листы: {all_sheets}")
        log.info(f"Доступные листы (без макросов): {st.session_state.available_sheets}")
        
        # Проверяем, были ли отфильтрованы листы с макросами
        if len(all_sheets) > len(filtered_sheets):
            log.warning(f"Обнаружены и отфильтрованы листы с макросами: {set(all_sheets) - set(filtered_sheets)}")
            # Если все листы были с макросами и отфильтрованы
            if not filtered_sheets:
                st.session_state.processing_error = "Внимание! Этот файл Excel содержит только макросы, а не обычные таблицы данных. Пожалуйста, выберите файл Excel с обычными листами, содержащими таблицы с артикулами и данными для обработки."
                return
        
        # --- Выбор листа по умолчанию --- 
        current_selection = st.session_state.get('selected_sheet')
        default_sheet = None
        if st.session_state.available_sheets:
            # Пытаемся найти первый "обычный" лист (не пустой, не скрытый - openpyxl может понадобиться для скрытых)
            # Простой вариант: просто берем первый
            default_sheet = st.session_state.available_sheets[0]
            log.info(f"Лист по умолчанию выбран: {default_sheet}")

        # Устанавливаем лист по умолчанию, если он еще не выбран или текущий выбор невалиден
        if default_sheet and (not current_selection or current_selection not in st.session_state.available_sheets):
             st.session_state.selected_sheet = default_sheet
             # Устанавливаем sheet_selector для корректной работы handle_sheet_change
             st.session_state.sheet_selector = default_sheet
             log.info(f"Установлен активный лист: {st.session_state.selected_sheet}")
             # Сбрасываем DataFrame, т.к. лист изменился (или был установлен впервые)
             st.session_state.df = None 
             st.session_state.processing_error = None

        # --- Загрузка данных с выбранного листа (если он есть) ---
        # Вызываем handle_sheet_change, чтобы загрузить данные для ВЫБРАННОГО листа
        # (это также обработает случай, когда лист был только что установлен по умолчанию)
        if st.session_state.selected_sheet:
            handle_sheet_change()  # Эта функция загрузит df и обработает ошибки
        else:
             # Если листов нет или выбрать по умолчанию не удалось
             st.session_state.df = None
             st.session_state.processing_error = "В файле не найдено листов для обработки."
             log.warning("Не удалось выбрать лист по умолчанию или листы отсутствуют.")

    except Exception as e:
        error_msg = f"Ошибка при чтении листов из Excel-файла: {e}"
        log.error(error_msg, exc_info=True)
        st.session_state.processing_error = error_msg
        st.session_state.available_sheets = []
        st.session_state.selected_sheet = None
        st.session_state.df = None

# Проверка валидности всех входных данных перед обработкой
def all_inputs_valid():
    """
    Проверяет, что все необходимые данные для обработки заполнены и валидны.
    
    Returns:
        bool: True, если все входные данные валидны, иначе False
    """
    # Подробная проверка с логированием
    valid = True
    log_msgs = []
    
    # 1. Проверяем наличие DataFrame
    if st.session_state.get('df') is None:
        log_msgs.append("DataFrame не загружен")
        valid = False
    else:
        log_msgs.append(f"DataFrame загружен, размер: {st.session_state.df.shape}")
        
    # 2. Проверяем, выбран ли лист в Excel
    if st.session_state.get('selected_sheet') is None:
        log_msgs.append("Лист Excel не выбран")
        valid = False
    else:
        log_msgs.append(f"Выбран лист: {st.session_state.selected_sheet}")

    # 3. Проверяем, выбрана ли колонка с артикулами
    if not st.session_state.get('article_column'): # Проверяем наличие и непустое значение
        log_msgs.append("Колонка с артикулами не выбрана")
        valid = False
    else:
        # Проверяем, что обозначение колонки - буква или число
        article_col = st.session_state.article_column
        if not (article_col.isalpha() or article_col.isdigit()):
            log_msgs.append(f"Неверное обозначение колонки с артикулами: '{article_col}'. Используйте букву (A, B, C...) или номер (1, 2, 3...)")
            valid = False
        else:
            log_msgs.append(f"Выбрана колонка артикулов: {article_col} ({article_col if article_col.isdigit() else f'столбец {article_col}'})")

    # 4. Проверяем папку с изображениями
    images_folder = config_manager.get_setting("paths.product_images_folder_path_1", "")
    if not images_folder:
        log_msgs.append("Папка с изображениями не указана в настройках")
        valid = False
    elif not os.path.exists(images_folder):
        log_msgs.append(f"Папка с изображениями не найдена: {images_folder}")
        valid = False
    else:
        log_msgs.append(f"Папка с изображениями найдена: {images_folder}")

    # Логируем результат проверки
    final_msg = "Проверка валидности завершена. Результат: " + ("Успешно" if valid else "Неуспешно")
    log.info(final_msg)
    for msg in log_msgs:
        log.info(f"- {msg}")
        
    return valid

def trigger_processing():
    """Sets a flag to start processing."""
    st.session_state.start_processing = True

# Функция для обработки изменения выбранного листа
def handle_sheet_change():
    """
    Обрабатывает изменение выбранного листа Excel и перезагружает данные.
    """
    # Обновляем выбранный лист из селектора, если он был изменен
    if 'sheet_selector' in st.session_state and st.session_state.get("sheet_selector") != st.session_state.selected_sheet:
        st.session_state.selected_sheet = st.session_state.get("sheet_selector")
        log.info(f"Выбран новый лист из селектора: {st.session_state.selected_sheet}")
    
    # Проверяем, что у нас есть выбранный лист
    selected_sheet = st.session_state.get('selected_sheet')
    if not selected_sheet:
        log.warning("Не выбран лист для загрузки данных")
        st.session_state.df = None
        st.session_state.processing_error = "Не выбран лист для загрузки данных"
        return
        
    # Перезагружаем данные с выбранного листа
    if st.session_state.temp_file_path and os.path.exists(st.session_state.temp_file_path):
        try:
            log.info(f"Загрузка данных с листа: {selected_sheet}")
            
            # Всегда используем фиксированные значения: без пропуска строк и заголовок в первой строке
            df = pd.read_excel(
                st.session_state.temp_file_path, 
                sheet_name=selected_sheet, 
                engine='openpyxl',
                skiprows=0,
                header=None
            )
            
            # Преобразуем все столбцы с объектами в строки для предотвращения ошибок с pyarrow
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].astype(str)
            
            # Проверка на пустой DataFrame
            log.info(f"Размер данных при смене листа: строк={df.shape[0]}, колонок={df.shape[1]}; пустой={df.empty}")
            
            if df.empty:
                error_msg = f"Лист '{selected_sheet}' не содержит данных"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
            
            if df.shape[0] == 0:
                error_msg = f"Лист '{selected_sheet}' не содержит строк с данными"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
                
            if df.shape[1] == 0:
                error_msg = f"Лист '{selected_sheet}' не содержит колонок с данными"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
            
            # Проверка на файл, который имеет колонки, но все значения в них NaN
            if df.notna().sum().sum() == 0:
                error_msg = f"Лист '{selected_sheet}' содержит только пустые ячейки"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
            
            # Все хорошо, сохраняем DataFrame
            st.session_state.df = df
            st.session_state.processing_error = None
            log.info(f"Лист '{selected_sheet}' успешно загружен. Найдено {len(df)} строк и {len(df.columns)} колонок")
            
        except Exception as e:
            error_msg = f"Ошибка при загрузке листа '{selected_sheet}': {str(e)}"
            log.error(error_msg)
            
            # Делаем сообщение об ошибке более понятным для пользователя
            user_friendly_msg = error_msg
            if "'dict' object has no attribute 'shape'" in str(e):
                user_friendly_msg = f"Лист '{selected_sheet}' не содержит табличных данных. Пожалуйста, выберите лист с необходимыми данными."
            elif "No sheet" in str(e) or "not found" in str(e):
                user_friendly_msg = f"Лист '{selected_sheet}' не найден в файле. Пожалуйста, выберите существующий лист."
            elif "Empty" in str(e) or "no data" in str(e):
                user_friendly_msg = f"Лист '{selected_sheet}' не содержит данных. Пожалуйста, выберите лист с данными."
            elif "ArrowTypeError" in str(e) or "Expected bytes" in str(e):
                user_friendly_msg = f"Ошибка преобразования типов данных. Попробуйте выбрать другой лист или перезагрузить файл."
                
            st.session_state.processing_error = user_friendly_msg
            st.session_state.df = None

# Функция для загрузки файла Excel
def file_uploader_section():
    """
    Отображает секцию для загрузки файла Excel.
    """
    with st.container():
        st.write("## Загрузка файла Excel")
        
        # CSS стили для кнопок и сообщений
        st.markdown("""
        <style>
        /* Стили для большой зеленой кнопки */
        .big-button-container {
            display: flex;
            justify-content: center;
            margin: 20px 0;
        }
        
        /* Увеличиваем высоту кнопок */
        .stButton > button:not([kind="secondary"]) {
            height: 80px !important;
            font-size: 20px !important;
            padding: 20px !important;
            width: 100% !important;
        }
        
        /* Специфичные стили для кнопки скачивания */
        div[data-testid="stDownloadButton"] button {
            height: 100px !important;
            font-size: 24px !important;
            padding: 25px !important;
            width: 100% !important;
            background-color: #4CAF50 !important;
            color: white !important;
            border: none !important;
            border-radius: 8px !important;
            transition: all 0.3s ease !important;
        }
        
        div[data-testid="stDownloadButton"] button:hover {
            background-color: #45a049 !important;
            transform: scale(1.02) !important;
        }
        
        /* Стиль для сообщений об ошибках */
        .error-message {
            color: #cc0000;
            background-color: #ffeeee;
            padding: 10px;
            border-radius: 5px;
            border-left: 5px solid #cc0000;
            margin: 10px 0;
        }
        
        /* Стиль для индикатора количества строк */
        .row-count {
            font-weight: bold;
            color: #1f77b4;
        }
        
        /* Стили для улучшения внешнего вида загрузчика файлов */
        div[data-testid="stFileUploader"] {
            border: 1px dashed #cccccc;
            padding: 10px;
            border-radius: 5px;
            background-color: #f8f9fa;
        }
        
        div[data-testid="stFileUploader"]:hover {
            border-color: #4CAF50;
            background-color: #f0f9f0;
        }
        
        /* Стили для лога */
        .log-container {
            max-height: 300px;
            overflow-y: auto;
            font-family: monospace;
            background-color: #f5f5f5;
            padding: 10px;
            border-radius: 5px;
            border: 1px solid #ddd;
            margin-top: 15px;
        }
        
        .log-entry {
            margin: 2px 0;
            font-size: 12px;
        }
        
        .log-info {
            color: #0366d6;
        }
        
        .log-warning {
            color: #e36209;
        }
        
        .log-error {
            color: #d73a49;
        }
        
        .log-success {
            color: #22863a;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Стили для кнопок
        big_green_button_style = """
            background-color: #4CAF50;
            color: white;
            padding: 15px 32px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            margin: 4px 2px;
            cursor: pointer;
            border-radius: 8px;
            border: none;
            box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
            transition: 0.3s;
        """
        
        inactive_button_style = """
            background-color: #cccccc;
            color: #666666;
            padding: 15px 32px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            margin: 4px 2px;
            cursor: not-allowed;
            border-radius: 8px;
            border: none;
            box-shadow: none;
        """
        
        # Загрузчик файлов Excel
        uploaded_file = st.file_uploader("Выберите Excel файл для обработки", type=["xlsx", "xls"], key="file_uploader",
                                     on_change=load_excel_file)

        # Отображение информации о загруженном файле
        if uploaded_file is not None:
            st.write(f"**Загружен файл:** {uploaded_file.name}")
            
            current_temp_path = st.session_state.get('temp_file_path', '')
            current_file_size = uploaded_file.size
            
            # Проверка необходимости обновления файла
            need_update = False
            
            # Проверяем, требуется ли обновление файла
            if not current_temp_path or not os.path.exists(current_temp_path):
                # Файла еще нет, нужно сохранить
                need_update = True
                log.info(f"Файл отсутствует, сохраняем новый: {uploaded_file.name}")
            elif os.path.basename(current_temp_path) != uploaded_file.name:
                # Имя файла изменилось, нужно сохранить новый
                need_update = True
                log.info(f"Имя файла изменилось: {os.path.basename(current_temp_path)} -> {uploaded_file.name}")
            else:
                # Файл с таким же именем уже существует, проверяем размер
                try:
                    previous_size = os.path.getsize(current_temp_path)
                    if previous_size != current_file_size:
                        # Размер изменился, заменяем файл
                        need_update = True
                        log.info(f"Размер файла изменился: {previous_size} -> {current_file_size}")
                        try:
                            os.remove(current_temp_path)
                            log.info(f"Удален предыдущий файл: {current_temp_path}")
                        except Exception as e:
                            log.error(f"Ошибка при удалении предыдущего файла: {e}")
                except Exception as e:
                    log.error(f"Ошибка при проверке размера файла {current_temp_path}: {e}")
                    need_update = True
            
            # Если требуется обновление, сохраняем файл
            if need_update:
                temp_dir = ensure_temp_dir()
                
                # Очищаем промежуточные файлы с префиксом temp_full_
                try:
                    # Удаляем все временные файлы с префиксом temp_full_
                    output_folder = os.path.dirname(temp_dir)
                    for filename in os.listdir(output_folder):
                        if filename.startswith("temp_full_"):
                            filepath = os.path.join(output_folder, filename)
                            try:
                                if os.path.isfile(filepath):
                                    os.remove(filepath)
                                    log.info(f"Удален промежуточный файл: {filepath}")
                            except Exception as e:
                                log.error(f"Ошибка при удалении промежуточного файла {filepath}: {e}")
                except Exception as e:
                    log.error(f"Ошибка при очистке промежуточных файлов: {e}")
                    
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.session_state.temp_file_path = temp_file_path
                add_log_message(f"Файл сохранен: {os.path.basename(temp_file_path)}", "INFO")
                load_excel_file()
            
            # Удалены настройки пропуска начальных строк и строки с заголовками
            # Инициализация переменных с фиксированными значениями
            st.session_state.skiprows = 0
            st.session_state.header_row = 0
                    
            # Отображение ошибки обработки, если есть
            if st.session_state.processing_error:
                st.markdown(f"""
                <div class="error-message">
                    <strong>Ошибка:</strong> {st.session_state.processing_error}
                </div>
                """, unsafe_allow_html=True)
                
                # Добавляем подсказку для решения проблемы с пустыми данными
                if "не содержит данных" in st.session_state.processing_error or "содержит только пустые ячейки" in st.session_state.processing_error:
                    st.info("""
                    **Рекомендации по решению проблемы:**
                    
                    1. Убедитесь, что файл Excel содержит данные в выбранном листе
                    2. Проверьте наличие невидимых форматирований или скрытых строк
                    3. Попробуйте открыть файл в Excel и пересохранить его
                    4. Убедитесь, что данные начинаются с первой строки и колонки
                    """)
                
            # Если данные успешно загружены, показываем предпросмотр и селекторы колонок
            if st.session_state.df is not None:
                # Отображение размерности данных
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"""
                    <div class="row-count">
                        Количество строк: {st.session_state.df.shape[0]}
                    </div>
                    """, unsafe_allow_html=True)
                with col2:
                    st.write(f"**Количество колонок:** {st.session_state.df.shape[1]}")
                
                # Добавляем предпросмотр данных
                with st.expander("Предпросмотр данных", expanded=False):
                    st.dataframe(st.session_state.df.head(10), use_container_width=True)
                    
                    # Добавляем статистику по колонкам
                    col_stats = pd.DataFrame({
                        'Колонка': st.session_state.df.columns,
                        'Тип данных': [str(dtype) for dtype in st.session_state.df.dtypes.values],
                        'Непустых значений': st.session_state.df.count().values,
                        'Процент заполнения': (st.session_state.df.count() / len(st.session_state.df) * 100).round(2).values
                    })
                    st.write("### Статистика по колонкам")
                    st.dataframe(col_stats, use_container_width=True)
                
                # Получение списка колонок
                column_options = list(st.session_state.df.columns)
                
                # Если колонки есть, показываем селекторы
                if column_options:
                    # Определяем индексы по умолчанию (если колонки A/B существуют)
                    default_article_index = column_options.index("A") if "A" in column_options else 0
                    
                    # Позволяем пользователю ввести буквенные или числовые обозначения колонок
                    selected_article_col = st.text_input(
                        "Колонка с артикулами", 
                        value=st.session_state.get('article_column', 'A'),
                        key="article_column_input",
                        help="Введите букву (A, B, C...) или номер (1, 2, 3...) колонки, содержащей артикулы товаров"
                    )
                    st.caption("Примеры: 'A' или '1', 'B' или '2'")
                    st.session_state.article_column = selected_article_col
                    
                    # Проверка всех необходимых полей перед обработкой
                    process_button_disabled = not all_inputs_valid()
                    
                    # Кнопка для запуска обработки
                    st.button("Обработать файл", 
                              disabled=process_button_disabled, 
                              type="primary", 
                              key="process_button",
                              on_click=trigger_processing,
                              use_container_width=True)  # Добавляем параметр для растягивания на всю ширину
                    
                    # Запускаем обработку, если установлен флаг
                    if st.session_state.get('start_processing', False):
                        st.info("Идет обработка файла. Не закрывайте страницу и не взаимодействуйте с интерфейсом до завершения.")
                        st.write("Это может занять некоторое время в зависимости от количества строк и изображений.")
                        
                        # Блокируем интерфейс на время обработки
                        with st.spinner("Обработка файла..."):
                            # Очищаем предыдущие результаты и ошибки
                            st.session_state.processing_result = None
                            st.session_state.processing_error = None
                            
                            # Выполняем обработку
                            success = process_files()
                            
                            # Записываем результат в session_state для отображения после перезагрузки
                            if success:
                                st.session_state.processing_result = "Файл успешно обработан! Вы можете скачать его ниже."
                                # Устанавливаем флаг для автоматического скролла к секции скачивания после перезагрузки
                                st.session_state.scroll_to_download = True
                            else:
                                st.session_state.processing_error_message = st.session_state.processing_error
                        
                        # Сбрасываем флаг обработки
                        st.session_state.start_processing = False
                        
                        # Форсируем перезагрузку страницы для обновления UI
                        st.rerun()
                    
            else:
                st.warning("Файл не содержит колонок для выбора. Проверьте структуру Excel-файла.")
                
        else:
                    st.warning("Файл не содержит колонок для выбора. Проверьте структуру Excel-файла.")
                    
        # Отображение результатов обработки и ошибок после обработки файла
        # Показываем только если обработка не выполняется сейчас
        if not st.session_state.get('start_processing', False):
            # Успешное завершение обработки
            if st.session_state.get('processing_result'):
                st.success(st.session_state.processing_result)
                # Если нужно автоматически прокрутить к секции скачивания
                if st.session_state.get('scroll_to_download', False):
                    st.markdown('<script>setTimeout(function() { window.scrollTo(0, document.body.scrollHeight); }, 500);</script>', 
                                unsafe_allow_html=True)
                    # Сбрасываем флаг скролла
                    st.session_state.scroll_to_download = False
    
            # Ошибка обработки
            if st.session_state.get('processing_error_message'):
                st.error(f"Ошибка при обработке файла: {st.session_state.processing_error_message}")
                # Очищаем сообщение об ошибке после отображения
                st.session_state.processing_error_message = None
            
            # Добавление кнопки скачивания, если файл был обработан
            if st.session_state.output_file_path and os.path.exists(st.session_state.output_file_path):
                # Создаем колонку для центрирования кнопки (опционально, для лучшего вида)
                col1, col2, col3 = st.columns([1,2,1])
                with col2:
                    with open(st.session_state.output_file_path, "rb") as file:
                        st.download_button(
                            label="СКАЧАТЬ ОБРАБОТАННЫЙ ФАЙЛ",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_file_path),
                            mime="application/pdf",
                            use_container_width=True,
                            type="primary",
                            key="download_button"
                        )
        
        # Проверяем, нужно ли отобразить отчет о результатах обработки
        if st.session_state.get('show_processing_report', False):
            # Удаляем вызов функции отображения отчета, поскольку функционал аналитики больше не требуется
            # Просто сбрасываем флаг, чтобы не пытаться отображать отчет повторно
            st.session_state.show_processing_report = False
        
        # Добавляем отображение логов вместо отладочной информации
        with st.expander("Журнал событий", expanded=False):
            # Отображаем сообщения из st.session_state.log_messages
            if 'log_messages' in st.session_state and st.session_state.log_messages:
                st.markdown('<div class="log-container">', unsafe_allow_html=True)
                for log_msg in st.session_state.log_messages:
                    # Определяем класс для стилизации
                    log_class = "log-info"
                    if "ERROR" in log_msg:
                        log_class = "log-error"
                    elif "WARNING" in log_msg:
                        log_class = "log-warning"
                    elif "SUCCESS" in log_msg:
                        log_class = "log-success"
                        
                    st.markdown(f'<div class="log-entry {log_class}">{log_msg}</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info("Журнал пуст")

# Функция для обработки файла
def process_files():
    """
    Основная функция для обработки файлов.
    Создает PDF-файл с карточками товаров.
    """
    try:
        log.info("===================== НАЧАЛО ОБРАБОТКИ ФАЙЛА =====================")
        add_log_message("Начало обработки файла", "INFO")
        st.session_state.is_processing = True
        st.session_state.processing_result = None
        st.session_state.processing_error = None
        
        with st.spinner("Идет обработка файла. Пожалуйста, подождите..."):
            cm = st.session_state.config_manager
            
            # Получаем пути к папкам с изображениями товаров
            product_image_folders = [
                cm.get_setting('paths.product_images_folder_path_1'),
                cm.get_setting('paths.product_images_folder_path_2'),
                cm.get_setting('paths.product_images_folder_path_3')
            ]
            
            # Получаем пути к папкам с изображениями упаковок
            package_image_folders = [
                cm.get_setting('paths.package_images_folder_path_1'),
                cm.get_setting('paths.package_images_folder_path_2'),
                cm.get_setting('paths.package_images_folder_path_3')
            ]

            df = st.session_state.df
            article_col = st.session_state.get('article_column')
            
            if df is None or article_col is None:
                st.session_state.processing_error = "Не загружен файл Excel или не выбран столбец с артикулами."
                return False
                
            temp_dir = ensure_temp_dir()
            
            output_path, inserted_cards, not_found_articles = create_pdf_cards(
                df=df,
                article_col_name=article_col,
                product_image_folders=product_image_folders,
                package_image_folders=package_image_folders,
                output_folder=temp_dir,
                progress_callback=lambda current, total: add_log_message(f"Обработано {current} из {total} строк", "INFO")
            )

            st.session_state.output_file_path = output_path
            
            if inserted_cards > 0:
                success_msg = f"Обработка завершена. Создано карточек: {inserted_cards}."
                st.session_state.processing_result = success_msg
                log.info(success_msg)
                add_log_message(success_msg, "SUCCESS")
                
                if not_found_articles:
                    st.session_state.not_found_articles = not_found_articles
                    warning_msg = f"Не найдены изображения для {len(not_found_articles)} артикулов."
                    add_log_message(warning_msg, "WARNING")
                
                log.info("===================== ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО =====================")
                return True
            else:
                error_msg = f"Не удалось создать PDF. Не найдено ни одного изображения для артикулов в указанном столбце."
                if not_found_articles:
                    error_msg += f" (проверено {len(not_found_articles)} артикулов)."
                st.session_state.processing_error = error_msg
                log.warning(error_msg)
                add_log_message(error_msg, "WARNING")
                log.info("===================== ОБРАБОТКА ЗАВЕРШЕНА (КАРТОЧКИ НЕ СОЗДАНЫ) =====================")
                return False

    except Exception as e:
        error_msg = f"Ошибка при создании PDF: {e}"
        st.session_state.processing_error = error_msg
        log.error(error_msg, exc_info=True)
        add_log_message(error_msg, "ERROR")
        log.info("===================== ОБРАБОТКА ЗАВЕРШЕНА С ОШИБКОЙ =====================")
        return False
    finally:
        st.session_state.is_processing = False

def show_results(stats: Dict[str, Any]):
    """
    Отображает результаты обработки файла.
    """
    st.write("**Результаты обработки:**")
    st.write(f"**Создано карточек:** {stats['inserted_cards']}")
    st.write(f"**Не найдены изображения для:** {', '.join(stats['not_found_articles'])}")

def initialize_session_state():
    """Инициализирует переменные в session state, если их нет"""
    # Словарь с переменными и их значениями по умолчанию
    defaults = {
        'df': None,
        'temp_file_path': None,
        'processing_result': None,
        'processing_error': None,
        'is_processing': False,
        'output_file_path': None,
        'selected_sheet': None,
        'available_sheets': [],
        'log_messages': [],
        'article_column': 'A',
        'start_processing': False,
        'show_processing_report': False,
        'processing_error_message': None,
        'scroll_to_download': False,
        'not_found_articles': []
    }
    # Проходим по словарю и инициализируем переменные, если их нет
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def check_required_modules():
    """
    Заглушка для функции проверки модулей.
    Эта функциональность обрабатывается в start.py
    """
    pass

def main():
    """Основная функция для запуска веб-приложения"""
    
    # Инициализация session_state
    initialize_session_state()

    # Заголовок приложения
    st.title("📇 Генератор PDF-карточек из Excel")
    st.write("Загрузите ваш Excel-файл, выберите столбец с артикулами, и приложение создаст PDF-документ с карточками товаров.")

    # Получаем менеджер конфигурации
    cm = init_config_manager()
    
    # Проверяем наличие необходимых модулей
    check_required_modules()

    # Отображаем UI
    show_settings()
    file_uploader_section()

if __name__ == "__main__":
    main()

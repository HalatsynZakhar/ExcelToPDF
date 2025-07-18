# ImageProcessor - Набор инструментов для обработки изображений и данных

## Описание проекта

ImageProcessor - это комплексный набор инструментов для работы с изображениями и данными, состоящий из нескольких взаимосвязанных подпроектов. Каждый подпроект решает специфические задачи в области обработки изображений, работы с Excel-файлами и создания PDF-документов.

## Структура проекта

Проект состоит из следующих основных компонентов:

### 1. PhotoProcessor

Инструмент для обработки фотографий и изображений. Позволяет выполнять различные операции с изображениями, включая изменение размера, обрезку, применение фильтров и эффектов, пакетную обработку и многое другое.

### 2. ExcelToPDF

Конвертер данных из Excel-файлов в структурированные PDF-карточки товаров. Идеально подходит для создания каталогов продукции, прайс-листов и других маркетинговых материалов. Подробнее см. в [README.md проекта ExcelToPDF](./ExcelToPDF/README.md).

### 3. ExcelWithImages

Инструмент для работы с изображениями в Excel-файлах. Позволяет автоматизировать процесс добавления, извлечения и обработки изображений в Excel-документах.

## Общие компоненты

Все подпроекты используют общие компоненты и библиотеки:

- **Управление конфигурацией**: Унифицированная система управления настройками и конфигурациями
- **Утилиты для работы с изображениями**: Общие функции для обработки изображений
- **Утилиты для работы с Excel**: Общие функции для работы с Excel-файлами

## Технические особенности

### Архитектура

Проект построен по модульному принципу, что позволяет использовать компоненты как вместе, так и по отдельности. Каждый подпроект имеет свою собственную структуру и может быть запущен независимо от других.

### Единый интерфейс и принципы работы

Все подпроекты следуют единым принципам работы, унаследованным от ExcelWithImages:

- **Унифицированный пользовательский интерфейс**: Все подпроекты используют схожую структуру веб-интерфейса на базе Streamlit с общими элементами управления
- **Единый подход к обработке файлов**: Механизмы загрузки, обработки и сохранения файлов реализованы по общему шаблону
- **Общая логика поиска и сопоставления изображений**: Алгоритмы поиска изображений по артикулам и другим идентификаторам основаны на подходах из ExcelWithImages
- **Согласованная обработка ошибок и логирование**: Все подпроекты используют общие механизмы обработки исключений и ведения журналов

### Используемые технологии

- **Python**: Основной язык программирования
- **Pandas**: Для работы с табличными данными
- **Pillow/OpenCV**: Для обработки изображений
- **FPDF**: Для создания PDF-документов
- **Streamlit**: Для создания пользовательских интерфейсов (в некоторых подпроектах)

## Установка и запуск

### Общие требования

- Python 3.8 или выше
- Зависимости из requirements.txt

### Установка

1. Клонируйте репозиторий:
   ```
   git clone <url-репозитория>
   cd ImageProcessor
   ```

2. Создайте виртуальное окружение и активируйте его:
   ```
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # Linux/Mac
   source .venv/bin/activate
   ```

3. Установите общие зависимости:
   ```
   pip install -r requirements.txt
   ```

4. Установите зависимости для конкретного подпроекта (при необходимости):
   ```
   pip install -r ExcelToPDF/requirements.txt
   pip install -r ExcelWithImages/requirements.txt
   pip install -r PhotoProcessor/requirements.txt
   ```

### Запуск проекта

#### Единый запуск через общий интерфейс
```
python start.py
```
Этот способ запустит общий интерфейс выбора подпроекта, унифицированный на основе интерфейса ExcelWithImages.

#### Запуск отдельных подпроектов

##### PhotoProcessor
```
python PhotoProcessor/start.py
```

##### ExcelToPDF
```
python ExcelToPDF/start.py
```

##### ExcelWithImages
```
python ExcelWithImages/start.py
```

## Текущее состояние проекта

Проект находится в активной разработке. Основные компоненты функциональны, но продолжают совершенствоваться и расширяться. Разработка ведется с сохранением единой архитектуры и принципов работы, заложенных в подпроекте ExcelWithImages.

### Планы развития

- Дальнейшая унификация интерфейсов всех подпроектов
- Расширение функциональности поиска и обработки изображений
- Улучшение производительности при работе с большими объемами данных
- Добавление новых форматов экспорта и импорта данных

### Последние изменения

- Унификация структуры проекта на основе архитектуры ExcelWithImages
- Реализация двухколоночной таблицы в ExcelToPDF для улучшения читаемости PDF-документов
- Оптимизация алгоритмов обработки изображений в PhotoProcessor
- Улучшение поддержки различных форматов данных в ExcelWithImages
- Создание единого интерфейса запуска подпроектов на базе шаблонов ExcelWithImages

### Планируемые улучшения

- Дальнейшая интеграция всех подпроектов в единый интерфейс на базе ExcelWithImages
- Унификация механизмов поиска и обработки изображений по шаблону ExcelWithImages
- Расширение возможностей обработки изображений с сохранением единой архитектуры
- Улучшение производительности при работе с большими объемами данных
- Добавление новых форматов экспорта данных с использованием общих компонентов

## Лицензия

Проект распространяется под лицензией MIT. Подробности см. в файле LICENSE.

## Контакты и поддержка

По вопросам работы с проектом и предложениям по улучшению обращайтесь через систему Issues на GitHub или по электронной почте.
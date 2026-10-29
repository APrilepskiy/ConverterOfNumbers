# Конвертер Numbers → Excel

Простое веб-приложение на Flask для конвертации файлов Apple Numbers (`.numbers`)
в документы Microsoft Excel (`.xlsx`). Интерфейс полностью на русском языке.

## Возможности
- Загрузка файла через диалог выбора или перетаскивание (drag & drop)
- Сохранение структуры листов и данных таблиц
- Перенос базового форматирования (жирный шрифт, цвет фона, выравнивание, ширина колонок)
- Понятные сообщения об ошибках на русском языке

## Установка и запуск

### ALT Linux (Sisyphus / p10)
```bash
# Системные зависимости (от root, через su или sudo)
apt-get install python3-module-pip python3-dev zlib-devel tar

# Виртуальное окружение и зависимости Python (от обычного пользователя)
cd numbers-converter
python3 -m venv venv
venv/bin/pip install -r requirements.txt

# Запуск
venv/bin/python app.py
```

### Другие системы
```bash
pip install -r requirements.txt
python app.py
```

Откройте в браузере http://localhost:5000

## Как это работает
1. Файл `.numbers` загружается через форму на странице.
2. Сервер читает его с помощью библиотеки `numbers-parser`.
3. Данные таблиц записываются в `.xlsx` с помощью `openpyxl`.
4. Готовый файл автоматически скачивается в браузере.

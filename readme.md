# Основная библиотека из реестра Т-Банка
pip install t-tech-investments --index-url https://opensource.tbank.ru/api/v4/projects/238/packages/pypi/simple

# Остальные зависимости
pip install openpyxl customtkinter

# Запуск с выбором режима
python main.py

# Или сразу GUI
python main.py --gui

# Или сразу консоль
python main.py --console

# Можно задать токен через переменную окружения
export TINKOFF_TOKEN="t."
python main.py

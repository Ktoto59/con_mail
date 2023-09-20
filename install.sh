#!/bin/bash

# Обновляем пакеты
sudo apt-get update

# Устанавливаем Python3 и pip3
sudo apt-get install -y python3 python3-pip

# Устанавливаем virtualenv
sudo pip3 install virtualenv

# Создаем виртуальное окружение
python3 -m venv venv

# Активируем виртуальное окружение
source venv/bin/activate

# Устанавливаем зависимости в виртуальное окружение
pip install openpyxl psycopg2

# Деактивируем виртуальное окружение
deactivate

# Создаем исполняемый файл для вашего скрипта
echo '#!/bin/bash
source venv/bin/activate
python3 con_emls.py
deactivate' > run.sh
chmod +x run.sh

# Добавляем задачу в CRON
(crontab -l ; echo "0 */2 * * * cd $(pwd) && ./run.sh") | crontab -

# Все готово
echo "Установка и настройка завершены!"
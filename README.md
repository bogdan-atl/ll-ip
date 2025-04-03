Автоматическая актуализация ip xlsx

Установка:
-----------
git clone https://github.com/bogdan-atl/ll-ip.git

apt install python3.10-venv 

python3 -m venv venv

source venv/bin/activate

pip install -r requirements.txt

Запуск:
----------

python3 ip.py

Инсрукция
-----------

Автоматически найдет xlsx файл для отработки 

и сохранит результат в файл log(название листа).txt

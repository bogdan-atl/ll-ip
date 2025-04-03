Автоматическая актуализация ip xlsx

Установка:

apt install python3.10-venv 

python3 -m venv venv

source venv/bin/activate

pip install -r requirements.txt

Запуск:

Указываем путь до xlsx файла в ip.py  раздел input_file = "file_path.xlsx" 

python3 ip.py

вывод результата в строке log_file = "logMCK.txt"


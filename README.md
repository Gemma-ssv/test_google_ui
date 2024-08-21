#ОПИСАНИЕ
>Это тестовая работа. Собранная на коленке и быстро во время кофепауз на работе. ФАЙЛ test_google.py скрипты смены пароля, задание нового имени и фамилии пользователю google,
>создания файла эксель, сохранения в файл эксель данных о пользователе google.
## КАК запустить?
Заранее подготовьте среду разработки, установить Python версии не ниже 3.12.x

```bash
sudo apt install python3.12
```

>Создайте отдельную папку на ПК.

```bash
mkdir my_bot_folder
cd my_bot_folder
```

>Создайте виртуальное окружение.

```bash
python -m venv test_env
```

>Запустите виртуальное окружение.

```bash
source test_env/bin/activate
```

>Склонируйте репозиторий себе на ПК в соответствующую директорию.

```bash
git clone https://github.com/username/repository.git
```

>Внесите соответствующие изменения в конфигурационный файл `configs.env`.

>Установите все библиотеки из файла `requirements.txt`.

```bash
pip install -r requirements.txt
```

>Замените соответсвующие передаваемые аргументы в функциях
```python
# ВАРИАНТ ИСПОЛЬЗОВАНИЯ

# СОЗДАЁМ ФАЙЛ
file = create_excel_file('test.xlsx')

# ОБНОВЛЯЕМ ПАРОЛЬ
new_passw = change_email_password('test@gmail.com','test_password', 'new_test_password')

# ОБНОВЛЯЕМ ИМЯ И ФАМИЛИЯ
change_name_and_surname_to_email('test@gmail.com', new_passw,'ТЕСТ', 'ТЕСТОВ')

# ПОЛУЧАЕМ ОБНОВЛЕННЫЕ ДАННЫЕ
data_list = get_data_email('test@gmail.com', new_passw)

# СОХРАНЯЕМ ДАННЫЕ В ФАЙЛ
save_data_email(file, data_list)
```

>Откройте консоль и запустите файл `test_google.py`.

```bash
python test_google.py
```
>?????????
>PROFIT!!!

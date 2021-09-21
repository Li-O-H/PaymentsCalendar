import psycopg2
from tkinter import messagebox


# Параметры БД
host = "localhost"
port = 5432
database = "payments_calendar"

# Таблица записей и ее столбцы
records_table = "records"
r_id_column = "id"
r_departament_id_column = "dept_id"
r_position_id_column = "pos_id"
r_currency_column = "currency_id"
r_direction_column = "direction"
r_date_column = "date"
r_value_column = "value"
r_comment_column = "comment"
r_timestamp_column = "record_timestamp"

# Таблица статей и ее столбцы
positions_table = "positions"
p_id_column = "id"
p_code_column = "code"
p_name_column = "name"
p_status_column = "status"
p_active_status = 1

# Таблица пользователей и ее столбцы
users_table = "users"
u_id_column = "id"
u_name_column = "uname"

# Таблица подразделений и ее столбцы
depts_table = "depts"
d_id_column = "id"
d_full_name_column = "fullname"
d_short_name_column = "shortname"
d_responsibility_column = "is_responsible"

# Таблица, сопоставляющая пользователю подразделение, и ее столбцы
depts_users_table = "depts_users"
du_dept_id_column = "dept_id"
du_user_id_column = "user_id"

# Таблица валют и ее столбцы
currency_table = "currency"
c_id_column = "id"
c_full_name_column = "fullname"
c_short_name_column = "shortname"

default_time = "13:00:00"

# Словарь, позволяющий найти короткое имя подразделения по id (используется при формировании комментария при своде)
departments_short_names = {}
# Словарь, позволяющий найти id валюты по короткому имени
currencies_ids = {}
# Словарь, позволяющий найти id статьи по коду
positions_ids = {}
# Сохраненная из БД таблица с записями
records_saved_table = []
# Словарь, в котором каждой дате сопоставлены строки в сохраненной таблице
dates_rows = {}


# Объект, представляющий запись, которую нужно сохранить в БД
class DbWriteRecord:
    def __init__(self, position_code, date, direction, currency, value, comment_text=None):
        self.position_code = position_code
        self.date = date
        self.direction = direction
        self.currency_name = currency
        self.value = value
        self.comment_text = comment_text


# Объект, представляющий значение, полученное из БД
class DbReadRecord:
    def __init__(self, value, comment_text):
        self.value = value
        self.comment_text = comment_text


# Установка соединения с БД
def connect_to_db(user, password):
    try:
        connection = psycopg2.connect(
            host=host,
            port=port,
            dbname=database,
            user=user,
            password=password
        )
    except psycopg2.Error:
        messagebox.showerror("Ошибка", f"Не удалось подключиться к базе данных")
        return None
    return connection


# Сохранение записей в БД
def write_records(connection, department_id, records):
    with connection.cursor() as cursor:
        # Получаем отметку времени
        try:
            cursor.execute("select now()::timestamptz(2);")
        except psycopg2.Error:
            messagebox.showerror("Ошибка", f"Ошибка при работе с базой данных")
            return False
        current_timestamp = cursor.fetchall()[0][0]
        values = ""
        for record in records:
            # Находим id валюты, соответствующий короткому названию валюты из записи
            currency_id = get_currency_id(connection, record.currency_name)
            if currency_id is None:
                return False
            # Находим id статьи, соответствующий коду статьи из записи
            position_id = get_position_id(connection, record.position_code)
            if position_id is None:
                return False
            if record.comment_text is None:
                comment_text = "null"
            else:
                comment_text = f"'{record.comment_text}'"
            date = f"{str(record.date)} {default_time}"
            # Для каждой записи формируем будущую запись в БД
            values = f"{values}({department_id}, {position_id}, {currency_id}, {record.direction}, " \
                     f"'{date}', '{record.value}', {comment_text}, '{current_timestamp}'), "
        values = values[:-2]
        # Записываем все записи в БД
        try:
            cursor.execute(f"insert into \"{records_table}\" (\"{r_departament_id_column}\", "
                           f"\"{r_position_id_column}\", \"{r_currency_column}\", \"{r_direction_column}\", "
                           f"\"{r_date_column}\", \"{r_value_column}\", \"{r_comment_column}\", "
                           f"\"{r_timestamp_column}\") values {values};")
        except psycopg2.Error:
            messagebox.showerror("Ошибка",
                                 f"Таблица {records_table} в базе данных повреждена, отсутствует или недоступна")
            return False
        connection.commit()
        return True


# Поиск записи в БД
def get_record(connection, department_id, position_code, date, direction, currency_name):
    # Находим id валюты, соответствующий currency_name
    currency_id = get_currency_id(connection, currency_name)
    if currency_id is None:
        return None
    # Находим id статьи, соответствующий position_code
    position_id = get_position_id(connection, position_code)
    if position_id is None:
        return None
    if dates_rows.get(date.toordinal()) is not None:
        for row in dates_rows[date.toordinal()]:
            if records_saved_table[row][1] == department_id and records_saved_table[row][3] == currency_id and \
                    records_saved_table[row][4] == direction and records_saved_table[row][2] == position_id:
                return DbReadRecord(records_saved_table[row][6], records_saved_table[row][7])
    return DbReadRecord(None, None)


# Обновление считанных записей (нужно выполнять перед поиском записей в БД)
def refresh_records(connection):
    with connection.cursor() as cursor:
        try:
            cursor.execute(f"select * from \"{records_table}\" order by \"{r_timestamp_column}\" desc;")
        except psycopg2.Error:
            messagebox.showerror("Ошибка",
                                 f"Таблица {records_table} в базе данных повреждена, отсутствует или недоступна")
            return False
        records_saved_table.clear()
        records_saved_table.extend(cursor.fetchall())
        dates_rows.clear()
        for i in range(len(records_saved_table)):
            if dates_rows.get(records_saved_table[i][5].toordinal()) is None:
                dates_rows[records_saved_table[i][5].toordinal()] = [i]
            else:
                dates_rows[records_saved_table[i][5].toordinal()].append(i)
        return True


# Получение id валюты по ее короткому имени
def get_currency_id(connection, currency_short_name):
    if len(currencies_ids) == 0:
        with connection.cursor() as cursor:
            try:
                cursor.execute(f"select \"{c_short_name_column}\", \"{c_id_column}\" from \"{currency_table}\";")
            except psycopg2.Error:
                messagebox.showerror("Ошибка",
                                     f"Таблица {currency_table} в базе данных повреждена, отсутствует или недоступна")
                return None
            db_respond = cursor.fetchall()
            if len(db_respond) == 0:
                messagebox.showerror("Ошибка", f"В таблице {currency_table} нет записей о валютах")
            for cur in db_respond:
                currencies_ids[cur[0]] = cur[1]
    if currencies_ids.get(currency_short_name) is None:
        messagebox.showerror("Ошибка", f"Валюта {currency_short_name} неизвестна")
        with connection.cursor() as cursor:
            try:
                cursor.execute(f"select \"{c_full_name_column}\", \"{c_short_name_column}\" from \"{currency_table}\";")
                currencies = ""
                for currency_name in cursor.fetchall():
                    currencies = f"{currencies}{currency_name[0]} - {currency_name[1]}\n"
                messagebox.showinfo("Подсказка", f"Доступные валюты: \n{currencies}")
            except psycopg2.Error:
                messagebox.showerror("Ошибка",
                                     f"Таблица {currency_table} в базе данных повреждена, отсутствует или недоступна")
            return None
    return currencies_ids.get(currency_short_name)


# Получение id статьи по ее коду
def get_position_id(connection, position_code):
    if len(positions_ids) == 0:
        with connection.cursor() as cursor:
            try:
                cursor.execute(f"select \"{p_code_column}\", \"{p_id_column}\" from \"{positions_table}\";")
            except psycopg2.Error:
                messagebox.showerror("Ошибка",
                                     f"Таблица {positions_table} в базе данных повреждена, отсутствует или недоступна")
                return None
            db_respond = cursor.fetchall()
            if len(db_respond) == 0:
                messagebox.showerror("Ошибка", f"В таблице {positions_table} нет записей о статьях")
            for pos in db_respond:
                positions_ids[pos[0]] = pos[1]
    if positions_ids.get(position_code) is None:
        messagebox.showerror("Ошибка", f"В таблице {positions_table} нет записей о статье {position_code}")
    return positions_ids.get(position_code)


# Определить подразделение пользователя
def get_user_department(connection, user):
    with connection.cursor() as cursor:
        try:
            cursor.execute(
                f"select \"{u_id_column}\" from \"{users_table}\" where \"{u_name_column}\" = '{user}';")
        except psycopg2.Error:
            messagebox.showerror("Ошибка",
                                 f"Таблица {users_table} в базе данных повреждена, отсутствует или недоступна")
            return None
        db_respond = cursor.fetchall()
        if len(db_respond) == 1:
            user_id = db_respond[0][0]
        else:
            messagebox.showerror("Ошибка", f"Нет информации о пользователе {user}")
            return None
        try:
            cursor.execute(f"select \"{du_dept_id_column}\" "
                           f"from \"{depts_users_table}\" where \"{du_user_id_column}\" = '{user_id}';")
        except psycopg2.Error:
            messagebox.showerror("Ошибка",
                                 f"Таблица {depts_users_table} в базе данных повреждена, отсутствует или недоступна")
            return None
        db_respond = cursor.fetchall()
        if len(db_respond) == 1:
            dept_id = db_respond[0][0]
        else:
            messagebox.showerror("Ошибка", f"Неизвестно подразделение пользователя {user}")
            return None
        return dept_id


# Получить список активных статей
def get_active_positions(connection):
    with connection.cursor() as cursor:
        try:
            cursor.execute(f"select \"{p_code_column}\" from \"{positions_table}\" "
                           f"where \"{p_status_column}\" = {p_active_status};")
        except psycopg2.Error:
            messagebox.showerror("Ошибка",
                                 f"Таблица {positions_table} в базе данных повреждена, отсутствует или недоступна")
            return None
        codes = []
        for record in cursor.fetchall():
            codes.append(record[0])
        return codes


# Получить список подразделений
def get_all_departments(connection):
    with connection.cursor() as cursor:
        try:
            cursor.execute(f"select \"{d_id_column}\" from \"{depts_table}\";")
        except psycopg2.Error:
            messagebox.showerror("Ошибка",
                                 f"Таблица {depts_table} в базе данных повреждена, отсутствует или недоступна")
            return None
        departments = []
        for record in cursor.fetchall():
            departments.append(record[0])
        return departments


# Получить короткое имя подразделения (для формирования комментария при своде)
def get_department_name(connection, department_id):
    if len(departments_short_names) == 0:
        with connection.cursor() as cursor:
            try:
                cursor.execute(f"select \"{d_id_column}\", \"{d_short_name_column}\" from \"{depts_table}\";")
            except psycopg2.Error:
                messagebox.showerror("Ошибка",
                                     f"Таблица {depts_table} в базе данных повреждена, отсутствует или недоступна")
                return None
            db_respond = cursor.fetchall()
            if len(db_respond) == 0:
                messagebox.showerror("Ошибка", f"В таблице {depts_table} нет записей о подразделениях")
            for dept in db_respond:
                departments_short_names[dept[0]] = dept[1]
    if departments_short_names.get(department_id) is None:
        messagebox.showerror("Ошибка", f"В таблице {depts_table} нет записей о подразделении с id {department_id}")
    return departments_short_names.get(department_id)


# Имеет ли право пользователь на свод/выгрузку
def is_user_responsible(connection, user):
    department_id = get_user_department(connection, user)
    with connection.cursor() as cursor:
        try:
            cursor.execute(f"select \"{d_responsibility_column}\" from \"{depts_table}\" "
                           f"where \"{d_id_column}\" = {department_id};")
        except psycopg2.Error:
            messagebox.showerror("Ошибка",
                                 f"Таблица {depts_table} в базе данных повреждена, отсутствует или недоступна")
            return None
        db_respond = cursor.fetchall()
        if len(db_respond) == 1:
            return db_respond[0][0]
        return False


# Получить все записи из records (для выгрузки)
def get_all_records(connection):
    with connection.cursor() as cursor:
        try:
            cursor.execute(
                f"select r.\"{r_id_column}\", d.\"{d_short_name_column}\", p.\"{p_code_column}\", "
                f"c.\"{c_short_name_column}\", r.\"{r_direction_column}\", r.\"{r_date_column}\", "
                f"r.\"{r_value_column}\", r.\"{r_comment_column}\", r.\"{r_timestamp_column}\" "
                f"from \"{records_table}\" as r join \"{depts_table}\" as d on "
                f"r.\"{r_departament_id_column}\" = d.\"{d_id_column}\" join \"{positions_table}\" as p on "
                f"r.\"{r_position_id_column}\" = p.\"{p_id_column}\" join \"{currency_table}\" as c on "
                f"r.\"{r_currency_column}\" = c.\"{c_id_column}\";")
        except psycopg2.Error:
            messagebox.showerror("Ошибка", f"Таблица {records_table}, {depts_table}, {positions_table} или "
                                           f"{currency_table} в базе данных повреждена, отсутствует или недоступна")
            return None
        return cursor.fetchall()

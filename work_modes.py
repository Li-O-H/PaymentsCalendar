from tkinter import messagebox

import db_worker
import excel_worker


# Первый режим - заполнение шаблона данными из БД
def mode1(user, password, input_file, output_file):
    if input_file == output_file:
        messagebox.showerror("Ошибка", "Выходной файл не должен совпадать со входным")
        return
    connection = db_worker.connect_to_db(user, password)
    if connection is None:
        return
    try:
        dept_id = db_worker.get_user_department(connection, user)
        if dept_id is None:
            return
        active_codes = db_worker.get_active_positions(connection)
        if active_codes is None:
            return
        if len(active_codes) == 0:
            messagebox.showinfo("Выполнение остановлено", "В базе данных нет информации об активных статьях")
            return
        if excel_worker.mode1_execute(input_file, connection, active_codes, dept_id, output_file):
            messagebox.showinfo("Выполнено", f"Значения из базы данных записаны в файл {output_file}")
    except Exception as e:
        messagebox.showerror("Непредвиденная ошибка", "Ошибка:\n" + str(e))
    finally:
        connection.close()


# Второй режим - сохранение ПК в БД
def mode2(user, password, input_file):
    connection = db_worker.connect_to_db(user, password)
    if connection is None:
        return
    try:
        dept_id = db_worker.get_user_department(connection, user)
        if dept_id is None:
            return
        active_codes = db_worker.get_active_positions(connection)
        if active_codes is None:
            return
        if len(active_codes) == 0:
            messagebox.showinfo("Выполнение остановлено", "В базе данных нет информации об активных статьях")
            return
        records = excel_worker.mode2_execute(input_file, active_codes)
        if records is None:
            return
        if len(records) == 0:
            messagebox.showinfo("Выполнение остановлено", "Нет значений для записи в базу данных")
            return
        if db_worker.write_records(connection, dept_id, records):
            messagebox.showinfo("Выполнено", f"Значения из файла {input_file} записаны в базу данных")
    except Exception as e:
        messagebox.showerror("Непредвиденная ошибка", "Ошибка:\n" + str(e))
    finally:
        connection.close()


# Третий режим - свод данных от подразделений
def mode3(user, password, input_file, output_file):
    if input_file == output_file:
        messagebox.showerror("Ошибка", "Выходной файл не должен совпадать со входным")
        return
    connection = db_worker.connect_to_db(user, password)
    if connection is None:
        return
    try:
        if not db_worker.is_user_responsible(connection, user):
            messagebox.showerror("Ошибка", "Пользователь не имеет права на свод")
            return
        departments = db_worker.get_all_departments(connection)
        if departments is None:
            return
        if len(departments) == 0:
            messagebox.showinfo("Выполнение остановлено", "В базе данных нет информации о подразделениях")
            return
        active_codes = db_worker.get_active_positions(connection)
        if active_codes is None:
            return
        if len(active_codes) == 0:
            messagebox.showinfo("Выполнение остановлено", "В базе данных нет информации об активных статьях")
            return
        if excel_worker.mode3_execute(input_file, connection, active_codes, departments, output_file):
            messagebox.showinfo("Выполнено", f"Значения из базы данных записаны в файл {output_file}")
    except Exception as e:
        messagebox.showerror("Непредвиденная ошибка", "Ошибка:\n" + str(e))
    finally:
        connection.close()


# Четвертый режим - выгрузка всех записей из БД в простую таблицу
def mode4(user, password, output_file):
    connection = db_worker.connect_to_db(user, password)
    if connection is None:
        return
    try:
        if not db_worker.is_user_responsible(connection, user):
            messagebox.showerror("Ошибка", "Пользователь не имеет права на выгрузку записей из базы данных")
            return
        records = db_worker.get_all_records(connection)
        if records is None:
            return
        if len(records) == 0:
            messagebox.showinfo("Выполнение остановлено", "В базе данных нет сохраненных записей")
            return
        if excel_worker.mode4_execute(records, output_file):
            messagebox.showinfo("Выполнено", f"Записи из базы данных выгружены в файл {output_file}")
    except Exception as e:
        messagebox.showerror("Непредвиденная ошибка", "Ошибка:\n" + str(e))
    finally:
        connection.close()

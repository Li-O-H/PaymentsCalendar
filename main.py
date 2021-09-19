import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import filedialog

import work_modes


# Создание вкладки 1 режима
def tab1_create(tab_control):
    tab1 = ttk.Frame(tab_control)
    tab_control.add(tab1, text="Заполнить пустой шаблон")
    tab1_info = Label(tab1, text="Шаблон платежного календаря заполнится сохраненными значениями из базы данных",
                      bg="white")
    tab1_info.pack(side=TOP, anchor=W, padx=10)

    tab1_var1 = tkinter.StringVar()
    tab1_var2 = tkinter.StringVar()
    tab1_var3 = tkinter.StringVar()
    tab1_var4 = tkinter.StringVar()

    def tab1_finish_button_check_activate(*args):
        if tab1_var1.get() != "" and tab1_var2.get() != "" and tab1_var3.get() != "" and tab1_var4.get() != "":
            tab1_finish_button['state'] = "normal"
        else:
            tab1_finish_button['state'] = "disabled"

    tab1_var1.trace("w", tab1_finish_button_check_activate)
    tab1_var2.trace("w", tab1_finish_button_check_activate)
    tab1_var3.trace("w", tab1_finish_button_check_activate)
    tab1_var4.trace("w", tab1_finish_button_check_activate)

    tab1_frame1 = Frame(tab1)
    tab1_text1_invite = Label(tab1_frame1, text="Выберите шаблон:")
    tab1_text1_invite.pack(side=LEFT, padx=10)

    def clicked_tab1_text1():
        file = filedialog.askopenfilename(filetype=[("Файлы Excel", "*.xlsx")], title="Выбор шаблона")
        tab1_var1.set(file)

    tab1_text1_button = Button(tab1_frame1, text="Выбрать файл", command=clicked_tab1_text1, bg="white")
    tab1_text1_button.pack(side=LEFT, padx=10)
    tab1_text1 = Label(tab1_frame1, textvariable=tab1_var1)
    tab1_text1.pack(side=LEFT, padx=10)
    tab1_frame1.pack(side=TOP, anchor=W, pady=5)

    tab1_frame2 = Frame(tab1)
    tab1_text2_invite = Label(tab1_frame2, text="Введите имя пользователя базы данных:")
    tab1_text2_invite.pack(side=LEFT, padx=10)
    tab1_text2_entry = Entry(tab1_frame2, width=30, textvariable=tab1_var2)
    tab1_text2_entry.pack(side=LEFT, padx=10)
    tab1_frame2.pack(side=TOP, anchor=W, pady=5)

    tab1_frame3 = Frame(tab1)
    tab1_text3_invite = Label(tab1_frame3, text="Введите пароль пользователя базы данных:")
    tab1_text3_invite.pack(side=LEFT, padx=10)
    tab1_text3_entry = Entry(tab1_frame3, width=30, show="*", textvariable=tab1_var3)
    tab1_text3_entry.pack(side=LEFT, padx=10)
    tab1_frame3.pack(side=TOP, anchor=W, pady=5)

    tab1_frame4 = Frame(tab1)
    tab1_text4_invite = Label(tab1_frame4, text="Выберите, куда сохранить заполненный шаблон:")
    tab1_text4_invite.pack(side=LEFT, padx=10)

    def clicked_tab1_text4():
        file = filedialog.asksaveasfilename(filetype=[("Файлы Excel", "*.xlsx")], title="Выбор места сохранения",
                                            defaultextension=".xlsx")
        tab1_var4.set(file)

    tab1_text4_button = Button(tab1_frame4, text="Выбрать файл", command=clicked_tab1_text4, bg="white")
    tab1_text4_button.pack(side=LEFT, padx=10)
    tab1_text4 = Label(tab1_frame4, textvariable=tab1_var4)
    tab1_text4.pack(side=LEFT, padx=10)
    tab1_frame4.pack(side=TOP, anchor=W, pady=5)

    def clicked_tab1_finish():
        work_modes.mode1(tab1_var2.get(), tab1_var3.get(), tab1_var1.get(), tab1_var4.get())
        tab1_var1.set("")
        tab1_text2_entry.delete(0, END)
        tab1_text3_entry.delete(0, END)
        tab1_var4.set("")

    tab1_finish_button = Button(tab1, text="Готово", command=clicked_tab1_finish, bg="white", font=("Arial", 12),
                                state="disabled")
    tab1_finish_button.pack(side=BOTTOM, anchor=W, padx=10, pady=15)

    def tab1_finish_button_invoke(text):
        tab1_finish_button.invoke()

    tab1_text2_entry.bind("<Return>", tab1_finish_button_invoke)
    tab1_text3_entry.bind("<Return>", tab1_finish_button_invoke)


# Создание вкладки 2 режима
def tab2_create(tab_control):
    tab2 = ttk.Frame(tab_control)
    tab_control.add(tab2, text="Сохранить заполненный ПК")
    tab2_info = Label(tab2, text="Заполненный платежный календарь сохранится в базу данных", bg="white")
    tab2_info.pack(side=TOP, anchor=W, padx=10)

    tab2_var1 = tkinter.StringVar()
    tab2_var2 = tkinter.StringVar()
    tab2_var3 = tkinter.StringVar()

    def tab2_finish_button_check_activate(*args):
        if tab2_var1.get() != "" and tab2_var2.get() != "" and tab2_var3.get() != "":
            tab2_finish_button['state'] = "normal"
        else:
            tab2_finish_button['state'] = "disabled"

    tab2_var1.trace("w", tab2_finish_button_check_activate)
    tab2_var2.trace("w", tab2_finish_button_check_activate)
    tab2_var3.trace("w", tab2_finish_button_check_activate)

    tab2_frame1 = Frame(tab2)
    tab2_text1_invite = Label(tab2_frame1, text="Выберите заполненный шаблон:")
    tab2_text1_invite.pack(side=LEFT, padx=10)

    def clicked_tab2_text1():
        file = filedialog.askopenfilename(filetype=[("Файлы Excel", "*.xlsx")], title="Выбор заполненного шаблона")
        tab2_var1.set(file)

    tab2_text1_button = Button(tab2_frame1, text="Выбрать файл", command=clicked_tab2_text1, bg="white")
    tab2_text1_button.pack(side=LEFT, padx=10)
    tab2_text1 = Label(tab2_frame1, textvariable=tab2_var1)
    tab2_text1.pack(side=LEFT, padx=10)
    tab2_frame1.pack(side=TOP, anchor=W, pady=5)

    tab2_frame2 = Frame(tab2)
    tab2_text2_invite = Label(tab2_frame2, text="Введите имя пользователя базы данных:")
    tab2_text2_invite.pack(side=LEFT, padx=10)
    tab2_text2_entry = Entry(tab2_frame2, width=30, textvariable=tab2_var2)
    tab2_text2_entry.pack(side=LEFT, padx=10)
    tab2_frame2.pack(side=TOP, anchor=W, pady=5)

    tab2_frame3 = Frame(tab2)
    tab2_text3_invite = Label(tab2_frame3, text="Введите пароль пользователя базы данных:")
    tab2_text3_invite.pack(side=LEFT, padx=10)
    tab2_text3_entry = Entry(tab2_frame3, width=30, show="*", textvariable=tab2_var3)
    tab2_text3_entry.pack(side=LEFT, padx=10)
    tab2_frame3.pack(side=TOP, anchor=W, pady=5)

    def clicked_tab2_finish():
        work_modes.mode2(tab2_var2.get(), tab2_var3.get(), tab2_var1.get())
        tab2_var1.set("")
        tab2_text2_entry.delete(0, END)
        tab2_text3_entry.delete(0, END)

    tab2_finish_button = Button(tab2, text="Готово", command=clicked_tab2_finish, bg="white", font=("Arial", 12),
                                state="disabled")
    tab2_finish_button.pack(side=BOTTOM, anchor=W, padx=10, pady=15)

    def tab2_finish_button_invoke(text):
        tab2_finish_button.invoke()

    tab2_text2_entry.bind("<Return>", tab2_finish_button_invoke)
    tab2_text3_entry.bind("<Return>", tab2_finish_button_invoke)


# Создание вкладки 3 режима
def tab3_create(tab_control):
    tab3 = ttk.Frame(tab_control)
    tab_control.add(tab3, text="Свод данных от подразделений")
    tab3_info = Label(tab3, text="Сведет (просуммирует) данные от всех подразделений в одну таблицу", bg="white")
    tab3_info.pack(side=TOP, anchor=W, padx=10)

    tab3_var1 = tkinter.StringVar()
    tab3_var2 = tkinter.StringVar()
    tab3_var3 = tkinter.StringVar()
    tab3_var4 = tkinter.StringVar()

    def tab3_finish_button_check_activate(*args):
        if tab3_var1.get() != "" and tab3_var2.get() != "" and tab3_var3.get() != "" and tab3_var4.get() != "":
            tab3_finish_button['state'] = "normal"
        else:
            tab3_finish_button['state'] = "disabled"

    tab3_var1.trace("w", tab3_finish_button_check_activate)
    tab3_var2.trace("w", tab3_finish_button_check_activate)
    tab3_var3.trace("w", tab3_finish_button_check_activate)
    tab3_var4.trace("w", tab3_finish_button_check_activate)

    tab3_frame1 = Frame(tab3)
    tab3_text1_invite = Label(tab3_frame1, text="Выберите шаблон:")
    tab3_text1_invite.pack(side=LEFT, padx=10)

    def clicked_tab3_text1():
        file = filedialog.askopenfilename(filetype=[("Файлы Excel", "*.xlsx")], title="Выбор шаблона")
        tab3_var1.set(file)

    tab3_text1_button = Button(tab3_frame1, text="Выбрать файл", command=clicked_tab3_text1, bg="white")
    tab3_text1_button.pack(side=LEFT, padx=10)
    tab3_text1 = Label(tab3_frame1, textvariable=tab3_var1)
    tab3_text1.pack(side=LEFT, padx=10)
    tab3_frame1.pack(side=TOP, anchor=W, pady=5)

    tab3_frame2 = Frame(tab3)
    tab3_text2_invite = Label(tab3_frame2, text="Введите имя пользователя базы данных:")
    tab3_text2_invite.pack(side=LEFT, padx=10)
    tab3_text2_entry = Entry(tab3_frame2, width=30, textvariable=tab3_var2)
    tab3_text2_entry.pack(side=LEFT, padx=10)
    tab3_frame2.pack(side=TOP, anchor=W, pady=5)

    tab3_frame3 = Frame(tab3)
    tab3_text3_invite = Label(tab3_frame3, text="Введите пароль пользователя базы данных:")
    tab3_text3_invite.pack(side=LEFT, padx=10)
    tab3_text3_entry = Entry(tab3_frame3, width=30, show="*", textvariable=tab3_var3)
    tab3_text3_entry.pack(side=LEFT, padx=10)
    tab3_frame3.pack(side=TOP, anchor=W, pady=5)

    tab3_frame4 = Frame(tab3)
    tab3_text4_invite = Label(tab3_frame4, text="Выберите, куда сохранить заполненный шаблон:")
    tab3_text4_invite.pack(side=LEFT, padx=10)

    def clicked_tab3_text4():
        file = filedialog.asksaveasfilename(filetype=[("Файлы Excel", "*.xlsx")], title="Выбор места сохранения",
                                            defaultextension=".xlsx")
        tab3_var4.set(file)

    tab3_text4_button = Button(tab3_frame4, text="Выбрать файл", command=clicked_tab3_text4, bg="white")
    tab3_text4_button.pack(side=LEFT, padx=10)
    tab3_text4 = Label(tab3_frame4, textvariable=tab3_var4)
    tab3_text4.pack(side=LEFT, padx=10)
    tab3_frame4.pack(side=TOP, anchor=W, pady=5)

    def clicked_tab3_finish():
        work_modes.mode3(tab3_var2.get(), tab3_var3.get(), tab3_var1.get(), tab3_var4.get())
        tab3_var1.set("")
        tab3_text2_entry.delete(0, END)
        tab3_text3_entry.delete(0, END)
        tab3_var4.set("")

    tab3_finish_button = Button(tab3, text="Готово", command=clicked_tab3_finish, bg="white", font=("Arial", 12),
                                state="disabled")
    tab3_finish_button.pack(side=BOTTOM, anchor=W, padx=10, pady=15)

    def tab3_finish_button_invoke(text):
        tab3_finish_button.invoke()

    tab3_text2_entry.bind("<Return>", tab3_finish_button_invoke)
    tab3_text3_entry.bind("<Return>", tab3_finish_button_invoke)


# Создание вкладки 4 режима
def tab4_create(tab_control):
    tab4 = ttk.Frame(tab_control)
    tab_control.add(tab4, text="Выгрузка записей из БД")

    tab4_info = Label(tab4, text="Все записи из базы данных будут выгружены в указанный файл", bg="white")
    tab4_info.pack(side=TOP, anchor=W, padx=10)

    tab4_var1 = tkinter.StringVar()
    tab4_var2 = tkinter.StringVar()
    tab4_var3 = tkinter.StringVar()

    def tab4_finish_button_check_activate(*args):
        if tab4_var1.get() != "" and tab4_var2.get() != "" and tab4_var3.get() != "":
            tab4_finish_button['state'] = "normal"
        else:
            tab4_finish_button['state'] = "disabled"

    tab4_var1.trace("w", tab4_finish_button_check_activate)
    tab4_var2.trace("w", tab4_finish_button_check_activate)
    tab4_var3.trace("w", tab4_finish_button_check_activate)

    tab4_frame1 = Frame(tab4)
    tab4_text1_invite = Label(tab4_frame1, text="Выберите, куда выгрузить файлы:")
    tab4_text1_invite.pack(side=LEFT, padx=10)

    def clicked_tab4_text1():
        file = filedialog.asksaveasfilename(filetype=[("Файлы Excel", "*.xlsx")], title="Выбор места сохранения",
                                            defaultextension=".xlsx")
        tab4_var1.set(file)

    tab4_text1_button = Button(tab4_frame1, text="Выбрать файл", command=clicked_tab4_text1, bg="white")
    tab4_text1_button.pack(side=LEFT, padx=10)
    tab4_text1 = Label(tab4_frame1, textvariable=tab4_var1)
    tab4_text1.pack(side=LEFT, padx=10)
    tab4_frame1.pack(side=TOP, anchor=W, pady=5)

    tab4_frame2 = Frame(tab4)
    tab4_text2_invite = Label(tab4_frame2, text="Введите имя пользователя базы данных:")
    tab4_text2_invite.pack(side=LEFT, padx=10)
    tab4_text2_entry = Entry(tab4_frame2, width=30, textvariable=tab4_var2)
    tab4_text2_entry.pack(side=LEFT, padx=10)
    tab4_frame2.pack(side=TOP, anchor=W, pady=5)

    tab4_frame3 = Frame(tab4)
    tab4_text3_invite = Label(tab4_frame3, text="Введите пароль пользователя базы данных:")
    tab4_text3_invite.pack(side=LEFT, padx=10)
    tab4_text3_entry = Entry(tab4_frame3, width=30, show="*", textvariable=tab4_var3)
    tab4_text3_entry.pack(side=LEFT, padx=10)
    tab4_frame3.pack(side=TOP, anchor=W, pady=5)

    def clicked_tab4_finish():
        work_modes.mode4(tab4_var2.get(), tab4_var3.get(), tab4_var1.get())
        tab4_var1.set("")
        tab4_text2_entry.delete(0, END)
        tab4_text3_entry.delete(0, END)

    tab4_finish_button = Button(tab4, text="Готово", command=clicked_tab4_finish, bg="white", font=("Arial", 12),
                                state="disabled")
    tab4_finish_button.pack(side=BOTTOM, anchor=W, padx=10, pady=15)

    def tab4_finish_button_invoke(text):
        tab4_finish_button.invoke()

    tab4_text2_entry.bind("<Return>", tab4_finish_button_invoke)
    tab4_text3_entry.bind("<Return>", tab4_finish_button_invoke)


if __name__ == "__main__":
    # Создание главного окна
    main_window = Tk()
    main_window.title("Платежный календарь")
    main_window.geometry("800x250")
    main_window.wm_minsize(width=660, height=250)
    tab_control = ttk.Notebook(main_window)
    tab1_create(tab_control)
    tab2_create(tab_control)
    tab3_create(tab_control)
    tab4_create(tab_control)
    tab_control.pack(expand=1, fill='both')
    main_window.mainloop()

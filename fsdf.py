import sqlite3, xlsxwriter, sys
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from PIL import Image
import pandas as pd
from tkinter.messagebox import showerror, showinfo
import os

abonent_name = ["№", "№ адреса", "Номер домашнего телефона", "№ служебного телефона"]
adres_name = ["№", "№ типа населенного пункта", "Населенный пункт", "Улица", "Номер дома"]
nas_pukt_name = ["№", "Наименование", "№ типа", "№ улицы"]
slujeb_tel_name = ["№", "№ предпреятия", "отдел", "Номер телефона"]

class AboutProgramWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("О программе")

        # Информация о программе
        program_name_label = tk.Label(self, text="Название программы: Городской телефонный справочник")
        program_version_label = tk.Label(self, text="Версия: 1.0")
        developer_label = tk.Label(self, text="Разработчик: Соловьев Максим, 2023")

        # Рамка для назначения программного средстваsssss
        purpose_frame = tk.LabelFrame(self, text="")  # Удалил текст "Назначение программного средства"
        purpose_text = (
            "Данное программное средство «Городской телефонный справочник» "
            "разрабатывается с целью автоматизации процесса введения отчетности"
        )
        purpose_label = tk.Label(purpose_frame, text=purpose_text, anchor="w", wraplength=300)

        # Размещение компонентов
        program_name_label.pack(pady=5)
        program_version_label.pack(pady=5)
        developer_label.pack(pady=5)

        purpose_frame.pack(pady=10, padx=10, ipadx=5, ipady=5)  # Добавлены ipadx и ipadytyryr
        purpose_label.pack(pady=5)

        # Кнопка "ОК" для закрытия окна
        ok_button = tk.Button(self, text="ОК", command=self.destroy)
        ok_button.pack(pady=10)

class WindowMain(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Городской телефонный справочник')
        self.wm_iconbitmap()
        self.iconphoto(True, tk.PhotoImage(file="image1\\icon.png"))
        self.last_headers = None

        # Создание фрейма для отображения таблицы
        self.table_frame = ctk.CTkFrame(self, width=700, height=400)
        self.table_frame.grid(row=0, column=0, padx=5, pady=5)

        # Загрузка фона
        bg = ctk.CTkImage(Image.open("image1\\fon.jpg"), size=(700, 400))
        lbl = ctk.CTkLabel(self.table_frame, image=bg, text='Городской телефонный справочник', font=("Calibri", 30))
        lbl.place(relwidth=1, relheight=1)

        # Создание меню
        self.menu_bar = tk.Menu(self, background='#555', foreground='white')

        # Меню "Файл"
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Выход", command=self.quit)
        self.menu_bar.add_cascade(label="Файл", menu=file_menu)

        # Меню "Таблицы"
        references_menu = tk.Menu(self.menu_bar, tearoff=0)
        references_menu.add_command(label="Абонет",
                                    command=lambda: self.show_table("SELECT * FROM abonent", abonent_name))
        references_menu.add_command(label="Адрес",
                                    command=lambda: self.show_table("SELECT * FROM adres", adres_name))
        references_menu.add_command(label="Cлужебные телефона",
                                    command=lambda: self.show_table("SELECT * FROM slyjeb_telephon", slujeb_tel_name))
        references_menu.add_command(label="Населённые пункты",
                                    command=lambda: self.show_table("SELECT * FROM id_nas_pynkt", nas_pukt_name))
        self.menu_bar.add_cascade(label="Таблицы", menu=references_menu)

        # Меню "Отчёты"
        reports_menu = tk.Menu(self.menu_bar, tearoff=0)
        reports_menu.add_command(label="Создать Отчёт", command=self.to_xlsx)
        self.menu_bar.add_cascade(label="Отчёты", menu=reports_menu)

        # Меню "Сервис"
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="Руководство пользователя", command=self.open_rykov)
        help_menu.add_command(label="O программе",command=self.open_about_window)
        self.menu_bar.add_cascade(label="Сервис", menu=help_menu)

        # Настройка цветов меню
        file_menu.configure(bg='#555', fg='white')
        references_menu.configure(bg='#555', fg='white')
        reports_menu.configure(bg='#555', fg='white')
        help_menu.configure(bg='#555', fg='white')

        # Установка меню в главное окно
        self.config(menu=self.menu_bar)

        btn_width = 150
        pad = 5

        # Создание кнопок и виджетов для поиска и редактирования данных
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=0, column=1)
        ctk.CTkButton(btn_frame, text="добавить", width=btn_width, command=self.add).pack(pady=pad)
        ctk.CTkButton(btn_frame, text="удалить", width=btn_width, command=self.delete).pack(pady=pad)
        ctk.CTkButton(btn_frame, text="изменить", width=btn_width, command=self.change).pack(pady=pad)

        search_frame = ctk.CTkFrame(self)
        search_frame.grid(row=1, column=0, pady=pad)
        self.search_entry = ctk.CTkEntry(search_frame, width=300)
        self.search_entry.grid(row=0, column=0, padx=pad)
        ctk.CTkButton(search_frame, text="Поиск", width=20, command=self.search).grid(row=0, column=1, padx=pad)
        ctk.CTkButton(search_frame, text="Искать далее", width=20, command=self.search_next).grid(row=0, column=2,
                                                                                                  padx=pad)
        ctk.CTkButton(search_frame, text="Сброс", width=20, command=self.reset_search).grid(row=0, column=3, padx=pad)

    def search_in_table(self, table, search_terms, start_item=None):
        table.selection_remove(table.selection())  # Сброс предыдущего выделения

        items = table.get_children('')
        start_index = items.index(start_item) + 1 if start_item else 0

        for item in items[start_index:]:
            values = table.item(item, 'values')
            for term in search_terms:
                if any(term.lower() in str(value).lower() for value in values):
                    table.selection_add(item)
                    table.focus(item)
                    table.see(item)
                    return item  # Возвращаем найденный элемент

    def reset_search(self):
        if self.last_headers:
            self.table.selection_remove(self.table.selection())
        self.search_entry.delete(0, 'end')

    def open_about_window(self):
        about_window = AboutProgramWindow(self)
        about_window.geometry("600x250")  # Установите размер окна по вашему усмотрению
        about_window.focus_set()
        about_window.grab_set()
        self.wait_window(about_window)

    def open_rykov(self):
        os.system(r"C:\Users\gysi-\PycharmProjects\pythonProject4\HTML/lf.html")

    def search(self):
        if self.last_headers:
            self.current_item = self.search_in_table(self.table, self.search_entry.get().split(','))

    def search_next(self):
        if self.last_headers:
            if self.current_item:
                self.current_item = self.search_in_table(self.table, self.search_entry.get().split(','),
                                                         start_item=self.current_item)

    def to_xlsx(self):
        if self.last_headers == abonent_name:
            sql_query = "SELECT * FROM abonent"
            table_name = "abonent"
        elif self.last_headers == adres_name:
            sql_query = "SELECT * FROM adres"
            table_name = "adres"
        elif self.last_headers == nas_pukt_name:
            sql_query = "SELECT * FROM id_nas_pynkt"
            table_name = "id_nas_pynkt"
        elif self.last_headers == slujeb_tel_name:
            sql_query = "SELECT * FROM slyjeb_telephon"
            table_name = "slyjeb_telephon"
        else:
            return

        dir = sys.path[0] + "\\export"
        os.makedirs(dir, exist_ok=True)
        path = dir + f"\\{table_name}.xlsx"

        # Подключение к базе данных SQLite
        conn = sqlite3.connect("gorod.db")
        cursor = conn.cursor()
        # Получите данные из базы данных
        cursor.execute(sql_query)
        data = cursor.fetchall()
        # Создайте DataFrame из данных
        df = pd.DataFrame(data, columns=self.last_headers)
        # Создайте объект writer для записи данных в Excel
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        # Запишите DataFrame в файл Excel
        df.to_excel(writer, 'Лист 1', index=False)
        # Сохраните результат
        writer.close()

        showinfo(title="Успешно", message=f"Данные экспортированы в {path}")

    def add(self):
        if self.last_headers == abonent_name:
            WindowAbonent("add")
        elif self.last_headers == nas_pukt_name:
            WindowNasPynkt("add")
        elif self.last_headers == slujeb_tel_name:
            WindowSlujeTel("add")
        elif self.last_headers == adres_name:
            WindowAdres("add")
        else:
            return

        self.withdraw()

    def delete(self):
        if self.last_headers:
            select_item = self.table.selection()
            if select_item:
                item_data = self.table.item(select_item[0])["values"]
            else:
                showerror(title="Ошибка", message="He выбранна запись")
                return
        else:
            return

        if self.last_headers == abonent_name:
            WindowAbonent("delete", item_data)
        elif self.last_headers == nas_pukt_name:
            WindowNasPynkt("delete", item_data)
        elif self.last_headers == slujeb_tel_name:
            WindowSlujeTel("delete", item_data)
        elif self.last_headers == adres_name:
            WindowAdres("delete", item_data)
        else:
            return

        self.withdraw()

    def change(self):
        if self.last_headers:
            select_item = self.table.selection()
            if select_item:
                item_data = self.table.item(select_item[0])["values"]
            else:
                showerror(title="Ошибка", message="He выбранна запись")
                return
        else:
            return

        if self.last_headers == abonent_name:
            WindowAbonent("change", item_data)
        elif self.last_headers == nas_pukt_name:
            WindowNasPynkt("change", item_data)
        elif self.last_headers == slujeb_tel_name:
            WindowSlujeTel("change", item_data)
        elif self.last_headers == adres_name:
            WindowAdres("change", item_data)
        else:
            return

        self.withdraw()

    def show_table(self, sql_query, headers=None):
        # Очистка фрейма перед отображением новых данных
        for widget in self.table_frame.winfo_children(): widget.destroy()

        # Подключение к базе данных SQLite
        conn = sqlite3.connect("gorod.db")
        cursor = conn.cursor()

        # Выполнение SQL-запроса
        cursor.execute(sql_query)
        self.last_sql_query = sql_query

        # Получение заголовков таблицы и данных
        if headers == None:  # если заголовки не были переданы используем те что в БД
            table_headers = [description[0] for description in cursor.description]
        else:  # иначе используем те что передали
            table_headers = headers
            self.last_headers = headers
        table_data = cursor.fetchall()

        # Закрытие соединения с базой данных
        conn.close()

        canvas = ctk.CTkCanvas(self.table_frame, width=865, height=480)
        canvas.pack(fill="both", expand=True)

        x_scrollbar = ttk.Scrollbar(self.table_frame, orient="horizontal", command=canvas.xview)
        x_scrollbar.pack(side="bottom", fill="x")

        canvas.configure(xscrollcommand=x_scrollbar.set)

        self.table = ttk.Treeview(self.table_frame, columns=table_headers, show="headings", height=23)
        for header in table_headers:
            self.table.heading(header, text=header)
            self.table.column(header,
                              width=len(header) * 10 + 15)  # установка ширины столбца исходя длины его заголовка
        for row in table_data: self.table.insert("", "end", values=row)

        canvas.create_window((0, 0), window=self.table, anchor="nw")

        self.table.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    def update_table(self):
        self.show_table(self.last_sql_query, self.last_headers)

class WindowAbonent(ctk.CTkToplevel):
    def __init__(self, operation, select_row=None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())

        conn = sqlite3.connect("gorod.db")

        conn.close

        if select_row:
            self.select_id_abonenta = select_row[0]
            self.select_id_adresa = select_row[1]
            self.select_nomer_dom_tel = select_row[2]
            self.select_id_slyjeb_telep = select_row[3]

        if operation == "add":
            self.title("Добаление")
            ctk.CTkLabel(self, text="Добаление в таблицу 'Абонент'").grid(row=0, column=0, pady=5, padx=5,
                                                                           columnspan=2)

            ctk.CTkLabel(self, text="id абонента").grid(row=1, column=0, pady=5, padx=5)
            self.id_abonenta = ctk.CTkEntry(self, width=300)
            self.id_abonenta.grid(row=1, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="id адреса").grid(row=2, column=0, pady=5, padx=5)
            self.id_adresa = ctk.CTkEntry(self, width=300)
            self.id_adresa.grid(row=2, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="Номер домашнего телефона").grid(row=3, column=0, pady=5, padx=5)
            self.nomer_dom_tel = ctk.CTkEntry(self, width=300)
            self.nomer_dom_tel.grid(row=3, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="id служебного телефона").grid(row=4, column=0, pady=5, padx=5)
            self.id_slyjeb_telep = ctk.CTkEntry(self, width=300)
            self.id_slyjeb_telep.grid(row=4, column=1, pady=5, padx=5)

            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=5, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Добавить", width=100, command=self.add).grid(row=5, column=1, pady=5, padx=5, sticky="e")

        elif operation == "delete":
            self.title("Удаление")
            ctk.CTkLabel(self, text="Вы действиельно хотите\n удалить запись из таблицы 'Абонент'?"
                         ).grid(row=0, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkLabel(self, text=f"{self.select_id_abonenta}. {self.select_id_adresa}"
                         ).grid(row=1, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkButton(self, text="Нет", width=100, command=self.quit_win).grid(row=2, column=0, pady=5, padx=5, sticky="w")
            ctk.CTkButton(self, text="Да", width=100, command=self.delete).grid(row=2, column=1, pady=5, padx=5, sticky="e")

        elif operation == "change":
            self.title("Изменение в таблице 'Адрес'")
            ctk.CTkLabel(self, text="Назввание поля").grid(row=0, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text="Текущее значение").grid(row=0, column=1, pady=5, padx=5)
            ctk.CTkLabel(self, text="Новое занчение").grid(row=0, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="id адреса").grid(row=1, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_id_adresa).grid(row=1, column=1, pady=5, padx=5)
            self.id_adresa = ctk.CTkEntry(self, width=300)
            self.id_adresa.grid(row=1, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="Номер домашнего телефона").grid(row=2, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_nomer_dom_tel).grid(row=2, column=1, pady=5, padx=5)
            self.nomer_dom_tel = ctk.CTkEntry(self, width=300)
            self.nomer_dom_tel.grid(row=2, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="id служебного телефона").grid(row=3, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_id_slyjeb_telep).grid(row=3, column=1, pady=5, padx=5)
            self.id_slyjeb_telep = ctk.CTkEntry(self, width=300)
            self.id_slyjeb_telep.grid(row=3, column=2, pady=5, padx=5)

            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=4, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Сохранить", width=100, command=self.change).grid(row=4, column=2, pady=5, padx=5,
                                                                                       sticky="e")

    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()

    def add(self):
        new_id_adres = self.id_abonenta.get()
        new_id_nomer_dom_tel = self.id_adresa.get()
        new_nomer_dom_tel = self.nomer_dom_tel.get()
        new_id_slyjeb_telep = self.id_slyjeb_telep.get()

        if new_id_nomer_dom_tel != "" and new_nomer_dom_tel != "":
            try:
                conn = sqlite3.connect("gorod.db")
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO abonent (id_abonenta, id_adresa, nomer_dom_telephona, id_slyjeb_telep) VALUES (?, ?, ?, ?)",
                    (new_id_adres, new_id_nomer_dom_tel, new_nomer_dom_tel, new_id_slyjeb_telep))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM abonent WHERE id_abonenta = ?", (self.select_id_abonenta,))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        new_id_nomer_dom_tel = self.id_adresa.get() or self.select_id_adresa
        new_nomer_dom_tel = self.nomer_dom_tel.get() or self.select_nomer_dom_tel
        new_id_slyjeb_telep = self.id_slyjeb_telep.get() or self.select_id_slyjeb_telep
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"""
                        UPDATE abonent SET (id_adresa, nomer_dom_telephona, id_slyjeb_telep) = (?, ?, ?)  WHERE id_abonenta= {self.select_id_abonenta}
                    """, (new_id_nomer_dom_tel, new_nomer_dom_tel, new_id_slyjeb_telep))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

class WindowSlujeTel(ctk.CTkToplevel):
    def __init__(self, operation, select_row=None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())

        conn = sqlite3.connect("gorod.db")

        conn.close

        if select_row:
            self.select_id_slyj_tel = select_row[0]
            self.select_id_predp = select_row[1]
            self.select_otdel = select_row[2]
            self.select_nomer_tel = select_row[3]

        if operation == "add":
            self.title("Добаление")
            ctk.CTkLabel(self, text="Добаление в таблицу 'Служебный телефон'").grid(row=0, column=0, pady=5, padx=5,
                                                                           columnspan=2)

            ctk.CTkLabel(self, text="id служебного тел.").grid(row=1, column=0, pady=5, padx=5)
            self.id_slyj_tel = ctk.CTkEntry(self, width=300)
            self.id_slyj_tel.grid(row=1, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="id Предприятия").grid(row=2, column=0, pady=5, padx=5)
            self.id_predp = ctk.CTkEntry(self, width=300)
            self.id_predp.grid(row=2, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="Отдел").grid(row=3, column=0, pady=5, padx=5)
            self.otdel = ctk.CTkEntry(self, width=300)
            self.otdel.grid(row=3, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="Номер телефона").grid(row=4, column=0, pady=5, padx=5)
            self.nomer_tel = ctk.CTkEntry(self, width=300)
            self.nomer_tel.grid(row=4, column=1, pady=5, padx=5)

            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=5, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Добавить", width=100, command=self.add).grid(row=5, column=1, pady=5, padx=5, sticky="e")

        elif operation == "delete":
            self.title("Удаление")
            ctk.CTkLabel(self, text="Вы действиельно хотите\n удалить запись из таблицы 'Служебный телефон'?"
                         ).grid(row=0, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkLabel(self, text=f"{self.select_id_slyj_tel}. {self.select_id_predp}"
                         ).grid(row=1, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkButton(self, text="Нет", width=100, command=self.quit_win).grid(row=2, column=0, pady=5, padx=5, sticky="w")
            ctk.CTkButton(self, text="Да", width=100, command=self.delete).grid(row=2, column=1, pady=5, padx=5, sticky="e")

        elif operation == "change":
            self.title("Изменение в таблице 'Адрес'")
            ctk.CTkLabel(self, text="Назввание поля").grid(row=0, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text="Текущее значение").grid(row=0, column=1, pady=5, padx=5)
            ctk.CTkLabel(self, text="Новое занчение").grid(row=0, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="id Предприятия").grid(row=1, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_id_predp).grid(row=1, column=1, pady=5, padx=5)
            self.id_predp = ctk.CTkEntry(self, width=300)
            self.id_predp.grid(row=1, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="Отдел").grid(row=2, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_otdel).grid(row=2, column=1, pady=5, padx=5)
            self.otdel = ctk.CTkEntry(self, width=300)
            self.otdel.grid(row=2, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="Номер телефона").grid(row=3, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_nomer_tel).grid(row=3, column=1, pady=5, padx=5)
            self.nomer_tel = ctk.CTkEntry(self, width=300)
            self.nomer_tel.grid(row=3, column=2, pady=5, padx=5)

            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=4, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Сохранить", width=100, command=self.change).grid(row=4, column=2, pady=5, padx=5,
                                                                                       sticky="e")

    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()

    def add(self):
        new_id_adres = self.id_slyj_tel.get()
        new_id_nomer_dom_tel = self.id_predp.get()
        new_nomer_dom_tel = self.otdel.get()
        new_id_slyjeb_telep = self.nomer_tel.get()

        if new_id_nomer_dom_tel != "" and new_nomer_dom_tel != "":
            try:
                conn = sqlite3.connect("gorod.db")
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO slyjeb_telephon (Id_slyj_telephona, id_predp, otdel, nomer_telephona) VALUES (?, ?, ?, ?)",
                    (new_id_adres, new_id_nomer_dom_tel, new_nomer_dom_tel, new_id_slyjeb_telep))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM slyjeb_telephon WHERE Id_slyj_telephona = ?", (self.select_id_slyj_tel,))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        new_id_nomer_dom_tel = self.id_predp.get() or self.select_id_predp
        new_nomer_dom_tel = self.otdel.get() or self.select_otdel
        new_id_slyjeb_telep = self.nomer_tel.get() or self.select_nomer_tel
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"""
                        UPDATE slyjeb_telephon SET (id_predp, otdel, nomer_telephona) = (?, ?, ?)  WHERE Id_slyj_telephona = {self.select_id_slyj_tel}
                    """, (new_id_nomer_dom_tel, new_nomer_dom_tel, new_id_slyjeb_telep))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))


class WindowNasPynkt(ctk.CTkToplevel):
    def __init__(self, operation, select_row=None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())

        conn = sqlite3.connect("gorod.db")

        conn.close

        if select_row:
            self.select_id_nas_pynkta = select_row[0]
            self.select_name = select_row[1]
            self.select_id_tipa = select_row[2]
            self.select_id_ylizi = select_row[3]

        if operation == "add":
            self.title("Добаление")
            ctk.CTkLabel(self, text="Добаление в таблицу 'Населённый пункт'").grid(row=0, column=0, pady=5, padx=5,
                                                                           columnspan=2)

            ctk.CTkLabel(self, text="id населённого пункта").grid(row=1, column=0, pady=5, padx=5)
            self.id_nas_pynkta = ctk.CTkEntry(self, width=300)
            self.id_nas_pynkta.grid(row=1, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="Название населённого пункта").grid(row=2, column=0, pady=5, padx=5)
            self.name = ctk.CTkEntry(self, width=300)
            self.name.grid(row=2, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="id типа").grid(row=3, column=0, pady=5, padx=5)
            self.id_tipa = ctk.CTkEntry(self, width=300)
            self.id_tipa.grid(row=3, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="id улицы").grid(row=4, column=0, pady=5, padx=5)
            self.id_ylizi = ctk.CTkEntry(self, width=300)
            self.id_ylizi.grid(row=4, column=1, pady=5, padx=5)

            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=5, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Добавить", width=100, command=self.add).grid(row=5, column=1, pady=5, padx=5, sticky="e")

        elif operation == "delete":
            self.title("Удаление")
            ctk.CTkLabel(self, text="Вы действиельно хотите\n удалить запись из таблицы 'Населённый пункт'?"
                         ).grid(row=0, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkLabel(self, text=f"{self.select_id_nas_pynkta}. {self.select_name}"
                         ).grid(row=1, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkButton(self, text="Нет", width=100, command=self.quit_win).grid(row=2, column=0, pady=5, padx=5, sticky="w")
            ctk.CTkButton(self, text="Да", width=100, command=self.delete).grid(row=2, column=1, pady=5, padx=5, sticky="e")

        elif operation == "change":
            self.title("Изменение в таблице 'Населённый пункт'")
            ctk.CTkLabel(self, text="Назввание поля").grid(row=0, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text="Текущее значение").grid(row=0, column=1, pady=5, padx=5)
            ctk.CTkLabel(self, text="Новое занчение").grid(row=0, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="Название населённого пункта").grid(row=1, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_name).grid(row=1, column=1, pady=5, padx=5)
            self.name = ctk.CTkEntry(self, width=300)
            self.name.grid(row=1, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="id типа").grid(row=2, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_id_tipa).grid(row=2, column=1, pady=5, padx=5)
            self.id_tipa = ctk.CTkEntry(self, width=300)
            self.id_tipa.grid(row=2, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="id улицы").grid(row=3, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_id_ylizi).grid(row=3, column=1, pady=5, padx=5)
            self.id_ylizi = ctk.CTkEntry(self, width=300)
            self.id_ylizi.grid(row=3, column=2, pady=5, padx=5)

            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=4, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Сохранить", width=100, command=self.change).grid(row=4, column=2, pady=5, padx=5,
                                                                                       sticky="e")

    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()

    def add(self):
        new_id_adres = self.id_nas_pynkta.get()
        new_id_nomer_dom_tel = self.name.get()
        new_nomer_dom_tel = self.id_tipa.get()
        new_id_slyjeb_telep = self.id_ylizi.get()

        if new_id_nomer_dom_tel != "" and new_nomer_dom_tel != "":
            try:
                conn = sqlite3.connect("gorod.db")
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO id_nas_pynkt (id_nas_pynkta, naimenovanie, id_tipa, id_ylizi) VALUES (?, ?, ?, ?)",
                    (new_id_adres, new_id_nomer_dom_tel, new_nomer_dom_tel, new_id_slyjeb_telep))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM id_nas_pynkt WHERE id_nas_pynkta = ?", (self.select_id_nas_pynkta,))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        new_id_nomer_dom_tel = self.name.get() or self.select_name
        new_nomer_dom_tel = self.id_tipa.get() or self.select_id_tipa
        new_id_slyjeb_telep = self.id_ylizi.get() or self.select_id_ylizi
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"""
                        UPDATE id_nas_pynkt SET (naimenovanie, id_tipa, id_ylizi) = (?, ?, ?)  WHERE id_nas_pynkta= {self.select_id_nas_pynkta}
                    """, (new_id_nomer_dom_tel, new_nomer_dom_tel, new_id_slyjeb_telep))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))


class WindowAdres(ctk.CTkToplevel):
    def __init__(self, operation, select_row=None):
        super().__init__()
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())

        conn = sqlite3.connect("gorod.db")

        conn.close

        if select_row:
            self.select_id_adresa = select_row[0]
            self.select_id_tip_nas_pynkta = select_row[1]
            self.select_nas_pynkt = select_row[2]
            self.select_yliza = select_row[3]
            self.select_n_doma = select_row[4]

        if operation == "add":
            self.title("Добаление")
            ctk.CTkLabel(self, text="Добаление в таблицу 'Адрес'").grid(row=0, column=0, pady=5, padx=5,
                                                                           columnspan=2)

            ctk.CTkLabel(self, text="id адреса").grid(row=1, column=0, pady=5, padx=5)
            self.id_adres = ctk.CTkEntry(self, width=300)
            self.id_adres.grid(row=1, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="id населённого пунтка").grid(row=2, column=0, pady=5, padx=5)
            self.adres = ctk.CTkEntry(self, width=300)
            self.adres.grid(row=2, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="Населённый пункт").grid(row=3, column=0, pady=5, padx=5)
            self.nas_punkt = ctk.CTkEntry(self, width=300)
            self.nas_punkt.grid(row=3, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="Улица").grid(row=4, column=0, pady=5, padx=5)
            self.uliza = ctk.CTkEntry(self, width=300)
            self.uliza.grid(row=4, column=1, pady=5, padx=5)

            ctk.CTkLabel(self, text="Номер дома").grid(row=5, column=0, pady=5, padx=5)
            self.nom_dom = ctk.CTkEntry(self, width=300)
            self.nom_dom.grid(row=5, column=1, pady=5, padx=5)


            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=6, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Добавить", width=100, command=self.add).grid(row=6, column=1, pady=5, padx=5, sticky="e")

        elif operation == "delete":
            self.title("Удаление")
            ctk.CTkLabel(self, text="Вы действиельно хотите\n удалить запись из таблицы 'Адрес'?"
                         ).grid(row=0, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkLabel(self, text=f"{self.select_id_adresa}. {self.select_id_tip_nas_pynkta}"
                         ).grid(row=1, column=0, pady=5, padx=5, columnspan=2)

            ctk.CTkButton(self, text="Нет", width=100, command=self.quit_win).grid(row=2, column=0, pady=5, padx=5, sticky="w")
            ctk.CTkButton(self, text="Да", width=100, command=self.delete).grid(row=2, column=1, pady=5, padx=5, sticky="e")

        elif operation == "change":
            self.title("Изменение в таблице 'Адрес'")
            ctk.CTkLabel(self, text="Назввание поля").grid(row=0, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text="Текущее значение").grid(row=0, column=1, pady=5, padx=5)
            ctk.CTkLabel(self, text="Новое занчение").grid(row=0, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="id населённого пункта").grid(row=1, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_id_tip_nas_pynkta).grid(row=1, column=1, pady=5, padx=5)
            self.adres = ctk.CTkEntry(self, width=300)
            self.adres.grid(row=1, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="Населённый пункт").grid(row=2, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_nas_pynkt).grid(row=2, column=1, pady=5, padx=5)
            self.nas_punkt = ctk.CTkEntry(self, width=300)
            self.nas_punkt.grid(row=2, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="Улица").grid(row=3, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_yliza).grid(row=3, column=1, pady=5, padx=5)
            self.uliza = ctk.CTkEntry(self, width=300)
            self.uliza.grid(row=3, column=2, pady=5, padx=5)

            ctk.CTkLabel(self, text="Номер дома").grid(row=4, column=0, pady=5, padx=5)
            ctk.CTkLabel(self, text=self.select_n_doma).grid(row=4, column=1, pady=5, padx=5)
            self.nom_dom = ctk.CTkEntry(self, width=300)
            self.nom_dom.grid(row=4, column=2, pady=5, padx=5)

            ctk.CTkButton(self, text="Отмена", width=100, command=self.quit_win).grid(row=5, column=0, pady=5, padx=5)
            ctk.CTkButton(self, text="Сохранить", width=100, command=self.change).grid(row=5, column=2, pady=5, padx=5,
                                                                                       sticky="e")

    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()

    def add(self):
        new_id_adres = self.id_adres.get()
        new_id_nas_punkt = self.adres.get()
        new_nas_punkt = self.nas_punkt.get()
        new_uliza = self.uliza.get()
        new_n_doma = self.nom_dom.get()

        if new_id_nas_punkt != "" and new_nas_punkt != "":
            try:
                conn = sqlite3.connect("gorod.db")
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO adres (id_adresa, id_tip_nas_pynkta, nas_pynkt, yliza, N_doma) VALUES (?, ?, ?, ?, ?)",
                    (new_id_adres, new_id_nas_punkt, new_nas_punkt, new_uliza, new_n_doma))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM adres WHERE id_adresa = ?", (self.select_id_adresa,))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

    def change(self):
        new_id_nas_punkt = self.adres.get() or self.select_id_tip_nas_pynkta
        new_nas_punkt = self.nas_punkt.get() or self.select_nas_pynkt
        new_uliza = self.uliza.get() or self.select_yliza
        new_n_doma = self.nom_dom.get() or self.select_n_doma
        try:
            conn = sqlite3.connect("gorod.db")
            cursor = conn.cursor()
            cursor.execute(f"""
                        UPDATE adres SET (id_tip_nas_pynkta, nas_pynkt, yliza, N_doma) = (?, ?, ?, ?)  WHERE id_adresa= {self.select_id_adresa}
                    """, (new_id_nas_punkt, new_nas_punkt, new_uliza, new_n_doma))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

if __name__ == "__main__":
    win = WindowMain()
    win.mainloop()
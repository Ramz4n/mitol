from tkinter import ttk
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import *
import tkinter.font as tkFont
from tkinter import messagebox as mb
from tkinter.messagebox import showinfo, askyesno
import mariadb
import datetime
from datetime import date, timedelta
import time
import pandas as pd
import subprocess
from contextlib import closing
import fileinput
from tkcalendar import DateEntry
from PIL import ImageTk, Image
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import babel.numbers
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker
import json
from openpyxl import load_workbook
import pymysql
import sys
import sqlite3


class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        root.title("МиТОЛ")
        root.state("zoomed")
        root.resizable(False, True)
        self.init_main()
        self.view_records()

    def init_main(self):
        m = Menu(root)
        root.config(menu=m)
        fm = Menu(m, font=20)
        m.add_cascade(label="МЕНЮ", menu=fm)
        fm.add_command(label="Открыть в экселе", command=self.open_bd_to_excel)
        # =======1 ОСНОВНОЙ TOOLBAR====================================================================
        toolbar2 = tk.Frame(borderwidth=1, relief="raised")
        toolbar2.pack(side=tk.TOP, fill=tk.X)
        # ==========================================================================================
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute('''select id, ФИО from workers where Должность="Диспетчер"''')
                data_workers = cursor.fetchall()
        except:
            mb.showerror("Ошибка","Нет подключения к базе данных")
            sys.exit()
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''select w.ФИО from zayavki z
                JOIN workers w ON z.id_диспетчер = w.id where pc_id = {pc_id} ORDER BY z.id DESC LIMIT 1''')
                data_worker = cursor.fetchone()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.disp_to_id = {i['ФИО'] for i in data_workers}
        first_element = next(iter(self.disp_to_id), None)
        if data_worker is None:
            data = first_element
        else:
            data = data_worker['ФИО']
        toolbar3 = tk.Frame(toolbar2, borderwidth=1, relief="raised")
        toolbar3.pack(side=tk.LEFT, fill=tk.Y)
        self.disp = tk.StringVar(value=data)
        self.label3 = tk.Label(toolbar3, borderwidth=1, width=18, relief="raised", text="Диспетчер", font='Calibri 14 bold')
        frame3 = tk.Frame(toolbar3)
        self.label3.pack(side=tk.TOP)
        for d in data_workers:
            lang_btn3 = tk.Radiobutton(toolbar3, text=d['ФИО'], value=d, variable=self.disp, font='Calibri  13')
            lang_btn3.pack(anchor=tk.NW, expand=True)
            if d['ФИО'] == self.disp.get():
                lang_btn3.select()
            else:
                lang_btn3.deselect()
        frame3.pack(side=tk.LEFT, anchor=tk.NW, expand=True)
        # =======2 ГОРОДА===========================================================================
        toolbar4 = tk.Frame(toolbar2, borderwidth=1, relief="raised")
        toolbar4.pack(side=tk.LEFT, fill=tk.Y)
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute("select id, город from goroda")
                data_towns = cursor.fetchall()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        default_town = data_towns[0] if data_towns else ''
        self.var2 = tk.StringVar(value=default_town)
        self.var2.trace("w", self.on_select_city)
        self.label2 = tk.Label(toolbar4, borderwidth=1, width=18, relief="raised", text="Город", font='Calibri 14 bold')
        self.label2.pack(side=tk.TOP, anchor=tk.W)
        frame2 = tk.Frame()
        for d in data_towns:
            lang_btn4 = tk.Radiobutton(toolbar4, text=d['город'], value=d, variable=self.var2, font='Calibri  13')
            lang_btn4.pack(anchor=tk.NW, expand=True)
        frame2.pack(side=tk.LEFT, anchor=tk.NW, expand=True)
        # =======3 БЛОК АДРЕСОВ========================================================================
        toolbar5 = tk.Frame(toolbar2, borderwidth=1, relief="raised")
        toolbar5.pack(side=tk.LEFT, fill=tk.Y)
        self.label3 = tk.Label(toolbar5, borderwidth=1, width=21, relief="raised", text="Адрес", font='Calibri 14 bold')
        self.frame3 = tk.Frame()
        self.entry_text3 = tk.StringVar()
        self.entry3 = tk.Entry(toolbar5, textvariable=self.entry_text3, width=33)
        self.entry3.bind('<KeyRelease>', self.check_input_address)
        self.label3.pack(side=tk.TOP)
        self.entry3.pack(side=tk.TOP, expand=True)
        self.listbox_values = tk.Variable()
        self.listbox = tk.Listbox(toolbar5, listvariable=self.listbox_values, width=25, font='Calibri 12')
        self.listbox.bind('<<ListboxSelect>>', self.on_change_selection_address)
        self.listbox.pack(side=tk.TOP, expand=True)
        self.frame3.pack(side=tk.LEFT, anchor=tk.NW, expand=True)
        # ======4 БЛОК КОД С ТИПАМИ ЛИФТОВ==================================================================
        toolbar6 = tk.Frame(toolbar2, borderwidth=1, relief="raised")
        toolbar6.pack(side=tk.LEFT, fill=tk.Y)
        self.label4 = tk.Label(toolbar6, borderwidth=1, width=12, relief="raised", text="Тип лифта", font='Calibri 14 bold')
        self.frame4 = tk.Frame()
        self.entry_text4 = tk.StringVar()
        self.entry4 = tk.Entry(toolbar6, textvariable=self.entry_text4, width=18)
        self.entry4.bind('<KeyRelease>', self.check_input_lifts)
        self.label4.pack(side=tk.TOP)
        self.entry4.pack(side=tk.TOP, expand=True)
        self.listbox_values_type = tk.Variable()
        self.listbox_type = tk.Listbox(toolbar6, listvariable=self.listbox_values_type, width=14, font='Calibri 12')
        self.listbox_type.bind('<<ListboxSelect>>', self.on_change_selection_lift)
        self.listbox_type.pack(side=tk.TOP, expand=True)
        self.frame4.pack(side=tk.LEFT, anchor=tk.NW, expand=True)
        # ======5 БЛОК ПРИЧИНА ОСТАНОВКИ ============================================
        toolbar7 = tk.Frame(toolbar2, borderwidth=1, relief="raised")
        toolbar7.pack(side=tk.LEFT, fill=tk.Y)
        frame5 = tk.Frame()
        label5 = tk.Label(toolbar7, borderwidth=1, width=14, relief="raised", text="Причина", font='Calibri 14 bold')
        label5.pack()
        self.prichina = ['Неисправность', 'Застревание', 'Остановлен', 'Связь']
        self.prich5 = tk.StringVar(value='?')
        for pr in self.prichina:
            lang_btn3 = tk.Radiobutton(toolbar7, text=pr, value=pr, variable=self.prich5, font='Calibri  13')
            lang_btn3.pack(anchor=tk.NW, expand=True)
        frame5.pack(side=tk.LEFT, anchor=tk.NW, expand=True)
        # ======6 БЛОК FIO МЕХАНИКА =====================================================
        toolbar9 = tk.Frame(toolbar2, borderwidth=1, relief="raised")
        toolbar9.pack(side=tk.LEFT, fill=tk.Y)
        self.label7 = tk.Label(toolbar9, borderwidth=1, width=18, relief="raised", text="ФИО механика", font='Calibri  14 bold')
        self.frame7 = tk.Frame()
        self.entry_text7 = tk.StringVar(value='')
        self.entry7 = tk.Entry(toolbar9, textvariable=self.entry_text7, width=28)
        self.entry7.bind('<KeyRelease>', self.check_input_fio)
        self.label7.pack(side=tk.TOP)
        self.entry7.pack(side=tk.TOP, expand=True)
        self.listbox_values7 = tk.Variable()
        self.listbox7 = tk.Listbox(toolbar9, width=21, listvariable=self.listbox_values7, font='Calibri 12')
        self.listbox7.bind('<<ListboxSelect>>', self.on_change_selection_fio)
        self.listbox7.pack(side=tk.TOP, expand=True)
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute("select id, ФИО from workers where Должность = 'Механик' and id_ЛК = 1 order by ФИО")
                self.data_meh = cursor.fetchall()
            for d in self.data_meh:
                self.data_meh_name = f"{d['ФИО']}"
                self.data_meh_id = f"{d['id']}"
                self.listbox7.insert(tk.END, self.data_meh_name)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.frame7.pack(side=tk.LEFT, anchor=tk.NW, expand=True)
        # =======NEW TOOLBAR==============================================================
        toolbar = tk.Frame(bd=2, borderwidth=1, relief="raised")
        toolbar.pack(side=tk.TOP, fill=tk.X, anchor=tk.N)
        helv36 = tkFont.Font(family='Helvetica', size=10, weight=tkFont.BOLD)
        # ================КНОПКИ========================================================
        tool1 = tk.Frame(toolbar, borderwidth=1, relief="raised")
        tool1.pack(side=tk.LEFT, fill=tk.X, anchor=tk.W)
        btn_open_dialog = tk.Button(tool1, text='Добавить заявку', command=self.sql_insert, bg='#d7d8e0', compound=tk.LEFT, width=14, height=1, font=helv36)
        btn_open_dialog.pack(side=tk.TOP)
        btn_refresh = tk.Button(tool1, text='Обновить', bg='#d7d8e0', compound=tk.TOP, command=self.view_records, width=14, font=helv36)
        btn_refresh.pack(side=tk.BOTTOM)
        btn_search = tk.Button(tool1, text='Поиск адреса', bg='#d7d8e0', compound=tk.TOP, command=self.open_search_dialog, width=14, font=helv36)
        btn_search.pack(side=tk.TOP)
        # =================КНОПКИ========================================================
        tool3 = tk.Frame(toolbar, borderwidth=1, relief="raised")
        tool3.pack(side=tk.LEFT, fill=tk.X, anchor=tk.W)
        btn_refresh = tk.Button(tool3, text='Запущенные лифты', bg='#00AD0E', compound=tk.TOP,command=self.start_lift, width=19, font=helv36)
        btn_refresh.pack(side=tk.BOTTOM)
        btn_refresh = tk.Button(tool3, text='Остановленные лифты', bg='#FFB3AB', compound=tk.TOP,command=self.stop_lift, width=19, font=helv36)
        btn_refresh.pack(side=tk.TOP)
        btn_refresh = tk.Button(tool3, text='НЕ Закрытые заявки', bg='#4897FF', compound=tk.TOP,command=self.non_start_lift, width=19, font=helv36)
        btn_refresh.pack(side=tk.BOTTOM)
        #=====================================================================================
        tool4 = tk.Frame(toolbar, borderwidth=1, relief="raised")
        tool4.pack(side=tk.LEFT, fill=tk.X, anchor=tk.W)
        btn_refresh = tk.Button(tool4, text='заявки до 25.03.24', bg='#d7d8e0', compound=tk.TOP,
        command=self.view_records_old, width=19, font=helv36)
        btn_refresh.pack(side=tk.BOTTOM)
        self.is_on = True
        self.enabled = IntVar()
        self.enabled.set(pc_id)
        self.my_label = Label(tool4,
                         text="Мои заявки",
                         fg="green",
                         font=("Helvetica", 10))
        self.my_label.pack()
        self.on = PhotoImage(file="on.png")
        self.off = PhotoImage(file="off.png")
        self.on_button = Button(tool4, image=self.off, bd=0, command=self.switch)
        self.on_button.pack()
        # === ПЕРЕЛИСТЫВАНИЕ БД ПО МЕСЯЦАМ=====================================================================
        self.months = ["Январь", "Февраль", "Март", "Апрель",
                       "Май", "Июнь", "Июль", "Август",
                       "Сентябрь","Октябрь","Ноябрь", "Декабрь"]
        btn_refresh = tk.Button(toolbar, text='Следующий месяц', bg='#d7d8e0', compound=tk.RIGHT, command=self.update_forward, width=19, font=helv36)
        btn_refresh.pack(side=tk.RIGHT)
        self.year_label = Label(toolbar, text='', font='Calibri 16 bold')
        self.year_label.pack(side=tk.RIGHT)
        self.month_label = Label(toolbar, text='', font='Calibri 16 bold')
        self.month_label.pack(side=tk.RIGHT)
        btn_refresh = tk.Button(toolbar, text='Предыдущий месяц', bg='#d7d8e0', compound=tk.RIGHT, command=self.update_backward, width=19, font=helv36)
        btn_refresh.pack(side=tk.RIGHT)
        # =======ИНФО СПРАВА О ПОДСЧЁТАХ============================================================================
        self.calendar = DateEntry(toolbar2, locale='ru_RU', font=1)
        self.calendar.bind("<<DateEntrySelected>>", self.print_info)
        self.calendar.pack(side=tk.TOP, anchor=tk.N)
        self.label_info_bd = tk.Label(toolbar2, borderwidth=1, width=50, height=10, relief="raised", text="", font='Times 13')
        style2 = ttk.Style()
        style2.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Helvetica', 12))
        style2.configure("mystyle.Treeview.Heading", font=('Helvetica', 12, 'bold'))
        style2.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])
        self.tree2 = ttk.Treeview(self.label_info_bd, style="mystyle.Treeview", columns=('Город', 'Застр', 'Неиспр', 'Кол_во'),
                                  height=7, show='headings')
        self.tree2.column('Город', width=115, anchor=tk.CENTER)
        self.tree2.column('Застр', width=120, anchor=tk.CENTER)
        self.tree2.column('Неиспр', width=145, anchor=tk.CENTER)
        self.tree2.column('Кол_во', width=70, anchor=tk.CENTER)

        self.tree2.heading('Город', text='Город')
        self.tree2.heading('Застр', text='Застреваний')
        self.tree2.heading('Неиспр', text='Неисправностей')
        self.tree2.heading('Кол_во', text='Кол-во')
        self.tree2.pack(side="left", fill="both")
        self.label_info_bd.pack(side=tk.TOP)
        # =======ВИЗУАЛ БАЗЫ ДАННЫХ =========================================================================
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Helvetica', 12))
        style.configure("mystyle.Treeview.Heading", font=('Helvetica', 12, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])
        self.tree = ttk.Treeview(self, style="mystyle.Treeview",
        columns=('ID', 'date', 'dispetcher', 'town', 'adress', 'type_lift', 'prichina', 'fio', 'date_to_go', 'comment', 'id2'),
                                 height=50, show='headings')
        self.tree.bind("<Button-3>", self.show_menu)
        self.tree.column('ID', width=50, anchor=tk.CENTER, stretch=False)
        self.tree.column('date', width=135, anchor=tk.CENTER, stretch=False)
        self.tree.column('dispetcher', width=120, anchor=tk.CENTER, stretch=False)
        self.tree.column('town', width=120, anchor=tk.CENTER, stretch=False)
        self.tree.column('adress', width=250, anchor=tk.CENTER, stretch=False)
        self.tree.column('type_lift', width=100, anchor=tk.CENTER, stretch=False)
        self.tree.column('prichina', width=120, anchor=tk.CENTER, stretch=False)
        self.tree.column('fio', width=170, anchor=tk.CENTER, stretch=False)
        self.tree.column('date_to_go', width=135, anchor=tk.CENTER, stretch=False)
        self.tree.column("comment", width=1000, anchor=tk.W, stretch=True)
        self.tree.column("id2", width=0, anchor=tk.CENTER)
        self.tree.column('#0', stretch=False)

        self.tree.heading('ID', text='№')
        self.tree.heading('date', text='Дата заявки')
        self.tree.heading('dispetcher', text='Диспетчер')
        self.tree.heading('town', text='Город')
        self.tree.heading('adress', text='Адрес')
        self.tree.heading('type_lift', text='Тип лифта')
        self.tree.heading('prichina', text='Причина')
        self.tree.heading('fio', text='ФИО механика')
        self.tree.heading('date_to_go', text='Дата запуска')
        self.tree.heading('comment', text='Комментарий', anchor=tk.W)
        self.tree.heading('id2', text='')

        self.scrollbar2 = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.tree.configure(yscrollcommand=self.scrollbar2.set)
        self.scrollbar2.config(command=self.tree.yview)
        self.scrollbar2.pack(side="right", fill="both")

        self.scrollbar = tk.Scrollbar(self, orient=tk.HORIZONTAL)
        self.tree.configure(xscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.tree.xview)
        self.scrollbar.pack(side="bottom", fill="both")
        self.tree.pack(side="left", fill="both")
        self.on_select_city()

    def switch(self):
        pc = {1, 2}
        if self.is_on:
            self.on_button.config(image=self.on)
            self.my_label.config(text="Чужие заявки", fg="red")
            new_value = (self.enabled.get() % len(pc)) + 1
            self.enabled.set(new_value)
            self.view_records()
            self.is_on = False
        else:
            self.on_button.config(image=self.off)
            self.my_label.config(text="Мои заявки", fg="green")
            new_value = (self.enabled.get() % len(pc))  + 1
            self.enabled.set(new_value)
            self.view_records()
            self.is_on = True

    def view_records_old(self):
        self.current_month_index = int((datetime.datetime.now(tz=None)).strftime("%m")) - 1
        self.current_year_index = int((datetime.datetime.now(tz=None)).strftime("%Y"))
        self.month_label.config(text=self.months[(self.current_month_index) % 12])
        self.year_label.config(text=self.current_year_index)
        self.tree.tag_configure("Red.Treeview", foreground="red")
        self.tree.tag_configure("Yellow.Treeview", foreground="yellow")
        date_obj = datetime.datetime.now()
        formatted_date = date_obj.strftime('%Y-%m-%d')
        try:
            with closing(sqlite3.connect('mitol.db')) as connection2:
                cursor = connection2.cursor()
                cursor.execute('''SELECT Номер_заявки,
                                    strftime('%d.%m.%Y, %H:%M', datetime(Дата_заявки, 'unixepoch')) AS Дата_заявки2,
                                    Диспетчер,
                                    Город,
                                    Адрес,
                                    Тип_лифта,
                                    Причина,
                                    ФИО_механика,
                                    COALESCE(strftime('%d.%m.%Y, %H:%M', datetime(Дата_запуска, 'unixepoch')), '') AS Дата_запуска2,
                                    Комментарий,
                                    id,
                                    Дата_запуска - Дата_заявки AS Разница,
                                    Дата_заявки
                                    FROM balash''')
                [self.tree.delete(i) for i in self.tree.get_children()]
                for row in cursor.fetchall():
                    if row[-5] == "" and row[-1] < int((time.time() + 3 * 60 * 60) - 86400):
                        self.tree.insert('', 'end', values=tuple(row), tags=('Red.Treeview',))
                    else:
                        self.tree.insert('', 'end', values=tuple(row))
                self.tree.yview_moveto(1.0)
                connection2.commit()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.print_info(formatted_date)
#ФУНКЦИЯ ПО ПОЛУЧЕНИЮ ГОРОДА И ДАЛЬНЕЙШЕМУ ПАРСИНГУ АДРЕСОВ В ОКОШКО ПО ГОРОДАМ
    def on_select_city(self, *args):
        selected_city_str = eval(self.var2.get())
        self.selected_city_id = selected_city_str['id']
        self.selected_city = selected_city_str['город']
        self.listbox.delete(0, tk.END)
        self.listbox_type.delete(0, tk.END)
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''
                    SELECT goroda.id as goroda_id, goroda.город, 
                    street.id as street_id, street.улица, 
                    doma.id as doma_id, doma.номер as дом, 
                    padik.id as padik_id, padik.номер as подъезд
                    FROM goroda
                    JOIN street ON goroda.id = street.город_id
                    JOIN doma ON street.id = doma.улица_id
                    JOIN padik ON doma.id = padik.дом_id
                    WHERE goroda.город = "{self.selected_city}" order BY street.улица, doma.`номер`, padik.`номер`''')
                self.data_streets = cursor.fetchall()
                for d in self.data_streets:
                    self.address_str = f"{d['улица']}, {d['дом']}, {d['подъезд']}"
                    self.listbox.insert(tk.END, self.address_str)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    def show_menu(self, event):
        menu = tk.Menu(self.tree, tearoff=False, font=20)
        settings_menu = tk.Menu(tearoff=False, font=20)
        settings_menu.add_command(label="ошибка а0", command=lambda: self.error("ошибка а0"))
        settings_menu.add_command(label="ошибка a2", command=lambda: self.error("ошибка a2"))
        settings_menu.add_command(label="ошибка 43", command=lambda: self.error("ошибка 43"))
        settings_menu.add_command(label="ошибка 44", command=lambda: self.error("ошибка 44"))
        settings_menu.add_command(label="ошибка 45", command=lambda: self.error("ошибка 45"))
        settings_menu.add_command(label="ошибка 48", command=lambda: self.error("ошибка 48"))
        settings_menu.add_command(label="ошибка 49", command=lambda: self.error("ошибка 49"))
        settings_menu.add_command(label="ошибка 50", command=lambda: self.error("ошибка 50"))
        settings_menu.add_command(label="ошибка 53", command=lambda: self.error("ошибка 53"))
        settings_menu.add_command(label="ошибка 56", command=lambda: self.error("ошибка 56"))
        settings_menu.add_command(label="ошибка 67", command=lambda: self.error("ошибка 67"))
        settings_menu.add_command(label="ошибка 70", command=lambda: self.error("ошибка 70"))
        settings_menu.add_command(label="ошибка 71", command=lambda: self.error("ошибка 71"))
        settings_menu.add_command(label="ошибка 96", command=lambda: self.error("ошибка 96"))
        settings_menu.add_command(label="ошибка 99", command=lambda: self.error("ошибка 99"))
        settings_menu.add_command(label="Частотник", command=lambda: self.error("Частотник"))
        settings_menu.add_command(label="Многократный реверс", command=lambda: self.error("Многократный реверс"))
        settings_menu.add_command(label="Пожарная сигнализация", command=lambda: self.error("Пожарная сигнализация"))
        settings_menu.add_command(label="До наладчика", command=lambda: self.error("До наладчика"))
        settings_menu.add_command(label="Ревизия/инспекция", command=lambda: self.error("Лифт в ревизии/инспекции"))
        settings_menu.add_command(label="Аварийная блокировка", command=lambda: self.error("Аварийная блокировка"))
        menu.add_cascade(label="Ошибка", command=lambda: self.error("Ошибка"), menu=settings_menu)
        menu.add_command(label="Редактировать", command=lambda: self.edit("Редактировать"))
        menu.add_command(label="Отметить Время", command=lambda: self.time_to("Отметить Время"))
        menu.add_command(label="Комментировать", command=lambda: self.open_comment("Комментировать"))
        menu.add_separator()
        menu.add_command(label="Ложная Заявка", command=lambda: self.lojnaya("Ложная Заявка"))
        menu.add_command(label="Отсутствие электроэнергии", command=lambda: self.error("Отсутствие электроэнергии"))
        menu.add_command(label="Вандальные действия", command=lambda: self.error("Вандальные действия"))
        menu.add_separator()
        menu.add_command(label="Удалить Заявку", command=lambda: self.delete("Удалить Заявку"))
        menu.post(event.x_root, event.y_root)


    # ===РЕДАКТИРОВАТЬ========================================================================
    def edit(self, event):
        if self.tree.selection():
            try:
                with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                    cursor = connection2.cursor(dictionary=True)
                    cursor.execute(f'''SELECT z.id,
                                       FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                       w.ФИО AS Диспетчер,
                                       g.город AS Город,
                                       CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                       тип_лифта,
                                       причина,
                                       m.ФИО,
                                       FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS дата_запуска,
                                       комментарий
                                FROM zayavki z
                                JOIN workers w ON z.id_диспетчер = w.id
                                JOIN goroda g ON z.id_город = g.id
                                JOIN street s ON z.id_улица = s.id
                                JOIN doma d ON z.id_дом = d.id
                                JOIN padik p ON z.id_подъезд = p.id
                                JOIN workers m ON z.id_механик = m.id
                                     where z.id=?''',
                                   (self.tree.set(self.tree.selection()[0], '#11'),))
                    rows = cursor.fetchall()
                    Child(rows)
            except mariadb.Error as e:
                showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        else:
            mb.showerror("Ошибка","Строка для редактирования не выбрана")
            return

    # ===ВСТАВКА ВРЕМЕНИ В БД=======================================================================
    def time_to(self, event):
        date_ = (datetime.datetime.now(tz=None)).strftime("%d.%m.%Y, %H:%M")
        time_obj = datetime.datetime.strptime(date_, time_format)
        unix_time = int(time_obj.timestamp())
        if self.tree.selection():
            try:
                with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                    cursor = connection2.cursor()
                    cursor.execute(f'''SELECT Дата_запуска from {table_zayavki} WHERE id=?''',
                                   [self.tree.set(self.tree.selection()[0], '#11')])
                    info = cursor.fetchone()
                    connection2.commit()
                    if info[0] is None:
                        try:
                            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                                cursor = connection2.cursor()
                                cursor.execute(
                                    f'''UPDATE {table_zayavki} set Дата_запуска=? WHERE ID=?''',
                                    [unix_time, self.tree.set(self.tree.selection()[0], '#11')])
                                connection2.commit()
                        except mariadb.Error as e:
                            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
                    else:
                        mb.showerror("Ошибка","Строка со временем уже заполнена")
                        return
            except mariadb.Error as e:
                showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        else:
            mb.showerror(
                "Ошибка",
                "Строка не выбрана")
            return
        self.view_records()
        msg = f"Время отмечено!"
        mb.showinfo("Информация", msg)

    # ===УДАЛЕНИЕ СТРОКИ=====================================================================
    def delete(self, event):
        if self.tree.selection():
            result = askyesno(title="Подтвержение операции", message="УДАЛИТЬ строчку?")
            if result:
                try:
                    with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                        cursor = connection2.cursor()
                        cursor.execute(f'''delete from {table_zayavki} WHERE ID=?''',
                                       [self.tree.set(self.tree.selection()[0], '#11')])
                        connection2.commit()
                except mariadb.Error as e:
                    showinfo('Информация', f"Ошибка удаления записи: {e}")
            else:
                showinfo("Результат", "Операция отменена")
        else:
            mb.showerror(
                "Ошибка",
                "Строка не выбрана")
            return
        self.view_records()

    # ===ОТМЕТИТЬ ЛОЖНУЮ==============================================================================
    def lojnaya(self, event):
        date_ = (datetime.datetime.now(tz=None)).strftime("%d.%m.%Y, %H:%M")
        time_obj = datetime.datetime.strptime(date_, time_format)
        unix_time = int(time_obj.timestamp())
        if self.tree.selection():
            try:
                with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                    cursor = connection2.cursor()
                    cursor.execute(f'''UPDATE {table_zayavki} set Дата_запуска=?, Комментарий=? WHERE ID=?''',
                                   [unix_time, f'{event}', self.tree.set(self.tree.selection()[0], '#11')])
                    connection2.commit()
            except mariadb.Error as e:
                showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        else:
            mb.showerror("Ошибка","Строка не выбрана")
            return
        self.view_records()
        msg = f"Запись отредактирована!"
        mb.showinfo("Информация", msg)

    # ===ОТМЕТИТЬ ОШИБКУ==============================================================================
    def error(self, event):
        print(event)
        if self.tree.selection():
            try:
                with closing(mariadb.connect(user=user, password=password, host=host, port=port,
                                             database=database)) as connection2:
                    cursor = connection2.cursor()
                    cursor.execute(f'''UPDATE {table_zayavki} set Комментарий=? WHERE ID=?''',
                                   [f'{event}', self.tree.set(self.tree.selection()[0], '#11')])
                    connection2.commit()
            except mariadb.Error as e:
                showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        else:
            mb.showerror("Ошибка", "Строка не выбрана")
            return
        self.view_records()
        msg = f"Запись отредактирована!"
        mb.showinfo("Информация", msg)

    # ===ФУНКЦИЯ КОММЕНТИРОВАНИЯ==================================================
    def open_comment(self, event):
        if self.tree.selection():
            Comment()
        else:
            mb.showerror("Ошибка","Строка не выбрана")
            return

    def update_forward(self):
        self.current_month_index = (self.current_month_index + 1) % 12
        self.month_label.config(text=self.months[self.current_month_index])
        if self.current_month_index == 0:
            self.current_year_index += 1
            self.year_label.config(text=self.current_year_index)
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f'''SELECT z.Номер_заявки,
                                   FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                   w.ФИО AS Диспетчер,
                                   g.город AS Город,
                                   CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                   тип_лифта,
                                   причина,
                                   m.ФИО,
                                   FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                                   комментарий,
                                   z.id
                            FROM zayavki z
                            JOIN workers w ON z.id_диспетчер = w.id
                            JOIN goroda g ON z.id_город = g.id
                            JOIN street s ON z.id_улица = s.id
                            JOIN doma d ON z.id_дом = d.id
                            JOIN padik p ON z.id_подъезд = p.id
                            JOIN workers m ON z.id_механик = m.id
                                WHERE FROM_UNIXTIME(Дата_заявки, '%m') = ?
                                and FROM_UNIXTIME(Дата_заявки, '%Y') = ? and z.pc_id = ?
                                order by z.id;''',
                               (f'{str(self.current_month_index + 1).zfill(2)}', f'{str(self.current_year_index)}', self.enabled.get()))
                [self.tree.delete(i) for i in self.tree.get_children()]
                for row in cursor.fetchall():
                    if row[-3] is None and int((datetime.datetime.strptime(row[1], time_format)).timestamp()) < int(
                            (time.time()) - 86400):
                        self.tree.insert('', 'end', values=tuple(row), tags=('Red.Treeview',))
                    elif row[-3] is None:
                        self.tree.insert('', 'end', values=tuple(row), tags=('Yellow.Treeview',))
                    elif row[-3] != None and row[-5] == 'Остановлен':
                        self.tree.insert('', 'end', values=tuple(row), tags=('Green.Treeview',))
                    else:
                        self.tree.insert('', 'end', values=tuple(row))
                connection.commit()
                self.tree.yview_moveto(1.0)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    def update_backward(self):
        self.current_month_index = (self.current_month_index - 1) % 12
        self.month_label.config(text=self.months[self.current_month_index])
        if self.current_month_index == 11:
            self.current_year_index -= 1
            self.year_label.config(text=self.current_year_index)
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port,
                                         database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f'''SELECT z.Номер_заявки,
                                   FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                   w.ФИО AS Диспетчер,
                                   g.город AS Город,
                                   CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                   тип_лифта,
                                   причина,
                                   m.ФИО,
                                   FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                                   комментарий,
                                   z.id
                            FROM zayavki z
                            JOIN workers w ON z.id_диспетчер = w.id
                            JOIN goroda g ON z.id_город = g.id
                            JOIN street s ON z.id_улица = s.id
                            JOIN doma d ON z.id_дом = d.id
                            JOIN padik p ON z.id_подъезд = p.id
                            JOIN workers m ON z.id_механик = m.id
                                WHERE FROM_UNIXTIME(Дата_заявки, '%m') = ?
                                and FROM_UNIXTIME(Дата_заявки, '%Y') = ? and z.pc_id = ?
                                order by z.id;''',
                               (f'{str(self.current_month_index + 1).zfill(2)}', f'{str(self.current_year_index)}',
                                self.enabled.get()))
                [self.tree.delete(i) for i in self.tree.get_children()]
                for row in cursor.fetchall():
                    if row[-3] is None and int((datetime.datetime.strptime(row[1], time_format)).timestamp()) < int(
                            (time.time()) - 86400):
                        self.tree.insert('', 'end', values=tuple(row), tags=('Red.Treeview',))
                    elif row[-3] is None:
                        self.tree.insert('', 'end', values=tuple(row), tags=('Yellow.Treeview',))
                    elif row[-3] != None and row[-5] == 'Остановлен':
                        self.tree.insert('', 'end', values=tuple(row), tags=('Green.Treeview',))
                    else:
                        self.tree.insert('', 'end', values=tuple(row))
                connection.commit()
                self.tree.yview_moveto(1.0)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    def print_info(self, event):
        self.selected_date = self.calendar.get_date()
        self.selected_date_str = self.selected_date.strftime('%d.%m.%Y')
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f'''SELECT g.город as Город,
                                    SUM(CASE WHEN z.Причина = 'Застревание' THEN 1 ELSE 0 END) as Застр,
                                    SUM(CASE WHEN z.Причина IN ('Остановлен', 'Неисправность') THEN 1 ELSE 0 END) as Неиспр,
                                    SUM(CASE WHEN z.Причина IN ('Остановлен', 'Неисправность', 'Застревание') THEN 1 ELSE 0 END) as Кол_во
                                    FROM zayavki z
                                    JOIN goroda g ON z.id_город = g.id
                                    WHERE DATE_FORMAT(FROM_UNIXTIME(z.Дата_заявки), '%d.%m.%Y') = ?
                                    GROUP BY Город;''',
                (self.selected_date_str,))
                [self.tree2.delete(i) for i in self.tree2.get_children()]
                [self.tree2.insert('', 'end', values=row) for row in cursor.fetchall()]
                connection.commit()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    # ===ОСТАНОВЛЕННЫЕ ЛИФТЫ==================================================================================
    def stop_lift(self):
        self.tree.tag_configure("Red.Treeview", foreground="red")
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f'''SELECT z.Номер_заявки,
                                       FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                       w.ФИО AS Диспетчер,
                                       g.город AS Город,
                                       CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                       тип_лифта,
                                       причина,
                                       m.ФИО,
                                       FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                                       комментарий,
                                       z.id
                                FROM zayavki z
                                JOIN workers w ON z.id_диспетчер = w.id
                                JOIN goroda g ON z.id_город = g.id
                                JOIN street s ON z.id_улица = s.id
                                JOIN doma d ON z.id_дом = d.id
                                JOIN padik p ON z.id_подъезд = p.id
                                JOIN workers m ON z.id_механик = m.id
                                WHERE Причина="Остановлен"
                                and Дата_запуска is Null and z.pc_id = ?
                                order by z.id;''', (self.enabled.get(),))
                [self.tree.delete(i) for i in self.tree.get_children()]
                for row in cursor.fetchall():
                    self.tree.insert('', 'end', values=tuple(row), tags=('Red.Treeview',))
                    connection.commit()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    # ===НЕ ЗАКРЫТЫЕ ЗАЯВКИ==================================================================================
    def non_start_lift(self):
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor()
                cursor.execute(f'''SELECT z.Номер_заявки,
                                       FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                       w.ФИО AS Диспетчер,
                                       g.город AS Город,
                                       CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                       тип_лифта,
                                       причина,
                                       m.ФИО,
                                       FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                                       комментарий,
                                       z.id,
                                       z.Дата_заявки
                                FROM zayavki z
                                JOIN workers w ON z.id_диспетчер = w.id
                                JOIN goroda g ON z.id_город = g.id
                                JOIN street s ON z.id_улица = s.id
                                JOIN doma d ON z.id_дом = d.id
                                JOIN padik p ON z.id_подъезд = p.id
                                JOIN workers m ON z.id_механик = m.id
                                WHERE Дата_запуска is Null and Причина<>"Остановлен" and z.pc_id = ?
                                order by z.id;''', (self.enabled.get(),))
                [self.tree.delete(i) for i in self.tree.get_children()]
                for row in cursor.fetchall():
                    if row[-4] == None and row[-1] < int((time.time()) - 86400):
                        self.tree.insert('', 'end', values=tuple(row), tags=('Red.Treeview',))
                    else:
                        self.tree.insert('', 'end', values=tuple(row), tags=('Yellow.Treeview',))
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    # ===ЗАПУЩЕННЫЕ ЛИФТЫ==================================================================================
    def start_lift(self):
        self.tree.tag_configure("Green.Treeview", foreground="#06B206")
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f'''SELECT z.Номер_заявки,
                                       FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                       w.ФИО AS Диспетчер,
                                       g.город AS Город,
                                       CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                       тип_лифта,
                                       причина,
                                       m.ФИО,
                                       FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                                       комментарий,
                                       z.id
                                FROM zayavki z
                                JOIN workers w ON z.id_диспетчер = w.id
                                JOIN goroda g ON z.id_город = g.id
                                JOIN street s ON z.id_улица = s.id
                                JOIN doma d ON z.id_дом = d.id
                                JOIN padik p ON z.id_подъезд = p.id
                                JOIN workers m ON z.id_механик = m.id
                                where Причина="Остановлен" 
                                and Дата_запуска is not Null and z.pc_id = ?
                                order by z.id;''', (self.enabled.get(),))
                [self.tree.delete(i) for i in self.tree.get_children()]
                for row in cursor.fetchall():
                    self.tree.insert('', 'end', values=row, tags=('Green.Treeview',))
                connection.commit()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    # ===Открытие файла excel================================================================
    def open_bd_to_excel(self):
        if self.tree.selection():
            selected_items = self.tree.selection()
            selected_ids = [self.tree.item(item, 'values')[-1] for item in selected_items]
            selected_id_str = ', '.join([f'"{id}"' for id in selected_ids])
            conn = pymysql.connect(user=user, password=password, host=host, port=port, database=database)
            cursor = conn.cursor()
            sql_query = f'''SELECT z.Номер_заявки,
                                   FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                   w.ФИО AS Диспетчер,
                                   g.город AS Город,
                                   CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                   Тип_лифта,
                                   Причина,
                                   m.ФИО as Механик,
                                   FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                                   Комментарий,
                                   z.id
                            FROM zayavki z
                            JOIN workers w ON z.id_диспетчер = w.id
                            JOIN goroda g ON z.id_город = g.id
                            JOIN street s ON z.id_улица = s.id
                            JOIN doma d ON z.id_дом = d.id
                            JOIN padik p ON z.id_подъезд = p.id
                            JOIN workers m ON z.id_механик = m.id 
                                    WHERE z.id in ({selected_id_str})'''
            cursor.execute(sql_query)
            data = cursor.fetchall()
            df = pd.DataFrame(data, columns=[i[0] for i in cursor.description])
            conn.close()
            Excel(df)
        else:
            msg = f"Нужно выделить одну или несколько строк, которые нужно вставить в excel"
            mb.showerror("Ошибка!", msg)

    # ===РЕДАКТИРОВАНИЕ ДАННЫХ В БД===================================================================
    def update_record(self, data, dispetcher, town, street, house, padik, type_lift, prichina, fio_meh, date_to_go, comment):
        try:
            date_object = datetime.datetime.strptime(data, time_format)
            if date_to_go == None or date_to_go == '':
                date_to_go = None
                val1 = (int(date_object.timestamp()),
                        dispetcher, town, street, house,
                        padik, type_lift, prichina, fio_meh,
                        date_to_go, comment)
            else:
                try:
                    date_object2 = datetime.datetime.strptime(date_to_go, time_format)
                    val1 = (int(date_object.timestamp()),
                            dispetcher, town, street, house,
                            padik, type_lift, prichina,fio_meh,
                        int(date_object2.timestamp()), comment)
                except ValueError:
                    msg = "Введите дату в формате ДД.ММ.ГГГГ, ЧЧ:ММ или нажмите на нужную заявку, а потом на кнопку 'Отметить время'"
                    mb.showerror("Ошибка", msg)
                    return
        except ValueError:
            msg = "Введите дату в формате ДД.ММ.ГГГГ, ЧЧ:ММ или нажмите на нужную заявку, а потом на кнопку 'Отметить время'"
            mb.showerror("Ошибка", msg)
            return
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f'''UPDATE zayavki z
                                    JOIN workers w_disp ON z.id_Диспетчер = w_disp.id
                                    JOIN goroda g ON z.id_город = g.id
                                    JOIN street s ON z.id_улица = s.id
                                    JOIN doma d ON z.id_дом = d.id
                                    JOIN padik p ON z.id_подъезд = p.id
                                    JOIN workers w_mech ON z.id_Механик = w_mech.id
                                    SET z.Дата_заявки = ?, 
                                        z.id_Диспетчер = ?, 
                                        z.id_город = ?, 
                                        z.id_улица = ?, 
                                        z.id_дом = ?, 
                                        z.id_подъезд = ?, 
                                        z.тип_лифта = ?, 
                                        z.Причина = ?, 
                                        z.id_Механик = ?, 
                                        z.Дата_запуска = ?, 
                                        z.Комментарий = ?
                                    WHERE z.ID = ?;''',
                               (val1 + (self.tree.set(self.tree.selection()[0], '#11'),)))
                connection.commit()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.view_records()
        msg = f"Запись отредактирована!"
        mb.showinfo("Информация", msg)

    # ===ФУНКЦИЯ КНОПКИ ЧЕРЕЗ КОММЕНТИРОВАНИЯ================================================
    def comment(self, comment):
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor()
                cursor.execute(f'''UPDATE {table_zayavki} SET Комментарий=? WHERE ID=?''',
                               (comment, self.tree.set(self.tree.selection()[0], '#11')))
                connection2.commit()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.view_records()
        msg = f"Комментарий добавлен!"
        mb.showinfo("Информация", msg)

    # ===ФУНКЦИЯ ОБНОВЛЕНИЯ ДАННЫХ В TREEVIEW==========================================
    def view_records(self):
        # Добавьте стиль и конфигурацию тега
        self.current_month_index = int((datetime.datetime.now(tz=None)).strftime("%m")) - 1
        self.current_year_index = int((datetime.datetime.now(tz=None)).strftime("%Y"))
        self.month_label.config(text=self.months[(self.current_month_index) % 12])
        self.year_label.config(text=self.current_year_index)
        self.tree.tag_configure("Red.Treeview", foreground="red")
        self.tree.tag_configure("Yellow.Treeview", foreground="#1437FF")
        date_obj = datetime.datetime.now()
        formatted_date = date_obj.strftime('%d.%m.%Y')
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f'''SELECT z.Номер_заявки,
                                       FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                                       w.ФИО AS Диспетчер,
                                       g.город AS Город,
                                       CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                                       тип_лифта,
                                       причина,
                                       m.ФИО,
                                       FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                                       комментарий,
                                       z.id
                                FROM zayavki z
                                JOIN workers w ON z.id_диспетчер = w.id
                                JOIN goroda g ON z.id_город = g.id
                                JOIN street s ON z.id_улица = s.id
                                JOIN doma d ON z.id_дом = d.id
                                JOIN padik p ON z.id_подъезд = p.id
                                JOIN workers m ON z.id_механик = m.id
                                WHERE DATE_FORMAT(FROM_UNIXTIME(z.Дата_заявки), '%m') = ?
                                AND DATE_FORMAT(FROM_UNIXTIME(z.Дата_заявки), '%Y') = ? and z.pc_id = ?
                                order by z.id;''',
                               (f'{str(self.current_month_index + 1).zfill(2)}', f'{str(self.current_year_index)}', self.enabled.get()))
                [self.tree.delete(i) for i in self.tree.get_children()]
                for row in cursor.fetchall():
                    if row[-3] is None and int((datetime.datetime.strptime(row[1], time_format)).timestamp()) < int((time.time()) - 86400):
                        self.tree.insert('', 'end', values=tuple(row), tags=('Red.Treeview',))
                    elif row[-3] is None:
                        self.tree.insert('', 'end', values=tuple(row), tags=('Yellow.Treeview',))
                    elif row[-3] != None and row[-5] == 'Остановлен':
                        self.tree.insert('', 'end', values=tuple(row), tags=('Green.Treeview',))
                    else:
                        self.tree.insert('', 'end', values=tuple(row))
                self.tree.yview_moveto(1.0)
                connection.commit()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.print_info(formatted_date)

    # ===ФУНКЦИЯ КНОПКИ ПОИСКА ПО АДРЕСУ===============================================
    def search_records(self, town, adress, calendar1, calendar2):
        self.town = town
        self.adress = adress
        self.calendar1 = calendar1
        self.calendar2 = calendar2
        time_obj1 = datetime.datetime.strptime(self.calendar1, '%d.%m.%Y')
        time_obj2 = datetime.datetime.strptime(self.calendar2, '%d.%m.%Y')
        unix_time1 = int(time_obj1.timestamp())
        unix_time2 = int(time_obj2.timestamp()) + 86400
        try:
            conn = mariadb.connect(user=user, password=password, host=host, port=port, database=database)
            cur = conn.cursor()
            cur.execute(f'''SELECT z.Номер_заявки,
                           FROM_UNIXTIME(z.Дата_заявки, '%d.%m.%Y, %H:%i') AS Дата_заявки,
                           w.ФИО AS Диспетчер,
                           g.город AS Город,
                           CONCAT(s.улица, ', ', d.номер, ', ', p.номер) AS Адрес,
                           тип_лифта,
                           причина,
                           m.ФИО,
                           FROM_UNIXTIME(дата_запуска, '%d.%m.%Y, %H:%i') AS Дата_запуска,
                           комментарий,
                           z.id
                            FROM zayavki z
                            JOIN workers w ON z.id_диспетчер = w.id
                            JOIN goroda g ON z.id_город = g.id
                            JOIN street s ON z.id_улица = s.id
                            JOIN doma d ON z.id_дом = d.id
                            JOIN padik p ON z.id_подъезд = p.id
                            JOIN workers m ON z.id_механик = m.id
                            WHERE CONCAT(s.улица, ', ', d.номер, ', ', p.номер) LIKE ? 
                            and Дата_заявки BETWEEN "{unix_time1}" and "{unix_time2}"''', (f"%{adress}%",))
            [self.tree.delete(i) for i in self.tree.get_children()]
            for row in cur.fetchall():
                self.tree.insert('', 'end', values=row)
            children = self.tree.get_children()
            if children:
                self.tree.selection_set(children)
                self.tree.focus(children[0])
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        except Exception as e:
            showinfo('Информация', f"Произошла непредвиденная ошибка: {e}")
        else:
            if not self.tree.get_children():
                showinfo('Информация', "Нет заявок")

    def open_search_dialog(self):
        Search(address_)

    # ===ОБНОВЛЕНИЕ СТРОК ФИО И АДРЕСОВ В ЛИСТБОКСАХ============================================
    def obnov(self):
        self.entry3.delete(0, tk.END)
        self.entry4.delete(0, tk.END)
        self.entry7.delete(0, tk.END)
        self.tree.yview_moveto(1)
        self.check_input_address()
        self.check_input_fio()
    # === Парсинг ФИО из списка бд фамилий в listbox ===
    def check_input_fio(self, _event=None):
        value7 = self.entry7.get().lower()
        names = []
        try:
            with closing(
                    mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute("select id, ФИО from workers where Должность = 'Механик' and id_ЛК = 1 order by ФИО")
                self.data_meh = cursor.fetchall()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        for i in self.data_meh:
            names.append(''.join(i['ФИО']))
        if value7 == '':
            self.listbox_values7.set(names)
        else:
            data7 = [item7 for item7 in names if item7.lower().startswith(value7)]
            self.listbox_values7.set(data7)

    # === Вставка выбранного ФИО из парсинга в entrybox ===
    def on_change_selection_fio(self, event):
        selection7 = event.widget.curselection()
        if selection7:
            self.index7 = selection7[0]
            self.data7 = event.widget.get(self.index7)
            for d in self.data_meh:
                if self.data7 == d['ФИО']:
                    self.selected_meh_id = d['id']
            self.entry_text7.set(self.data7)
            self.check_input_fio()

    # ===ПАРСИНГ ТИПА ЛИФТОВ ИЗ СПИСКА ЛИФТОВ В ЛИСТБОКС======================
    def check_input_lifts(self, _event=None):
        selected_address = self.data3
        types = []
        street, house, entrance = selected_address.split(', ')
        try:
            with closing(
                    mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''SELECT lifts.тип_лифта
                                    FROM lifts
                                    JOIN padik ON lifts.id_подъезд = padik.id
                                    JOIN doma ON lifts.id_дом = doma.id
                                    JOIN street ON lifts.id_улица = street.id
                                    JOIN goroda ON lifts.id_город = goroda.id
                                    WHERE goroda.город = "{self.selected_city}" AND street.улица = "{street}"
                                    and doma.номер = "{house}" and padik.номер = "{entrance}" order BY street.улица, doma.`номер`, padik.`номер`''')
                data_lifts = cursor.fetchall()
                for lift in data_lifts:
                    lift_str = f"{lift['тип_лифта']}"
                    types.append(lift_str)
                    self.listbox_type.insert(tk.END, lift_str)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        if self.entry_text4.get() == '':
            self.listbox_values_type.set(types)
            self.entry4.delete(0, tk.END)
        else:
            data2 = [item for item in types if
                    self.entry_text4.get().lower() in item.lower()]
            self.listbox_values_type.set(data2)

    # ===ПАРСИНГ АДРЕСОВ ИЗ СПИСКА АДРЕСОВ В ЛИСТБОКС=========================
    def check_input_address(self, _event=None):
        self.listbox_type.delete(0, tk.END)
        names = []
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''
                    SELECT street.улица, doma.номер as дом, padik.номер as подъезд
                    FROM street 
                    JOIN doma ON street.id = doma.улица_id
                    JOIN padik ON doma.id = padik.дом_id
                    JOIN goroda ON street.город_id = goroda.id
                    WHERE goroda.город = "{self.selected_city}" 
                    order BY street.улица, doma.`номер`, padik.`номер`''')
                data_streets = cursor.fetchall()
                for d in data_streets:
                    address_str = f"{d['улица']}, {d['дом']}, {d['подъезд']}"
                    names.append(address_str)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        if self.entry3.get().lower() == '':
            self.listbox_values.set(names)
            self.entry4.delete(0, tk.END)
        else:
            data = [item for item in names if self.entry3.get().lower() in item.lower()]
            self.listbox_values.set(data)
            self.on_change_selection_address

    # ===ВСТАВКА ЛИФТА ИЗ ПАРСИНГА В ЛИСТБОКС=============================
    def on_change_selection_lift(self, event):
        self.selection = event.widget.curselection()
        if self.selection:
            self.index4 = self.selection[0]
            self.data4 = event.widget.get(self.index4)
            self.entry_text4.set(self.data4)
        self.check_input_lifts()
    # ===ВСТАВКА АДРЕСА ИЗ ПАРСИНГА В ЛИСТБОКС=============================
    def on_change_selection_address(self, event):
        self.selection = event.widget.curselection()
        if self.selection:
            self.index3 = self.selection[0]
            self.data3 = event.widget.get(self.index3)
            self.entry_text3.set(self.data3)
            self.check_input_address()
        self.check_input_lifts()

    def create_combobox(self, combo_text, town):
        frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
        label = ttk.Label(self, text=combo_text)
        label.pack(anchor=NW)
        combo = ttk.Combobox(self, values=town)
        combo.pack(anchor=NW, side=LEFT)
        return frame

    # --------------------------------------------------------------
    def select(self):
        self.type_li = self.entry4.get()
        self.adress_excel = self.entry_text3.get()
        self.ala = self.prich5.get()
        fio_excel = self.entry_text7.get()

    def sql_insert(self):
        selected_disp_str = eval(self.disp.get())
        self.selected_disp_id = selected_disp_str['id']
        self.selected_disp_name = selected_disp_str['ФИО']
        date_ = (datetime.datetime.now(tz=None)).strftime("%d.%m.%Y, %H:%M")
        time_obj = datetime.datetime.strptime(date_, time_format)
        unix_time = int(time_obj.timestamp())
        d = (datetime.datetime.now(tz=None)).strftime("%d")
        val2 = [self.selected_disp_name, self.selected_city, self.entry_text3.get(),
                self.entry_text4.get(), self.prich5.get(), self.entry_text7.get()]
        naz = ['Диспетчер', 'Город', 'Адрес', 'Тип лифта', 'Причина остановки', 'ФИО механика']
        for i in range(len(val2)):
            if len(val2[i]) < 2:
                mb.showerror(
                    "Ошибка",
                    f"Введите данные в строке: {naz[i]}")
                return
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor()
                cursor.execute(f"SELECT COALESCE(MAX(Номер_заявки), 0) FROM {table_zayavki} WHERE DATE_FORMAT(FROM_UNIXTIME("
                    f"Дата_заявки), '%Y-%m') = ?",
                    ((datetime.datetime.now()).strftime('%Y-%m'),))
                res = cursor.fetchone()[0]
                if res is None:
                    self.num_request = 1
                else:
                    self.num_request = res + 1
                #===========================================
                data = self.entry_text3.get()
                parts = data.split(',')
                try:
                    with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                        cursor = connection.cursor()
                        cursor.execute(f'''SELECT goroda.id AS goroda_id, street.id AS street_id, doma.id AS doma_id, padik.id AS padik_id, lifts.id as id_лифт
                                                FROM lifts
                                                JOIN padik ON lifts.id_подъезд = padik.id
                                                JOIN doma ON lifts.id_дом = doma.id
                                                JOIN street ON lifts.id_улица = street.id
                                                JOIN goroda ON lifts.id_город = goroda.id
                                                WHERE goroda.город = "{self.selected_city}" AND street.улица = "{parts[0].strip()}" AND doma.номер = "{parts[1].strip()}" 
                                                AND padik.номер = "{parts[2].strip()}" and тип_лифта="{self.entry_text4.get()}";''')
                        data_lifts = cursor.fetchall()
                except mariadb.Error as e:
                    showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
                print(data_lifts)
                gorod, street, dom, padik, lift_id = data_lifts[0]
                val = (self.num_request, unix_time, self.selected_disp_id,
                       gorod, street, dom, padik, self.entry4.get(),
                    self.prich5.get(), self.selected_meh_id, None, '', lift_id, pc_id)
                try:
                    with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                        cursor = connection.cursor()
                        cursor.execute(f'''INSERT INTO zayavki (
                                        Номер_заявки,
                                        Дата_заявки,
                                        id_Диспетчер,
                                        id_город,
                                        id_улица,
                                        id_дом,
                                        id_подъезд,
                                        тип_лифта,
                                        Причина,
                                        id_Механик,
                                        Дата_запуска,
                                        Комментарий,
                                        id_Лифт,
                                        pc_id) 
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);''',(val))
                        connection.commit()
                except mariadb.Error as e:
                    showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        msg = f"Запись успешно добавлена! Её порядковый номер - {res + 1}"
        mb.showinfo("Информация", msg)
        self.view_records()
        self.obnov()

    # ======ФУНКЦИЯ СПРОСА О ЗАКРЫТИИ ПРОГРАММЫ=================================================
    def on_closing(self):
        result = askyesno(title="Подтвержение действия", message="Закрыть программу?")
        if result:
            root.destroy()
        else:
            showinfo("Результат", "Действие отменено.")

# ====ВЫЗОВ ФУНКЦИЙ КНОПОК РЕДАКТИРОВАНИЯ==================================================================
class Child(tk.Toplevel):
    def __init__(self, rows):
        self.rows = rows
        super().__init__(root)
        self.init_child()
        self.view = app

    def init_child(self):
        self.title('Редактировать')
        self.geometry('500x500')
        self.resizable(False, False)
        self.wm_attributes('-topmost', 1)

        self.bind('<Unmap>', self.on_unmap)

        font10 = tkFont.Font(family='Helvetica', size=10, weight=tkFont.BOLD)
        font12 = tkFont.Font(family='Helvetica', size=12, weight=tkFont.BOLD)

        label_select = tk.Label(self, text='Дата заявки:', font=font12)
        label_select.place(x=20, y=50)
        label_sum = tk.Label(self, text='Диспетчер:', font=font12)
        label_sum.place(x=20, y=80)
        label_select = tk.Label(self, text='Город:', font=font12)
        label_select.place(x=20, y=110)
        label_sum = tk.Label(self, text='Адрес:', font=font12)
        label_sum.place(x=20, y=140)
        label_sum = tk.Label(self, text='Тип лифта:', font=font12)
        label_sum.place(x=20, y=170)
        label_sum = tk.Label(self, text='Причина остановки:', font=font12)
        label_sum.place(x=20, y=200)
        label_sum = tk.Label(self, text='ФИО механика:', font=font12)
        label_sum.place(x=20, y=230)
        label_sum = tk.Label(self, text='Дата запуска:', font=font12)
        label_sum.place(x=20, y=260)
        label_sum = tk.Label(self, text='Комментарий:', font=font12)
        label_sum.place(x=20, y=290)
#===============================================================================================
        self.text_entry_data = tk.StringVar(value=self.rows[0]['Дата_заявки'])
        self.calen1 = tk.Entry(self, textvariable=self.text_entry_data, font=font10)
        self.calen1.place(x=200, y=50)
#================================================================================================
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute("SELECT id, ФИО FROM workers WHERE Должность = 'Диспетчер'")
                read = cursor.fetchall()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        # Создание словаря с соответствием ФИО и ID диспетчеров
        self.fio_to_id = {i['ФИО']: i['id'] for i in read}
        # Получение списка ФИО диспетчеров
        fio_d = [i['ФИО'] for i in read]
        # Вставляем текущего диспетчера в начало списка
        l1 = [j for j in fio_d]
        l1.insert(0, self.rows[0]['Диспетчер'])
        # Создание Combobox с выбором диспетчера
        self.dispetcher = ttk.Combobox(self, values=list(dict.fromkeys(l1)), font=font10)
        self.dispetcher.current(0)
        self.dispetcher.place(x=200, y=80)
#============================================================================================
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute("select id, город from goroda")
                data_towns = cursor.fetchall()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.town_to_id = {i['город']: i['id'] for i in data_towns}
        # Получение списка ФИО диспетчеров
        town_d = [i['город'] for i in data_towns]
        g1 = [j for j in town_d]
        g1.insert(0, self.rows[0]['Город'])
        self.combobox_town = ttk.Combobox(self, values=list(dict.fromkeys(g1)), font=font10)
        self.combobox_town.current(0)
        self.combobox_town.place(x=200, y=110)
        self.combobox_town.bind("<<ComboboxSelected>>", self.on_town_select)
#=============================================================================================
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''SELECT goroda.id as goroda_id, goroda.город, 
                    CONCAT(street.улица, ', ', doma.номер, ', ', padik.номер) as Адрес,
                    street.id as street_id, street.улица as улица,
                    doma.id as doma_id, doma.номер as дом, 
                    padik.id as padik_id, padik.номер as подъезд
                    FROM goroda
                    JOIN street ON goroda.id = street.город_id
                    JOIN doma ON street.id = doma.улица_id
                    JOIN padik ON doma.id = padik.дом_id
                    WHERE goroda.город = "{self.combobox_town.get()}"
                    order by street.улица, doma.номер, padik.номер''')
                adreses = cursor.fetchall()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.street_to_id = {i['улица']: i['street_id'] for i in adreses}
        self.house_to_id = {i['дом']: i['doma_id'] for i in adreses}
        self.padik_to_id = {i['подъезд']: i['padik_id'] for i in adreses}
        adres_list = [i['Адрес'] for i in adreses]
        self.selected_address = tk.StringVar(value=self.rows[0]['Адрес'])
        self.address_combobox = ttk.Combobox(self, textvariable=self.selected_address, font=font10, width=30)
        adres_list.insert(0, self.rows[0]['Адрес'])
        self.address_combobox['values'] = adres_list
        self.address_combobox.place(x=200, y=140)
        self.address_combobox.bind("<<ComboboxSelected>>", self.on_address_select)
        self.street, self.house, self.entrance = self.address_combobox.get().split(', ')

#=============================================================================================
        self.selected_type = tk.StringVar(value=self.rows[0]['тип_лифта'])
        self.combobox_lift = ttk.Combobox(self, textvariable=self.selected_type, font=font10)
        self.street, self.house, self.entrance = self.address_combobox.get().split(', ')
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor(dictionary=True)
                cursor.execute(f'''SELECT lifts.тип_лифта
                                    FROM lifts
                                    JOIN padik ON lifts.id_подъезд = padik.id
                                    JOIN doma ON lifts.id_дом = doma.id
                                    JOIN street ON lifts.id_улица = street.id
                                    JOIN goroda ON lifts.id_город = goroda.id
                                    WHERE goroda.город = "{self.combobox_town.get()}" AND street.улица = "{self.street}"
                                    and doma.номер = "{self.house}" and padik.номер = "{self.entrance}"''')
                data_lifts = cursor.fetchall()
                self.add_type_lifts = [lift['тип_лифта'] for lift in data_lifts]
                # Обновление значений комбобокса с типами лифтов
                self.combobox_lift['values'] = self.add_type_lifts
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.combobox_lift.place(x=200, y=170)
#==============================================================================================
        self.combobox_stop = ttk.Combobox(self, values=list(dict.fromkeys(
                    [self.rows[0]['причина'], 'Неисправность', 'Застревание', 'Остановлен'])), font=font10)
        self.combobox_stop.current(0)
        self.combobox_stop.place(x=200, y=200)
#==============================================================================================
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute("select ФИО, id from workers where Должность = 'Механик' and id_ЛК = 1 order by ФИО")
                read = cursor.fetchall()
        except mariadb.Error as e:
            showerror('Информация', f"Ошибка при работе с базой данных: {e}")
        self.meh_to_id = {i['ФИО']: i['id'] for i in read}
        meh = [i['ФИО'] for i in read]
        m1 = [j for j in meh]
        m1.insert(0, self.rows[0]['ФИО'])
        self.combobox_meh = ttk.Combobox(self, values=list(dict.fromkeys(m1)), font=font10)
        self.combobox_meh.current(0)
        self.combobox_meh.place(x=200, y=230)
#====================================================================================
        self.text_entry_zapusk = tk.StringVar(value=self.rows[0]['дата_запуска'])
        self.calen2 = tk.Entry(self, textvariable=self.text_entry_zapusk, font=font10)
        self.calen2.place(x=200, y=260)
#=======================================================================================
        self.text_entry_comment = tk.StringVar(value=self.rows[0]['комментарий'])
        self.entry_comment = tk.Entry(self, textvariable=self.text_entry_comment, font=font10)
        self.entry_comment.place(x=200, y=290)
        btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        btn_cancel.place(x=300, y=350)
        self.btn_ok = ttk.Button(self, text='Сохранить', command=self.save_and_close)
        self.btn_ok.place(x=200, y=350)

    def on_unmap(self, event):
        self.deiconify()  # Отменяем сворачивание дочернего окна

    def deiconify(self):
        if self.state() == 'iconic':
            self.state('normal')

    def on_town_select(self, event):
        selected_town = self.combobox_town.get()
        # Запрос к базе данных для получения типов лифтов на основе выбранного адреса
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''SELECT CONCAT(street.улица, ', ', CAST(doma.номер AS CHAR), ', ', CAST(padik.номер AS CHAR)) AS Адрес
                    FROM goroda
                    JOIN street ON goroda.id = street.город_id
                    JOIN doma ON street.id = doma.улица_id
                    JOIN padik ON doma.id = padik.дом_id
                    WHERE goroda.город = "{self.combobox_town.get()}"
                    order by street.улица, doma.номер, padik.номер''')
                adreses = cursor.fetchall()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.add_adres = [adres['Адрес'] for adres in adreses]
        # Обновление значений комбобокса с типами лифтов
        self.address_combobox['values'] = self.add_adres
        # Сбрасываем текущее выбранное значение типа лифта
        self.selected_address.set('ВЫБРАТЬ АДРЕС')
        self.selected_type.set('ВЫБРАТЬ ЛИФТ')

    def on_address_select(self, event):
        # Запрос к базе данных для получения типов лифтов на основе выбранного адреса
        self.street, self.house, self.entrance = self.address_combobox.get().split(', ')
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection:
                cursor = connection.cursor(dictionary=True)
                cursor.execute(f'''SELECT lifts.тип_лифта
                                FROM lifts
                                JOIN padik ON lifts.id_подъезд = padik.id
                                JOIN doma ON lifts.id_дом = doma.id
                                JOIN street ON lifts.id_улица = street.id
                                JOIN goroda ON lifts.id_город = goroda.id
                                WHERE goroda.город = "{self.combobox_town.get()}" AND street.улица = "{self.street}"
                                and doma.номер = "{self.house}" and padik.номер = "{self.entrance}"''')
                data_lifts = cursor.fetchall()
                self.add_type_lifts = [lift['тип_лифта'] for lift in data_lifts]
                # Обновление значений комбобокса с типами лифтов
                self.combobox_lift['values'] = self.add_type_lifts
                # Сбрасываем текущее выбранное значение типа лифта
                self.selected_type.set('ВЫБРАТЬ ЛИФТ')
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

    def get_selected_dispetcher_id(self):
        selected_dispetcher_fio = self.dispetcher.get()
        selected_dispetcher_id = self.fio_to_id.get(selected_dispetcher_fio)
        return selected_dispetcher_id

    def get_selected_town_id(self):
        selected_town = self.combobox_town.get()
        selected_town_id = self.town_to_id.get(selected_town)
        return selected_town_id

    def get_selected_adres_id(self):
        if self.address_combobox.get() == 'ВЫБРАТЬ АДРЕС':
            return self.address_combobox.get()
        else:
            street, house, padik = self.address_combobox.get().split(',')
            selected_street_id = self.street_to_id.get(street.strip())
            selected_house_id = self.house_to_id.get(house.strip())
            selected_padik_id = self.padik_to_id.get(padik.strip())
            return selected_street_id, selected_house_id, selected_padik_id

    def get_selected_meh_id(self):
        selected_meh_fio = self.combobox_meh.get()
        selected_meh_id = self.meh_to_id.get(selected_meh_fio)
        return selected_meh_id

    def save_and_close(self):
        adres_id = self.get_selected_adres_id()
        if not adres_id or adres_id == 'ВЫБРАТЬ АДРЕС':
            mb.showerror("Ошибка", "Выберите адрес")
            return

        self.view.update_record(
            self.calen1.get(),
            self.get_selected_dispetcher_id(),
            self.get_selected_town_id(),
            self.get_selected_adres_id()[0],
            self.get_selected_adres_id()[1],
            self.get_selected_adres_id()[2],
            self.combobox_lift.get(),
            self.combobox_stop.get(),
            self.get_selected_meh_id(),
            self.calen2.get(),
            self.entry_comment.get())
        self.destroy()

class Search(tk.Toplevel):
    def __init__(self, address_):
        address_ = address_
        super().__init__()
        self.init_search()
        self.view = app

    def init_search(self):
        self.title('Поиск')
        self.geometry('600x350+400+300')
        self.resizable(False, False)
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute("select id, город from goroda")
                data_towns = cursor.fetchall()
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        # Установить значение по умолчанию
        default_town = data_towns[0] if data_towns else ''
        self.var2 = tk.StringVar(value=default_town)
        # Привязать метод on_select_city к изменению значения переменной
        self.var2.trace("w", self.on_select_city)
        self.label2 = tk.Label(self, text="Город", font='Calibri 14 bold')
        self.label2.pack(side=tk.TOP, anchor=tk.W)
        frame2 = tk.Frame()
        for d in data_towns:
            lang_btn4 = tk.Radiobutton(self, text=d['город'], value=d, variable=self.var2, font='Calibri  13')
            lang_btn4.pack(side=tk.TOP, anchor=tk.NW, padx=0, pady=0, fill='y')
        frame2.pack(side=tk.LEFT, anchor=tk.NW, expand=True)
#================================================================================================================
        label_adres = ttk.Label(self, text='Адрес', font='Calibri 14 bold')
        label_adres.place(x=170, y=0)
        label_data = ttk.Label(self, text='Дата', font='Calibri 14 bold')
        label_data.place(x=420, y=0)
        label_c = ttk.Label(self, text='с', font='Calibri 15 bold')
        label_c.place(x=390, y=40)
        self.calendar1 = DateEntry(self, locale='ru_RU', font=1)
        self.calendar1.bind("<<DateEntrySelected>>")
        self.calendar1.place(x=420, y=40)
        label_po = ttk.Label(self, text='по', font='Calibri 15 bold')
        label_po.place(x=380, y=90)
        self.calendar2 = DateEntry(self, locale='ru_RU', font=1)
        self.calendar2.bind("<<DateEntrySelected>>")
        self.calendar2.place(x=420, y=90)
        self.frame11 = tk.Frame(borderwidth=1)
        self.entry_text11 = tk.StringVar()
        self.entry11 = tk.Entry(self, textvariable=self.entry_text11, width=33)
        self.entry11.bind('<KeyRelease>', self.check_input_11)
        self.entry11.place(x=170, y=40)
        self.listbox_values11 = tk.Variable()
        self.listbox11 = tk.Listbox(self, listvariable=self.listbox_values11, width=25, font='Calibri 12')
        self.listbox11.bind('<<ListboxSelect>>', self.on_change_selection_11)
        self.listbox11.place(x=170, y=90)
        selected_city_str = eval(self.var2.get())
        self.selected_city_id = selected_city_str['id']
        self.selected_city = selected_city_str['город']
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''SELECT goroda.id as goroda_id, goroda.город, 
                    street.id as street_id, street.улица, 
                    doma.id as doma_id, doma.номер as дом, 
                    padik.id as padik_id, padik.номер as подъезд
                    FROM goroda
                    JOIN street ON goroda.id = street.город_id
                    JOIN doma ON street.id = doma.улица_id
                    JOIN padik ON doma.id = padik.дом_id
                    WHERE goroda.город = "{self.selected_city}" order BY street.улица, doma.`номер`, padik.`номер`''')
                self.data_streets = cursor.fetchall()
                for d in self.data_streets:
                    self.address_str = f"{d['улица']}, {d['дом']}, {d['подъезд']}"
                    self.listbox11.insert(tk.END, self.address_str)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        self.frame11.pack(side=tk.LEFT, anchor=tk.NW)
        style = ttk.Style()
        # Создаем стиль для кнопки "Поиск" с зеленым цветом
        style.configure("Green.TButton", foreground="green", background="#50C878", font='Calibri 12')
        # Создаем стиль для кнопки "Закрыть" с красным цветом
        style.configure("Red.TButton", foreground="red", background="#FCA195", font='Calibri 12')
        # Создание и размещение кнопки "Поиск" большого размера слева снизу
        btn_search = ttk.Button(self, text='Поиск', style="Green.TButton", command=self.destroy)
        btn_search.place(x=0, y=300, width=300, height=50)  # Установить координаты, ширину и высоту
        # Создание и размещение кнопки "Закрыть" большого размера справа снизу
        btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy, style="Red.TButton")
        btn_cancel.place(x=300, y=300, width=300, height=50)  # Установить координаты, ширину и высоту
        btn_search.bind('<Button-1>', lambda event: (self.view.search_records(self.var2, self.entry11.get(), self.calendar1.get(), self.calendar2.get())))

    def on_select_city(self, *args):
        selected_city_str = eval(self.var2.get())
        self.selected_city_id = selected_city_str['id']
        self.selected_city = selected_city_str['город']
        # Очистить Listbox перед добавлением новых улиц
        self.listbox11.delete(0, tk.END)
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''SELECT goroda.id as goroda_id, goroda.город, 
                    street.id as street_id, street.улица, 
                    doma.id as doma_id, doma.номер as дом, 
                    padik.id as padik_id, padik.номер as подъезд
                    FROM goroda
                    JOIN street ON goroda.id = street.город_id
                    JOIN doma ON street.id = doma.улица_id
                    JOIN padik ON doma.id = padik.дом_id
                    WHERE goroda.город = "{self.selected_city}" order BY street.улица, doma.`номер`, padik.`номер`''')
                self.data_streets = cursor.fetchall()
                for d in self.data_streets:
                    self.address_str = f"{d['улица']}, {d['дом']}, {d['подъезд']}"
                    self.listbox11.insert(tk.END, self.address_str)
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")

        # =========================================================================
    def check_input_11(self, _event=None):
        selected_city_str = eval(self.var2.get())
        self.selected_city_id = selected_city_str['id']
        self.selected_city = selected_city_str['город']
        value = self.entry_text11.get().lower()
        names = []
        try:
            with closing(mariadb.connect(user=user, password=password, host=host, port=port, database=database)) as connection2:
                cursor = connection2.cursor(dictionary=True)
                cursor.execute(f'''SELECT goroda.id as goroda_id, goroda.город, 
                    street.id as street_id, street.улица, 
                    doma.id as doma_id, doma.номер as дом, 
                    padik.id as padik_id, padik.номер as подъезд
                    FROM goroda
                    JOIN street ON goroda.id = street.город_id
                    JOIN doma ON street.id = doma.улица_id
                    JOIN padik ON doma.id = padik.дом_id
                    WHERE goroda.город = "{self.selected_city}" order BY street.улица, doma.`номер`, padik.`номер`''')
                self.data_streets = cursor.fetchall()
                for d in self.data_streets:
                    self.address_str = f"{d['улица']}, {d['дом']}, {d['подъезд']}"  # Парсим адреса из файла
                    names.append(''.join(self.address_str).strip())
        except mariadb.Error as e:
            showinfo('Информация', f"Ошибка при работе с базой данных: {e}")
        if value == '':  # Если поле ввода пустое, показываем все адреса
            self.listbox_values11.set(names)
        else:  # Если введено что-то, показываем только подходящие по шаблону адреса
            data = [item for item in names if value in item.lower()]
            self.listbox_values11.set(data)

    def on_change_selection_11(self, event):
        self.selection = event.widget.curselection()
        if self.selection:
            self.index3 = self.selection[0]
            self.data3 = event.widget.get(self.index3)
            self.entry_text11.set(self.data3)
            self.check_input_11()


# ==========================================================================================
class Comment(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.init_comm()
        self.view = app

    def init_comm(self):
        self.title('Комментировать')
        self.geometry('350x150+400+300')
        self.resizable(False, False)
        label_comm = ttk.Label(self, text='Комментировать:')
        label_comm.place(x=1, y=10)
        self.frame123 = tk.Frame(borderwidth=1)
        self.frame123.pack(side=tk.TOP)
        self.entry_text123 = tk.StringVar()
        self.entry123 = tk.Entry(self, textvariable=self.entry_text123, width=33)
        self.entry123.place(x=110, y=10)
        btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        btn_cancel.place(x=235, y=50)
        btn_search = ttk.Button(self, text='Комментировать', command=self.save_and_close)
        btn_search.place(x=110, y=50)
        self.entry123.bind("<Button-3>", self.show_menu)
        self.entry123.bind_all("<Control-v>", self.paste_text)

    def save_and_close(self):
        self.view.comment(self.entry123.get())
        self.destroy()

    def paste_text(self, event=None):
        text_to_paste = self.clipboard_get()
        self.entry123.insert(tk.INSERT, text_to_paste)

    def show_menu(self, event):
        # Создаем меню
        menu = tk.Menu(self, tearoff=False, font=20)
        menu.add_command(label="Вставить", command=self.paste_text)
        # Показываем меню в указанной позиции
        menu.post(event.x_root, event.y_root)

class Excel():
    def __init__(self, df):
        self.df = df
        self.open_ex()

    def open_ex(self):
        # Экспортировать данные в Excel
        self.df.to_excel('данные.xlsx', index=False)
        process = subprocess.Popen("данные.xlsx", shell=True)
        process.communicate()


if __name__ == "__main__":
    with open('config.json', 'r') as file:
        data = json.loads(file.read())

    pc_id = data['pc_id']
    host = data['db_host']
    user = data['db_user']
    password = data['db_password']
    port = data['db_port']
    database = data['db_name']
    table_goroda = data['table_goroda']
    table_street = data['table_street']
    table_doma = data['table_doma']
    table_padik = data['table_padik']
    table_lifts = data['table_lifts']
    table_lifts_company = data['table_lifts_company']
    table_uk = data['table_uk']
    table_zayavki = data['table_zayavki']
    table_workers = data['table_workers']

    fio_ = 'fio_meh.csv'
    address_ = 'adreses.csv'
    fio_dispetchers = 'fio_dispetchers.csv'
    goroda = 'goroda.csv'
    time_format = "%d.%m.%Y, %H:%M"

    root = tk.Tk()
    root.resizable(False, True)
    app = Main(root)
    app.pack()
    root.title("МиТОЛ")
    root.state("zoomed")
    root.iconphoto(False, tk.PhotoImage(file='mitol.png'))
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

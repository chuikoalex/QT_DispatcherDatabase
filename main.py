# -*- coding: utf-8 -*-
import os.path
import sys
import datetime
import time
import sqlite3
import logging
import xlsxwriter
from xlsxwriter.exceptions import FileCreateError
from pyexcel_ods import save_data
from collections import OrderedDict

from shutil import copyfile

from PyQt5.QtSql import QSqlDatabase, QSqlRelation, QSqlRelationalTableModel, QSqlTableModel
from PyQt5.QtWidgets import QApplication, QMainWindow, QAbstractItemView
from PyQt5.QtWidgets import QDialog, QMessageBox, QInputDialog, QColorDialog, QFileDialog
from PyQt5.QtGui import QFont, QTextCursor, QColor
from PyQt5.QtCore import QTimer, Qt, pyqtSlot

import start_ui_dialog
import ui_dispatcher

START = False
DISPATCHER = 'Администратор'

logging.basicConfig(
    level=logging.DEBUG,
    filename="log.txt",
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s"
)


def file_creation_date(filename_and_path):
    """ Функция получения даты создания файла

     Функция возвращает дату файла 'filename_and_path'
     в виде строки (dd/mm/yyyy)"""

    created = os.path.getctime(filename_and_path)
    year, month, day = time.localtime(created)[:3]
    return "(%02d/%02d/%d)" % (day, month, year)


def fill_list(path_db, flag='dispatchers'):
    """Функция загрузки списков из БД

    Функция загружает списки диспетчеров (всех/активных) и адресов
    flag=dispatchers / all_dispatchers / address
    Используется в выпадающих списках Панели фильтра и Панели ввода данных"""

    try:
        connection = sqlite3.connect(path_db)
        if flag == 'address':
            result_list = connection.cursor().execute("SELECT building FROM address ORDER BY building").fetchall()
        elif flag == 'all_dispatcher':
            result_list = connection.cursor().execute("SELECT name FROM dispatcher ORDER BY name").fetchall()
        else:
            result_list = connection.cursor().execute("SELECT name FROM dispatcher"
                                                      " WHERE active='да' OR active='Да' OR active='ДА'"
                                                      " ORDER BY name").fetchall()
    except sqlite3.OperationalError:
        logging.exception("Exception occurred")
        return False
    else:
        connection.close()
        return result_list


def id_dispatchers_or_address(path_db, field, flag='dispatchers'):
    """Функция получения id из связанных таблиц по имени или адресу

    Функция получает аргументы адрес или имя диспетчера и возвращает соотв id из БД
    flag=dispatchers / address
    Используется в фильтрации"""

    if field != "ВСЕ":
        try:
            connection = sqlite3.connect(path_db)
            if flag == 'address':
                result = connection.cursor().execute(f"SELECT id FROM address WHERE building='{field}'").fetchall()
            else:
                result = connection.cursor().execute(f"SELECT id FROM dispatcher WHERE name='{field}'").fetchall()
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
        else:
            connection.close()
            return result[0][0]
    return 0


def dispatchers_or_address_from_id(path_db, field_id, flag='dispatchers'):
    """Функция получения диспетчера или адреса из связанных таблиц по id

    Функция получает аргументы id и возвращает адрес или имя диспетчера из БД
    flag=dispatchers / address
    Используется в загрузке отчетов"""
    try:
        connection = sqlite3.connect(path_db)
        if flag == 'address':
            result = connection.cursor().execute(f"SELECT building FROM address WHERE id={field_id}").fetchall()
        else:
            result = connection.cursor().execute(f"SELECT name FROM dispatcher WHERE id={field_id}").fetchall()
    except sqlite3.OperationalError:
        logging.exception("Exception occurred")
        return 'ВСЕ'
    else:
        connection.close()
        return result[0][0]


def text_clearing_characters(text, type_text='text'):
    """Функция удаления лишних символов из текста

    Функция используется при формировании данных для внесения в БД. Имя диспетчера, адрес, название отчета"""

    clearing_txt = ''
    allowed_characters = '.-/ ' if type_text != 'text' else '№.,-+/*=?() '
    for symbol in text:
        if symbol.isalpha() or symbol.isdigit() or (symbol in allowed_characters):
            clearing_txt += symbol
    return clearing_txt


class StartWindow(QDialog, start_ui_dialog.Ui_Dialog):
    """Стартовое окно программы.

    Стартовое окно - объект проверяет наличие файла базы и архивов базы.
    Если база в наличии, проводится запрос активных диспетчеров, для выбора пользователя.
    При ошибках производится запись проблемы в файл 'log.txt'. """

    def __init__(self, path_db):
        super().__init__()
        self.setupUi(self)

        if os.path.isfile(path_db):
            self.progressBar.setProperty("value", 20)
            file_data = f"Файл базы данных - файл найден {file_creation_date(path_db)}"
            self.label_db0.setText(file_data)
            list_dispatcher_name = fill_list(path_db, 'dispatchers')
            if not list_dispatcher_name:
                self.label_status.setText("ОШИБКА работы базы данных! Обратитесь к администратору.")
            else:
                self.progressBar.setProperty("value", 40)

                all_archives_ok = ''
                for i in range(1, 4):
                    path_archive_db = f'Archive_db/archive_db{i}.db'
                    if os.path.isfile(path_archive_db):
                        command = f'self.label_db{i}.setText("Архив базы {i}:Ок {file_creation_date(path_archive_db)}")'
                    else:
                        command = f'self.label_db{i}.setText("Архив базы {i}: no")'
                        all_archives_ok = 'Не забывайте создавать архивы базы.'
                    eval(command)
                    self.progressBar.setProperty("value", 40 + 20 * i)
                self.label_status.setText("База данных готова к работе. " + all_archives_ok)

                self.comboBox_user.addItems([name[0] for name in list_dispatcher_name])
                self.pushButton.clicked.connect(self.start_db)

        else:
            self.label_db0.setText("Файл базы данных - файл не найден!")
            self.label_status.setText("База данных не найдена! Обратитесь к администратору.")
            logging.warning("Файл базы данных не найден.")

    def start_db(self):
        global START
        global DISPATCHER
        START = True
        DISPATCHER = self.comboBox_user.currentText()
        self.close()


class MainWindow(QMainWindow, ui_dispatcher.Ui_MainWindow):
    """Основное окно программы"""

    def __init__(self, path_db):
        super().__init__()
        self.path_db = path_db
        self.setupUi(self)

        # Загружаем настройки диспетчера в виде словаря
        self.setting_dict = self.setting_dispatcher_load()

        # Запускаем таймер (для отображения часов в программе)
        self.timer = QTimer()
        self.timer.start(15000)

        # Таймер для одноразовых срабатываний
        self.single_timer = QTimer()

        # Обновляем информацию в верхней части
        self.create_head_panel()

        # Подключаем базу данных
        self.db = QSqlDatabase.addDatabase('QSQLITE')
        self.db.setDatabaseName(self.path_db)

        # Получаем списки Диспетчеров и Адресов
        self.all_dispatcher = fill_list(self.path_db, 'all_dispatcher')
        self.dispatcher = fill_list(self.path_db, 'dispatcher')
        self.address = fill_list(self.path_db, 'address')

        # Создаем панель фильтрации
        self.create_filter_panel()

        # Создаем модели БД для связи
        self.dbmodel = QSqlRelationalTableModel(self)
        self.dbmodel_report = QSqlTableModel(self)
        self.dbmodel_dispatcher = QSqlTableModel(self)
        self.dbmodel_address = QSqlTableModel(self)
        self.dbmodel_pattern = QSqlTableModel(self)

        # Строим (настраиваем) модели
        self.create_dbmodel()
        self.create_dbmodel_report()
        self.create_dbmodel_dispatcher()
        self.create_dbmodel_address()
        self.create_dbmodel_pattern()

        # Связываем модели и таблицы для их отображения
        self.table_database.setModel(self.dbmodel)
        self.report_table.setModel(self.dbmodel_report)
        self.conf_dispatcher_table.setModel(self.dbmodel_dispatcher)
        self.conf_address_table.setModel(self.dbmodel_address)
        self.conf_pattern_table.setModel(self.dbmodel_pattern)

        # Фильтруем данные в модели по умолчанию (здесь для первого вывода таблицы при старте)
        self.filtering_dbmodel(self.set_filter_default())

        # Настраиваем отображение таблицы с данными
        self.create_table_database()

        # Создаем панель ввода данных
        self.create_data_input_panel()

        # Создаем панель отчетов
        self.create_table_report()

        # Создаем панель настройки
        self.create_table_conf_dispatcher()
        self.create_table_conf_address()
        self.create_table_conf_pattern()
        self.create_archive_panel()

        # загружаем информацию в поле логирования
        self.create_logging_box()

        # Обработчики событий:
        self.btn_filter.clicked.connect(self.start_filtering)
        self.btn_cancel_filter.clicked.connect(self.cancel_filtering)

        self.table_database.clicked.connect(self.selected_row_click)
        self.table_database.doubleClicked.connect(self.selected_row_doubleclick)
        self.data_slider_horizontal_size.valueChanged.connect(self.table_database_horizontal_size)
        self.data_slider_font.valueChanged.connect(self.table_database_font_size)

        self.widget_4_working_area.currentChanged.connect(self.step_working_area)

        self.btn_datetime_now.clicked.connect(self.datetime_in_data_input_panel)
        self.btn_pattern_messages.clicked.connect(self.add_pattern_messages)
        self.btn_add_data_db.clicked.connect(self.add_data_input)
        self.btn_cancel_data_db.clicked.connect(self.cansel_data_input)

        self.btn_report_excel.clicked.connect(self.report_excel)
        self.btn_report_openoffice.clicked.connect(self.report_openoffice)
        self.btn_report_load.clicked.connect(self.report_load)
        self.report_table.doubleClicked.connect(self.report_load)
        self.btn_report_save.clicked.connect(self.report_save)
        self.btn_report_del.clicked.connect(self.report_del)

        self.btn_conf_dispatcher_edit.clicked.connect(self.conf_dispatcher_edit)
        self.btn_conf_dispatcher_cansel.clicked.connect(self.conf_dispatcher_cancel)
        self.btn_conf_dispatcher_new.clicked.connect(self.conf_dispatcher_new)

        self.btn_conf_address_edit.clicked.connect(self.conf_address_edit)
        self.btn_conf_address_cansel.clicked.connect(self.conf_address_cancel)
        self.btn_conf_address_new.clicked.connect(self.conf_address_new)
        self.btn_conf_address_streets.clicked.connect(self.conf_select_street)

        self.btn_conf_pattern_edit.clicked.connect(self.conf_pattern_edit)
        self.btn_conf_pattern_cancel.clicked.connect(self.conf_pattern_cancel)
        self.btn_conf_pattern_del.clicked.connect(self.conf_pattern_del)
        self.btn_conf_pattern_new.clicked.connect(self.conf_pattern_new)

        self.btn_archive_database.clicked.connect(self.archive_database)
        self.btn_conf_setting_color.clicked.connect(self.setting_color_dispatcher)
        self.btn_conf_setting_save.clicked.connect(self.setting_dispatcher_save)

        self.logging_slider_font.valueChanged.connect(self.logging_font_size)

        self.timer.timeout.connect(self.create_head_panel)

    # -------- загрузка настроек --------------------------
    def setting_dispatcher_load(self, ):
        """Функция загружает настройки активного диспетчера"""

        id_setting = id_dispatchers_or_address(self.path_db, DISPATCHER)
        try:
            connection = sqlite3.connect(self.path_db)
            result_list = connection.cursor().execute(f"SELECT * FROM setting "
                                                      f"WHERE id_dispatcher={id_setting}").fetchall()
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            return False
        else:
            connection.close()
        tmp_setting = {
            "id": result_list[0][0],
            "id_dispatcher": result_list[0][1],
            "color_text": "#" + result_list[0][2]
        }
        return tmp_setting

    # -------- верхняя панель ---------------------
    @pyqtSlot()
    def create_head_panel(self):
        """Функция отображения информации в верхней строке"""

        dt = datetime.datetime.now()
        self.label_date_now.setText(dt.strftime('%d.%m.%Y года'))
        self.label_time_now.setText(dt.strftime('%H:%M'))
        name_style = "QLabel {font: 87 11pt 'Arial Black'; color:" + self.setting_dict["color_text"] + "}"
        self.label_dispatcher_now.setStyleSheet(name_style)
        self.label_dispatcher_now.setText(DISPATCHER)

    # --------- модели БД -------------------------
    def create_dbmodel(self):
        """Функция построения основной модели базы данных (таблица с заявками)"""

        self.db.open()

        self.dbmodel.setTable('application')

        self.dbmodel.select()
        while self.dbmodel.canFetchMore():
            self.dbmodel.fetchMore()

        self.dbmodel.setRelation(3, QSqlRelation('address', 'id', 'building'))
        self.dbmodel.setRelation(11, QSqlRelation('dispatcher', 'id', 'name'))
        self.dbmodel.setSort(1, Qt.AscendingOrder)

        columns = ['№ п/п',
                   'Дата',
                   'Время',
                   'Адрес',
                   'Пар-я',
                   'Этаж',
                   'Кв.',
                   'Заявитель',
                   'Телефон',
                   'Сообщение',
                   'Примечание',
                   'Диспетчер',
                   'Передано',
                   'Исполнение']
        for i in range(14):
            self.dbmodel.setHeaderData(i, Qt.Horizontal, columns[i])
        self.db.close()

    def create_dbmodel_report(self):
        self.db.open()
        self.dbmodel_report.setTable('report')
        self.dbmodel_report.setSort(1, Qt.AscendingOrder)
        self.dbmodel_report.select()
        while self.dbmodel_report.canFetchMore():
            self.dbmodel_report.fetchMore()
        columns = ['№ п/п',
                   'Дата',
                   'Наименование']
        for i in range(3):
            self.dbmodel_report.setHeaderData(i, Qt.Horizontal, columns[i])
        self.db.close()

    def create_dbmodel_dispatcher(self):
        """Функция построения модели базы данных. Таблица с диспетчерами"""

        self.db.open()
        self.dbmodel_dispatcher.setTable('dispatcher')
        self.dbmodel_dispatcher.setSort(1, Qt.AscendingOrder)
        self.dbmodel_dispatcher.select()
        while self.dbmodel_dispatcher.canFetchMore():
            self.dbmodel_dispatcher.fetchMore()
        columns = ['№ п/п',
                   'ФИО',
                   'Актив.']
        for i in range(3):
            self.dbmodel_dispatcher.setHeaderData(i, Qt.Horizontal, columns[i])
        self.dbmodel_dispatcher.setEditStrategy(QSqlTableModel.OnManualSubmit)
        self.db.close()

    def create_dbmodel_address(self):
        """Функция построения модели базы данных. Таблица с адресами"""

        self.db.open()
        self.dbmodel_address.setTable('address')
        self.dbmodel_address.setSort(1, Qt.AscendingOrder)
        self.dbmodel_address.setFilter("building<>'' ")
        self.dbmodel_address.select()
        while self.dbmodel_address.canFetchMore():
            self.dbmodel_address.fetchMore()
        self.dbmodel_address.setHeaderData(1, Qt.Horizontal, 'Адрес')
        self.dbmodel_address.setEditStrategy(QSqlTableModel.OnManualSubmit)
        self.db.close()

    def create_dbmodel_pattern(self):
        """Функция построения модели базы данных. Таблица с шаблонами обращений"""

        self.db.open()
        self.dbmodel_pattern.setTable('pattern_message')
        self.dbmodel_pattern.select()
        self.dbmodel_pattern.setHeaderData(1, Qt.Horizontal, 'Шаблон сообщения')
        self.dbmodel_pattern.setEditStrategy(QSqlTableModel.OnManualSubmit)
        self.db.close()

    def filtering_dbmodel(self, data_filter):
        """Функция фильтрации модели по полученным данным 'data_filter'

        data_filter(дата_начало, дата_конец, адрес, имя_заявителя, тел_заявителя, диспетчер, не_исполнение_заявки)"""

        self.db.open()

        date_start = data_filter[0]
        date_finish = data_filter[1]
        address = "<>''" if data_filter[2] == 0 else f"='{data_filter[2]}'"
        consumer_name = f"%%" if data_filter[3] == 'ВСЕ' else f"%{data_filter[3]}%"
        consumer_phone = f"%%" if data_filter[4] == 'ВСЕ' else f"%{data_filter[4]}%"
        dispatcher = "<>''" if data_filter[5] == 0 else f"='{data_filter[5]}'"
        execution = "execution is NULL OR execution=''" if data_filter[6] else "TRUE"

        self.dbmodel.setFilter(f"(date BETWEEN '{date_start}' AND '{date_finish}')"
                               f" AND (address{address})"
                               f" AND (consumer_name LIKE '{consumer_name}')"
                               f" AND (consumer_phone LIKE '{consumer_phone}' OR consumer_phone is NULL)"
                               f" AND (id_dispatcher{dispatcher})"
                               f" AND ({execution})"
                               )
        while self.dbmodel.canFetchMore():
            self.dbmodel.fetchMore()

        self.db.close()
        self.message_logging(f"Фильтру соответствует {self.dbmodel.rowCount()} строк данных. ", 'info', False)

    # ---------- основная таблица с БД 'Обращения' --------------
    def create_table_database(self):
        """Функция настройки главной таблицы для вывода БД 'обращения'"""

        self.table_database.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_database.setSelectionMode(1)
        self.table_database.setSelectionBehavior(1)
        self.table_database.setTabKeyNavigation(False)
        self.table_database.setAlternatingRowColors(True)
        self.table_database.verticalHeader().hide()
        self.table_database.horizontalHeader().setSectionResizeMode(9, 1)
        self.table_database.horizontalHeader().setSectionResizeMode(13, 1)
        self.table_database.setColumnWidth(0, 40)
        self.table_database.setColumnWidth(1, 80)
        self.table_database.setColumnWidth(2, 50)
        self.table_database.setColumnWidth(3, 170)
        self.table_database.setColumnWidth(4, 20)
        self.table_database.setColumnWidth(5, 20)
        self.table_database.setColumnWidth(6, 30)
        self.table_database.setColumnWidth(7, 120)
        self.table_database.setColumnWidth(8, 90)
        self.table_database.setColumnWidth(11, 100)
        self.table_database.scrollToBottom()

        self.table_database_horizontal_size()
        self.table_database_font_size()

    @pyqtSlot()
    def table_database_horizontal_size(self):
        """Функция установки высоты строк в основной таблице"""

        horizontal_size = 15 + self.data_slider_horizontal_size.value() * 5
        self.table_database.verticalHeader().setDefaultSectionSize(horizontal_size)

    @pyqtSlot()
    def table_database_font_size(self):
        """Функция изменяет размер шрифта в основной таблице"""

        self.table_database.setFont(QFont("Arial", self.data_slider_font.value()))

    @pyqtSlot()
    def selected_row_click(self):
        """Функция обработки одного клика мышью по строке основной таблицы"""

        row = self.table_database.currentIndex().row()
        self.table_database_horizontal_size()
        self.table_database.setRowHeight(row, 90)

    @pyqtSlot()
    def selected_row_doubleclick(self):
        """Функция обработки двойного клика мышью по строке основной таблицы

        Функция берет данные из выделенной строки в таблице и размещает их в панели ввода данных"""

        data_update = [data.data() for data in self.table_database.selectedIndexes()]

        self.data_id_number.setText(str(data_update[0]))
        self.data_date.setDate(datetime.datetime.strptime(data_update[1], "%Y-%m-%d"))
        self.data_time.setTime(datetime.datetime.strptime(data_update[2], "%H:%M").time())
        self.data_address.setCurrentText(data_update[3])
        self.data_door.setText(data_update[4])
        self.data_floor.setText(data_update[5])
        self.data_flat.setText(data_update[6])
        self.data_consumer_name.setText(data_update[7])
        self.data_consumer_phone.setText(data_update[8])
        self.data_messages.setText(data_update[9])
        self.data_source.setText(data_update[10])
        self.data_dispatcher_name.setCurrentText(data_update[11])
        self.data_time_start.setText(data_update[12])
        self.data_execution.setText(data_update[13])

        self.btn_add_data_db.setText("&Изменить")

    # --------- фильтрация данных ------------------
    def create_filter_panel(self):
        """Функция построения панели фильтра

        добавляет диспетчеров и адреса в выпадающие списки панели фильтра"""
        self.filter_address.addItems(['ВСЕ'])
        self.filter_address.addItems([name[0] for name in self.address])
        self.create_filter_combobox_dispatcher()

    def create_filter_combobox_dispatcher(self, flag='active'):
        """Функция добавляет всех или активных диспетчеров в выпадающий список в панели фильтра"""

        self.filter_dispatcher_name.clear()
        self.filter_dispatcher_name.addItems(['ВСЕ'])
        if flag == 'all':
            self.filter_dispatcher_name.addItems([name[0] for name in self.all_dispatcher])
        else:
            self.filter_dispatcher_name.addItems([name[0] for name in self.dispatcher])

    def set_filter_default(self):
        """Функция настройки значений фильтра по умолчанию (используется и для сброса панели фильтра)"""

        dt = datetime.datetime.now()
        dt_start = dt - datetime.timedelta(days=30)
        self.filter_date_start.setDate(dt_start.date())
        self.filter_date_finish.setDate(dt.date())
        self.filter_address.setCurrentIndex(0)
        self.filter_consumer_name.setText('ВСЕ')
        self.filter_consumer_phone.setText('ВСЕ')
        self.filter_dispatcher_name.setCurrentIndex(0)
        self.filter_no_close.setChecked(False)
        filter_default = (dt_start.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d'), 0, 'ВСЕ', 'ВСЕ', 0, False)
        return filter_default

    @pyqtSlot()
    def start_filtering(self):
        """Функция запуска фильтра по запросу пользователя

        Функция подготавливает кортеж с набором данных из Панели фильтра и
        передает его в функцию фильтрации основной модели БД.
        А также возвращает этот кортеж (используется для сохранения подготовленного отчета)"""

        dt_start = self.filter_date_start.date().toPyDate().strftime('%Y-%m-%d')
        dt = self.filter_date_finish.date().toPyDate().strftime('%Y-%m-%d')
        if (self.filter_address.currentText(),) not in self.address:
            address = 0
            self.filter_address.setCurrentIndex(0)
        else:
            address = id_dispatchers_or_address(self.path_db, self.filter_address.currentText(), 'address')
        consumer_name = self.filter_consumer_name.text()
        consumer_phone = self.filter_consumer_phone.text()
        dispatcher = id_dispatchers_or_address(self.path_db, self.filter_dispatcher_name.currentText(), 'dispatchers')
        flag = self.filter_no_close.isChecked()

        filtering = (dt_start, dt, address, consumer_name, consumer_phone, dispatcher, flag)

        self.filtering_dbmodel(filtering)
        self.create_table_database()
        return filtering

    @pyqtSlot()
    def cancel_filtering(self):
        """Функция сброса фильтра на значения по умолчанию"""

        self.filtering_dbmodel(self.set_filter_default())
        self.create_table_database()

    # --------- рабочая зона с закладками ----------------
    def step_working_area(self):
        """Функция переключает настройки рабочей зоны в зависимости от активной закладки

        0 - <Редактирование данных> (в фильтре список только активных диспетчеров)
        1 - <Отчеты> ()
        2 - <Настройки> ()
        3 - <Просмотр> (уменьшает высоту рабочей области для удобного просмотра таблицы)"""

        index_tab = self.widget_4_working_area.currentIndex()
        if index_tab == 0:
            self.create_filter_combobox_dispatcher()
        else:
            self.create_filter_combobox_dispatcher(flag='all')

        if index_tab == 3:
            self.widget_4_working_area.setMaximumHeight(60)
        else:
            self.widget_4_working_area.setMaximumHeight(270)

    # --------- панель ввода данных в БД 'Обращения' -----
    def create_data_input_panel(self):
        """Функция построения панели ввода данных в БД"""

        global DISPATCHER
        self.data_dispatcher_name.addItems([DISPATCHER])
        self.data_dispatcher_name.addItems([name[0] for name in self.dispatcher if name[0] != DISPATCHER])
        self.data_address.addItems([name[0] for name in self.address])
        self.datetime_in_data_input_panel()

    @pyqtSlot()
    def add_pattern_messages(self):
        """Функция добавления шаблона обращения заявителя в поле 'Обращение'"""

        try:
            connection = sqlite3.connect(self.path_db)
            pattern_list = connection.cursor().execute("SELECT pattern"
                                                       " FROM pattern_message"
                                                       " ORDER BY pattern").fetchall()
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging('Ошибка работы с БД! Обратитесь к администратору.', 'error', False)
            return False
        else:
            connection.close()
        pattern_list = [data[0] for data in pattern_list]
        pattern_messages = QInputDialog.getItem(self,
                                                'Шаблон обращения',
                                                'Выберите шаблон для его вставки в поле',
                                                pattern_list,
                                                editable=True)
        if pattern_messages[1]:
            self.data_messages.insertPlainText(pattern_messages[0] + ' ')
            return True
        return False

    @pyqtSlot()
    def datetime_in_data_input_panel(self):
        """Функция установки текущего времени в панель ввода данных"""

        dt = datetime.datetime.now()
        self.data_date.setDate(dt.date())
        self.data_time.setTime(dt.time())

    @pyqtSlot()
    def cansel_data_input(self):
        """Функция очистки панели ввода данных в БД"""

        self.data_dispatcher_name.setCurrentIndex(0)
        self.datetime_in_data_input_panel()
        self.data_address.setCurrentIndex(0)
        self.data_door.clear()
        self.data_floor.clear()
        self.data_flat.clear()
        self.data_consumer_name.clear()
        self.data_consumer_phone.clear()
        self.data_messages.clear()
        self.data_source.clear()
        self.data_time_start.clear()
        self.data_execution.clear()
        self.btn_add_data_db.setText("&Добавить")

    @pyqtSlot()
    def add_data_input(self):
        """Функция записи данных из панели редактирования в БД 'Обращения'

        Функция записывает новые данные / обновляет старые ('Добавить' / 'Изменить')"""

        input_data = self.create_data_list() if self.create_data_list() else []
        if input_data:
            try:
                connection = sqlite3.connect(self.path_db)
            except sqlite3.OperationalError:
                logging.exception("Exception occurred")
                self.message_logging(f"ОШИБКА БД! Повторите попытку.", 'error', False)
            else:
                if self.btn_add_data_db.text() == "&Добавить":
                    try:
                        connection.cursor().execute(
                            f"INSERT INTO application(date, time, address, door, floor, flat, "
                            f"consumer_name, consumer_phone, messages, source, id_dispatcher, "
                            f"time_start, execution) "
                            f"VALUES('{input_data[0]}', "  # дата
                            f"'{input_data[1]}', "  # время
                            f"'{input_data[2]}', "  # id адреса
                            f"'{input_data[3]}', "  # подъезд 
                            f"'{input_data[4]}', "  # этаж
                            f"'{input_data[5]}', "  # квартира
                            f"'{input_data[6]}', "  # имя заявителя
                            f"'{input_data[7]}', "  # телефон заявителя
                            f"'{input_data[8]}', "  # сообщение-заявка
                            f"'{input_data[9]}', "  # примечание
                            f"'{input_data[10]}', "  # id диспетчера
                            f"'{input_data[11]}', "  # время передачи
                            f"'{input_data[12]}'"  # исполнение
                            f");")
                        connection.commit()
                        connection.close()
                    except sqlite3.OperationalError:
                        logging.exception("Exception occurred")
                        self.message_logging(f"Ошибка записи данных! Повторите попытку.", 'error', False)
                    finally:
                        self.message_logging(
                            f"Новая запись в БД. "
                            f"Дата: {input_data[0]} {input_data[1]}, "
                            f"Заявитель ({input_data[6]}),"
                            f"Обращение ({input_data[8][:15]}...)", 'info')
                        self.cansel_data_input()
                        self.cancel_filtering()

                elif self.btn_add_data_db.text() == "&Изменить":
                    try:
                        connection.cursor().execute(
                            f"UPDATE application SET "
                            f"date='{input_data[0]}', "  # дата
                            f"time='{input_data[1]}', "  # время
                            f"address='{input_data[2]}', "  # id адреса
                            f"door='{input_data[3]}', "  # подъезд 
                            f"floor='{input_data[4]}', "  # этаж
                            f"flat='{input_data[5]}', "  # квартира
                            f"consumer_name='{input_data[6]}', "  # имя заявителя
                            f"consumer_phone='{input_data[7]}', "  # телефон заявителя
                            f"messages='{input_data[8]}', "  # сообщение-заявка
                            f"source='{input_data[9]}', "  # примечание
                            f"id_dispatcher='{input_data[10]}', "  # id диспетчера
                            f"time_start='{input_data[11]}', "  # время передачи
                            f"execution='{input_data[12]}' "  # исполнение
                            f"WHERE id='{input_data[13]}';")  # id для изменения
                        connection.commit()
                        connection.close()
                    except sqlite3.OperationalError:
                        logging.exception("Exception occurred")
                        self.message_logging(f"Ошибка внесения изменений! Повторите попытку.", 'error', False)
                    finally:
                        self.message_logging(
                            f"Изменения внесены в id {input_data[13]} - "
                            f"Заявитель ({input_data[6]}),"
                            f"Обращение ({input_data[8][:15]}...)", 'info')
                        self.cansel_data_input()
                        self.start_filtering()
        return False

    def create_data_list(self):
        """Функция подготовки данных для отправки в БД 'Обращения' """

        data_list = [self.data_date.date().toPyDate().strftime('%Y-%m-%d'),
                     self.data_time.time().toPyTime().strftime('%H:%M')]

        address_building = self.data_address.currentText()
        if (self.data_address.currentText(),) not in self.address:
            self.data_address.setCurrentIndex(0)
            self.message_logging(f'Адрес {address_building} не существует!')
            return False
        else:
            data_list.append(id_dispatchers_or_address(self.path_db, self.data_address.currentText(), 'address'))

        data_list.append(self.data_door.text())
        data_list.append(self.data_floor.text())
        data_list.append(self.data_flat.text())

        if len(self.data_consumer_name.text()) < 2:
            self.message_logging('Имя заявителя отсутствует или слишком короткое!')
            return False
        else:
            data_list.append(self.data_consumer_name.text())

        if self.data_consumer_phone.text() != '':
            if not self.data_consumer_phone.text().isdigit():
                self.message_logging('Телефон заявителя должен содержать только числа!')
                return False
            elif len(self.data_consumer_phone.text()) > 11 or len(self.data_consumer_phone.text()) < 7:
                self.message_logging('Длина номера телефона не может быть меньше 7 цифр или больше 11')
                return False
        data_list.append(self.data_consumer_phone.text())

        if len(self.data_messages.toPlainText()) < 2:
            self.message_logging('Сообщение заявка отсутствует или текст слишком короткий!')
            return False
        else:
            data_list.append(self.data_messages.toPlainText())

        data_list.append(self.data_source.text())
        data_list.append(id_dispatchers_or_address(self.path_db, self.data_dispatcher_name.currentText()))
        data_list.append(self.data_time_start.text())
        data_list.append(self.data_execution.toPlainText())

        data_list.append(self.data_id_number.text())
        for i in range(len(data_list)):
            if isinstance(data_list[i], str):
                data_list[i] = data_list[i].replace('"', '=')
                data_list[i] = data_list[i].replace("'", "")
        return data_list

    # --------- панель отчеты ---------------------------
    def create_table_report(self):
        """Функция настройки таблицы для вывода БД 'отчеты' """

        self.report_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.report_table.setSelectionMode(1)
        self.report_table.setSelectionBehavior(1)
        self.report_table.setTabKeyNavigation(False)
        self.report_table.verticalHeader().hide()
        self.report_table.setColumnWidth(0, 50)
        self.report_table.setColumnWidth(1, 70)
        self.report_table.hideColumn(3)
        self.report_table.horizontalHeader().setStretchLastSection(True)
        self.report_table.verticalHeader().setDefaultSectionSize(20)

    @pyqtSlot()
    def report_excel(self):
        """Функция выгружает отчет в формат .XLSX"""

        name_report = QFileDialog.getSaveFileName(parent=self,
                                                  caption='Сохранение отчета',
                                                  directory='Reports',
                                                  filter='*.xlsx',
                                                  )
        if name_report[0] == '':
            return False
        self.table_database.setSelectionMode(2)
        self.table_database.selectAll()
        data_report = [data.data() for data in self.table_database.selectedIndexes()]
        self.table_database.clearSelection()
        self.table_database.setSelectionMode(2)

        workbook = xlsxwriter.Workbook(name_report[0])
        worksheet = workbook.add_worksheet(name='Отчет')
        worksheet.set_landscape()
        worksheet.set_paper(9)
        cell_format_center = workbook.add_format()
        cell_format_left = workbook.add_format()
        cell_format_center.set_align('center')
        cell_format_center.set_align('vcenter')
        cell_format_left.set_align('left')
        cell_format_left.set_align('vcenter')
        worksheet.write_row(0, 0, [
            '№ п/п',
            'Дата',
            'Время',
            'Адрес',
            'Пар-я',
            'Этаж',
            'Кв.',
            'Заявитель',
            'Телефон',
            'Сообщение',
            'Примечание',
            'Диспетчер',
            'Передано',
            'Исполнение'], cell_format_center)
        worksheet.set_column(0, 0, 7, cell_format_center)
        worksheet.set_column(1, 1, 12, cell_format_center)
        worksheet.set_column(2, 2, 7, cell_format_center)
        worksheet.set_column(4, 6, 7, cell_format_center)
        worksheet.set_column(3, 3, 24, cell_format_left)
        worksheet.set_column(7, 7, 24, cell_format_left)
        worksheet.set_column(8, 8, 20, cell_format_left)
        worksheet.set_column(9, 9, 50, cell_format_left)
        worksheet.set_column(10, 10, 17, cell_format_left)
        worksheet.set_column(11, 11, 17, cell_format_left)
        worksheet.set_column(12, 12, 17, cell_format_left)
        worksheet.set_column(13, 13, 40, cell_format_left)
        row = 1
        for i in range(0, len(data_report), 14):
            worksheet.write_row(row, 0, data_report[i:i + 14])
            row += 1
        try:
            workbook.close()
        except FileCreateError:
            self.message_logging('Ошибка доступа к файлу и/или к каталогу при выгрузке отчета в .xlsx', 'error')
            return False
        else:
            self.message_logging(f'Отчет сформирован. Имя/путь: {name_report[0]}', 'info')
            return True

    @pyqtSlot()
    def report_openoffice(self):
        """Функция выгружает отчет в формат .ODS"""

        name_report = QFileDialog.getSaveFileName(parent=self,
                                                  caption='Сохранение отчета в формате .ods',
                                                  directory='Reports',
                                                  filter='*.ods',
                                                  )
        if name_report[0] == '':
            return False
        self.table_database.setSelectionMode(2)
        self.table_database.selectAll()
        data_tmp = [data.data() for data in self.table_database.selectedIndexes()]
        self.table_database.clearSelection()
        self.table_database.setSelectionMode(2)
        data_report = [data_tmp[i:i + 14] for i in range(0, len(data_tmp), 14)]
        data_report = [[
            '№ п/п',
            'Дата',
            'Время',
            'Адрес',
            'Пар-я',
            'Этаж',
            'Кв.',
            'Заявитель',
            'Телефон',
            'Сообщение',
            'Примечание',
            'Диспетчер',
            'Передано',
            'Исполнение']] + data_report
        data_dict = OrderedDict()
        data_dict.update({"Отчет": [*data_report]})
        try:
            save_data(name_report[0], data_dict)
        except PermissionError:
            self.message_logging('Ошибка доступа к файлу и/или к каталогу при выгрузке отчета в .ods', 'error')
            return False
        else:
            self.message_logging(f'Отчет сформирован. Имя/путь: {name_report[0]}', 'info')
            return True

    @pyqtSlot()
    def report_save(self):
        """Функция сохраняет отчет

        В БД 'Отчет' (report) сохраняются настройки фильтра, по которому в данный момент
        в основную таблицу выведены данные"""

        report_date = datetime.datetime.now().date().strftime('%Y-%m-%d')
        report_name = text_clearing_characters(self.report_name.text(), 'text')
        if len(report_name) < 5:
            self.message_logging("Имя отчета должно быть не менее пяти символов. Запрещенные символы удаляются.")
            self.report_name.setText(report_name)
            return False
        report_select = '|'.join([str(data) for data in self.start_filtering()])
        report_save_data = [report_date, report_name, report_select]
        try:
            connection = sqlite3.connect(self.path_db)
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging(f"ОШИБКА БД! Повторите попытку.", 'error', False)
            return False
        else:
            try:
                connection.cursor().execute(
                    f"INSERT INTO report(date_creation, name_report, report) "
                    f"VALUES('{report_save_data[0]}', "
                    f"'{report_save_data[1]}', "
                    f"'{report_save_data[2]}')")
                connection.commit()
                connection.close()
            except sqlite3.OperationalError:
                logging.exception("Exception occurred")
                self.message_logging(f"Ошибка сохранения отчета! Повторите попытку.", 'error')
                return False
            finally:
                self.message_logging(
                    f"Сохранен новый отчет. "
                    f"Имя отчета '{report_save_data[1]}'", 'info')
                self.report_name.clear()
                self.create_dbmodel_report()
                self.create_table_report()
        return True

    @pyqtSlot()
    def report_load(self):
        """Функция загружает выбранный отчет и выводит данные в основной таблице"""

        self.report_table.showColumn(3)
        report_load_data = [data.data() for data in self.report_table.selectedIndexes()]
        self.report_table.hideColumn(3)
        report_load = report_load_data[-1].split('|')

        if report_load[2] == '0':
            report_address = "ВСЕ"
        else:
            report_address = dispatchers_or_address_from_id(self.path_db, report_load[2], 'address')
        if report_load[5] == '0':
            report_dispatcher = "ВСЕ"
        else:
            report_dispatcher = dispatchers_or_address_from_id(self.path_db, report_load[5], 'dispatchers')

        report_load[2] = int(report_load[2])
        report_load[5] = int(report_load[5])
        report_load[6] = True if report_load[6] == 'True' else False
        if report_load_data:
            self.filter_date_start.setDate(datetime.datetime.strptime(report_load[0], "%Y-%m-%d").date())
            self.filter_date_finish.setDate(datetime.datetime.strptime(report_load[1], "%Y-%m-%d").date())
            self.filter_address.setCurrentText(report_address)
            self.filter_consumer_name.setText(report_load[3])
            self.filter_consumer_phone.setText(report_load[4])
            self.filter_dispatcher_name.setCurrentText(report_dispatcher)
            self.filter_no_close.setChecked(report_load[6])

            self.filtering_dbmodel(report_load)
            self.create_table_database()
        else:
            self.message_logging(f"Для загрузки отчета выберите его название в списке.")
            return False
        return True

    @pyqtSlot()
    def report_del(self):
        """Функция удаляет выбранный отчет"""

        if self.report_table.currentIndex().row() == -1:
            self.message_logging(f"Для удаления отчета выберите его название в списке.")
            return False
        check = QMessageBox(QMessageBox.Question, 'Удаление отчета', 'Вы точно хотите удалить сохраненный отчет?',
                            parent=self
                            )
        check.addButton('Да', 5)
        check.addButton('Отменить', 6)
        check.exec()
        if check.clickedButton().text() != 'Да':
            return False
        self.report_table.showColumn(3)
        report_del_data = [data.data() for data in self.report_table.selectedIndexes()]
        self.report_table.hideColumn(3)
        try:
            connection = sqlite3.connect(self.path_db)
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging(f"ОШИБКА БД! Повторите попытку.", 'error', False)
            return False
        else:
            try:
                connection.cursor().execute(f"DELETE FROM report WHERE id={report_del_data[0]}")
                connection.commit()
                connection.close()
            except sqlite3.OperationalError:
                logging.exception("Exception occurred")
                self.message_logging(f"Ошибка удаления отчета! Повторите попытку.", 'error')
                return False
            finally:
                self.message_logging(
                    f"Удален отчет. "
                    f"Имя отчета '{report_del_data[2]}'", 'info')
                self.create_dbmodel_report()
                self.create_table_report()
        return True

    # --------- панель настройки -------------------------
    # --------- таблица диспетчеры
    def create_table_conf_dispatcher(self):
        """Функция настройки таблицы для вывода БД 'диспетчеры' """

        self.conf_dispatcher_table.setSelectionMode(1)
        self.conf_dispatcher_table.verticalHeader().hide()
        self.conf_dispatcher_table.hideColumn(0)
        self.conf_dispatcher_table.resizeRowToContents(1)
        self.conf_dispatcher_table.horizontalHeader().setStretchLastSection(True)
        self.conf_dispatcher_table.verticalHeader().setDefaultSectionSize(20)

    @pyqtSlot()
    def conf_dispatcher_edit(self):
        if self.dbmodel_dispatcher.isDirty():
            check = QMessageBox(QMessageBox.Question, 'Изменение таблицы <Диспетчеры>',
                                'После сохранения изменений, обязательно перезапустите программу! \n'
                                'Внести изменения?', parent=self
                                )
            check.addButton('Да', 5)
            check.addButton('Отменить', 6)
            check.exec()
            if check.clickedButton().text() != 'Да':
                self.conf_dispatcher_cancel()
                return False
            self.db.open()
            transferred = self.dbmodel_dispatcher.submitAll()
            self.db.close()
            self.create_dbmodel_dispatcher()
            self.create_table_conf_dispatcher()
            if transferred:
                self.message_logging("Изменения в таблице 'dispatcher' сохранены", 'info')
                return True
            else:
                self.message_logging("Изменения в таблице не сохранены! Попробуйте снова.")
                return False
        return False

    @pyqtSlot()
    def conf_dispatcher_cancel(self):
        self.conf_dispatcher_new_name.clear()
        if self.dbmodel_dispatcher.isDirty():
            self.create_dbmodel_dispatcher()
            self.create_table_conf_dispatcher()

    @pyqtSlot()
    def conf_dispatcher_new(self):
        dispatcher_new_name = text_clearing_characters(self.conf_dispatcher_new_name.text(), 'dispatcher')
        self.conf_dispatcher_new_name.clear()
        if len(dispatcher_new_name) < 4:
            self.message_logging("Введите фамилию и инициалы нового диспетчера (не может быть короче 4 символов).")
            return False
        try:
            connection = sqlite3.connect(self.path_db)
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging(f"ОШИБКА БД! Повторите попытку.", 'error', False)
            return False
        else:
            try:
                connection.cursor().execute(
                    f"INSERT INTO dispatcher(name, active) "
                    f"VALUES('{dispatcher_new_name}', "
                    f"'да')")
                connection.commit()
                id_new_dispatchers = id_dispatchers_or_address(self.path_db, dispatcher_new_name, 'dispatchers')
                connection.cursor().execute(
                    f"INSERT INTO setting(id_dispatcher, color_text) "
                    f"VALUES('{id_new_dispatchers}', "
                    f"'000000')")
                connection.commit()
                connection.close()
            except sqlite3.OperationalError:
                logging.exception("Exception occurred")
                self.message_logging(f"Ошибка добавления диспетчера! Повторите попытку.", 'error')
                return False
            finally:
                QMessageBox.information(self,
                                        'Добавлен новый диспетчер',
                                        'Для применения изменений перезапустите программу!',
                                        buttons=QMessageBox.Ok)
                self.message_logging(
                    f"В базу данных добавлен новый диспетчер. "
                    f"Имя: '{dispatcher_new_name}'. Статус активности: 'да'", 'info')
                self.create_dbmodel_dispatcher()
                self.create_table_conf_dispatcher()
                return True

    # --------- таблица адреса
    def create_table_conf_address(self):
        """Функция настройки таблицы для вывода БД 'адреса' """

        self.conf_address_table.setSelectionMode(1)
        self.conf_address_table.verticalHeader().hide()
        self.conf_address_table.hideColumn(0)
        self.conf_address_table.horizontalHeader().setStretchLastSection(True)
        self.conf_address_table.verticalHeader().setDefaultSectionSize(20)

    @pyqtSlot()
    def conf_address_edit(self):
        if self.dbmodel_address.isDirty():
            check = QMessageBox(QMessageBox.Question, 'Изменение таблицы <Адреса>',
                                'После сохранения изменений, обязательно перезапустите программу! \n'
                                'Внести изменения?', parent=self
                                )
            check.addButton('Да', 5)
            check.addButton('Отменить', 6)
            check.exec()
            if check.clickedButton().text() != 'Да':
                self.conf_address_cancel()
                return False
            self.db.open()
            transferred = self.dbmodel_address.submitAll()
            self.db.close()
            self.create_dbmodel_address()
            self.create_table_conf_address()
            if transferred:
                self.message_logging("Изменения в таблице 'address' сохранены", 'info')
                return True
            else:
                self.message_logging("Изменения в таблице не сохранены! Попробуйте снова.")
                return False
        return False

    @pyqtSlot()
    def conf_address_cancel(self):
        self.conf_address_new_building.clear()
        if self.dbmodel_address.isDirty():
            self.create_dbmodel_address()
            self.create_table_conf_address()

    @pyqtSlot()
    def conf_address_new(self):
        address_new_building = text_clearing_characters(self.conf_address_new_building.text(), 'address')
        self.conf_address_new_building.clear()
        if len(address_new_building) < 1:
            self.message_logging("Введите новый адрес (улицу можно подгрузить из списка 'Улицы').")
            return False
        try:
            connection = sqlite3.connect(self.path_db)
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging(f"ОШИБКА БД! Повторите попытку.", 'error', False)
            return False
        else:
            try:
                connection.cursor().execute(
                    f"INSERT INTO address(building) "
                    f"VALUES('{address_new_building}')")
                connection.commit()
                connection.close()
            except sqlite3.OperationalError:
                logging.exception("Exception occurred")
                self.message_logging(f"Ошибка сохранения адреса! Повторите попытку.", 'error')
                return False
            finally:
                QMessageBox.information(self,
                                        'Добавлен новый адрес',
                                        'Для применения изменений перезапустите программу!',
                                        buttons=QMessageBox.Ok)
                self.message_logging(f"В базу данных добавлен новый адрес: '{address_new_building}'", 'info')
                self.create_dbmodel_address()
                self.create_table_conf_address()
                return True

    @pyqtSlot()
    def conf_select_street(self):
        try:
            connection = sqlite3.connect(self.path_db)
            street_list = connection.cursor().execute("SELECT title"
                                                      " FROM streets"
                                                      " ORDER BY title").fetchall()
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging('Ошибка работы с БД! Обратитесь к администратору.', 'error', False)
            return False
        else:
            connection.close()
        street_list = [data[0] for data in street_list]
        street_dialog = QInputDialog.getItem(self, 'Названия улиц Кронштадта',
                                             'Выберите улицу для ее вставки в поле',
                                             street_list,
                                             editable=True)
        if street_dialog[1]:
            self.conf_address_new_building.setText(street_dialog[0] + ' ')
            return True
        return False

    # --------- таблица шаблоны обращений
    def create_table_conf_pattern(self):
        """Функция настройки таблицы для вывода БД 'шаблоны обращений' """

        self.conf_pattern_table.setSelectionMode(1)
        self.conf_pattern_table.verticalHeader().hide()
        self.conf_pattern_table.hideColumn(0)
        self.conf_pattern_table.horizontalHeader().setStretchLastSection(True)
        self.conf_pattern_table.verticalHeader().setDefaultSectionSize(20)

    @pyqtSlot()
    def conf_pattern_edit(self):
        if self.dbmodel_pattern.isDirty():
            check = QMessageBox(QMessageBox.Question, 'Изменение таблицы <Шаблоны>',
                                'После сохранения изменений, обязательно перезапустите программу! \n'
                                'Внести изменения?', parent=self
                                )
            check.addButton('Да', 5)
            check.addButton('Отменить', 6)
            check.exec()
            if check.clickedButton().text() != 'Да':
                self.conf_pattern_cancel()
                return False
            self.db.open()
            transferred = self.dbmodel_pattern.submitAll()
            self.db.close()
            self.create_dbmodel_pattern()
            self.create_table_conf_pattern()
            if transferred:
                self.message_logging("Изменения в таблице 'pattern' сохранены", 'info')
                return True
            else:
                self.message_logging("Изменения в таблице не сохранены! Попробуйте снова.")
                return False
        return False

    @pyqtSlot()
    def conf_pattern_cancel(self):
        self.conf_pattern_new_message.clear()
        if self.dbmodel_pattern.isDirty():
            self.create_dbmodel_pattern()
            self.create_table_conf_pattern()

    @pyqtSlot()
    def conf_pattern_del(self):
        pattern_del_index = self.conf_pattern_table.currentIndex().row()
        if pattern_del_index == -1:
            self.message_logging("Для удаления выберите шаблон в списке.")
            return False
        else:
            self.dbmodel_pattern.removeRow(pattern_del_index)
            self.db.open()
            transferred = self.dbmodel_pattern.submitAll()
            self.db.close()
            self.create_dbmodel_pattern()
            self.create_table_conf_pattern()
            if transferred:
                self.message_logging("Выполнено удаление шаблона обращений.", 'info')
                return True
            else:
                self.message_logging("Ошибка при удалении! Попробуйте снова.")
                return False

    @pyqtSlot()
    def conf_pattern_new(self):
        pattern_new = text_clearing_characters(self.conf_pattern_new_message.text(), 'text')
        self.conf_pattern_new_message.clear()
        if len(pattern_new) < 1:
            self.message_logging("Введите шаблон обращения.")
            return False
        try:
            connection = sqlite3.connect(self.path_db)
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging(f"ОШИБКА БД! Повторите попытку.", 'error', False)
            return False
        else:
            try:
                connection.cursor().execute(
                    f"INSERT INTO pattern_message(pattern) "
                    f"VALUES('{pattern_new}')")
                connection.commit()
                connection.close()
            except sqlite3.OperationalError:
                logging.exception("Exception occurred")
                self.message_logging(f"Ошибка сохранения шаблона! Повторите попытку.", 'error')
                return False
            finally:
                QMessageBox.information(self,
                                        'Новый шаблон обращения',
                                        'Для применения изменений перезапустите программу!',
                                        buttons=QMessageBox.Ok)
                self.message_logging(f"В базу данных добавлен новый шаблон обращения: '{pattern_new}'", 'info')
                self.create_dbmodel_pattern()
                self.create_table_conf_pattern()
                return True

    # ----------- конфигурация
    @pyqtSlot()
    def archive_database(self):
        archive3 = 'Archive_db/archive_db3.db'
        archive2 = 'Archive_db/archive_db2.db'
        archive1 = 'Archive_db/archive_db1.db'
        if os.path.isfile(archive3):
            try:
                os.remove(archive3)
            except OSError:
                logging.exception("Exception occurred")
        if os.path.isfile(archive2):
            try:
                os.rename(archive2, archive3)
            except OSError:
                logging.exception("Exception occurred")
        if os.path.isfile(archive1):
            try:
                os.rename(archive1, archive2)
            except OSError:
                logging.exception("Exception occurred")
        try:
            copyfile(self.path_db, archive1)
        except OSError:
            logging.exception("Exception occurred")
            return False
        finally:
            self.create_archive_panel()
            self.message_logging('Архив базы данных успешно создан.', 'info')

    def create_archive_panel(self):
        self.conf_list_archive_db.clear()
        for i in range(1, 4):
            path_archive_db = f'Archive_db/archive_db{i}.db'
            if os.path.isfile(path_archive_db):
                self.conf_list_archive_db.insertPlainText(f"Архив базы {i}: {file_creation_date(path_archive_db)}\n")
            else:
                self.conf_list_archive_db.insertPlainText(f"Архив базы {i}: отсутствует\n")
        self.conf_list_archive_db.moveCursor(QTextCursor.Start)

    @pyqtSlot()
    def setting_color_dispatcher(self):
        color_dispatcher = QColorDialog.getColor(initial=QColor(self.setting_dict['color_text']),
                                                 parent=self,
                                                 title='Выберите цвет отображения имени',
                                                 options=QColorDialog.DontUseNativeDialog)
        if color_dispatcher.isValid():
            self.setting_dict['color_text'] = color_dispatcher.name()
            self.create_head_panel()
            return True
        return False

    @pyqtSlot()
    def setting_dispatcher_save(self):
        """Функция сохраняет настройки активного диспетчера"""
        try:
            connection = sqlite3.connect(self.path_db)
            connection.cursor().execute(
                f"UPDATE setting SET "
                f"color_text='{self.setting_dict['color_text'][1:]}' "
                f"WHERE id_dispatcher='{self.setting_dict['id_dispatcher']}';")
            connection.commit()
            connection.close()
        except sqlite3.OperationalError:
            logging.exception("Exception occurred")
            self.message_logging("Не удалось сохранить настройки")
            return False
        else:
            connection.close()
            self.message_logging("Настройки сохранены.", 'info')
            return True

    # ------- панель логирования ------------------
    def create_logging_box(self):
        try:
            with open('log.txt', 'r') as file:
                self.logging.setText(file.read())
                self.logging.moveCursor(QTextCursor.End)
        except FileNotFoundError:
            pass

    def message_logging(self, message, type_message='warning', file_write=True):
        """Функция передачи сообщений в поле логирования и запись в log.txt

        message - передаваемое сообщение
        type_message='warning' - тип сообщения (warning, info, error)
        file_write=True - производится ли запись в файл"""

        if type_message == 'info':
            self.logging.setStyleSheet("QTextEdit {background-color:#9cde9c}")
            if file_write:
                logging.info(f"{DISPATCHER}: {message}")
        elif type_message == 'error':
            self.logging.setStyleSheet("QTextEdit {background-color:#ff5050}")
            if file_write:
                logging.error(f"{DISPATCHER}: {message}")
        else:
            self.logging.setStyleSheet("QTextEdit {background-color:#fdc0c0}")
            if file_write:
                logging.warning(f"{DISPATCHER}: {message}")
        self.logging.moveCursor(QTextCursor.End)
        self.logging.insertPlainText(f"{message}\n")
        self.single_timer.singleShot(300, self.logging_background_white)
        return True

    @pyqtSlot()
    def logging_background_white(self):
        """Функция устанавливает белый фон для поля логирования (вызывается таймером)"""

        self.logging.setStyleSheet("QTextEdit {background-color:white}")

    @pyqtSlot()
    def logging_font_size(self):
        """Функция изменения шрифта в поле логирования"""

        self.logging.setFont(QFont("Arial", self.logging_slider_font.value() * 2))
        self.logging.moveCursor(QTextCursor.End)


def main_application():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # ['Breeze', 'Oxygen', 'QtCurve', 'Windows', 'Fusion']

    path_db = 'DataBase/TEST_Dispatcher_db.db'
    start_window = StartWindow(path_db)
    start_window.show()
    start_window.exec()

    if START:
        logging.info(f"Запуск БД. Пользователь: {DISPATCHER}")
        program_bd = MainWindow(path_db)
        program_bd.show()
        sys.exit(app.exec())

    app.quit()


if __name__ == "__main__":
    main_application()

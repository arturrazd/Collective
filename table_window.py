import calendar
import time
import locale
from datetime import datetime, timedelta
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QIcon, QColor
from PyQt5.QtWidgets import *
from psycopg2 import connect
from data_base import DataBase
from styles import Styles
from styles import AlignDelegate


locale.setlocale(category=locale.LC_ALL, locale="Russian")


class TableWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.table_report = TableReport(self)
        main_btn_size = QSize(30, 30)
        input_btn_size = QSize(150, 30)
        # фильтр по должности
        self.combo_role = QComboBox(self, objectName='combo_role')
        self.list_role = self.get_list_role()
        self.combo_role.addItems(self.list_role)
        self.combo_role.setFixedSize(input_btn_size)
        self.combo_role.setStyleSheet(Styles.workers_combo())
        self.combo_role.activated.connect(self.table_report.set_filter_table)
        self.combo_role.activated.connect(self.table_report.check_table)
        # фильтр по отделу/цеху
        self.combo_gild = QComboBox(self, objectName='combo_gild')
        self.list_gild = self.get_list_gild()
        self.combo_gild.addItems(self.list_gild)
        self.combo_gild.setFixedSize(input_btn_size)
        self.combo_gild.setStyleSheet(Styles.workers_combo())
        self.combo_gild.activated.connect(self.table_report.set_filter_table)
        self.combo_gild.activated.connect(self.table_report.check_table)
        # выбор года
        self.combo_years = QComboBox(self, objectName='combo_years')
        self.list_years = list(self.get_list_dates().keys())
        self.combo_years.addItems(self.list_years)
        self.combo_years.setCurrentText(str(datetime.now().date().year))
        self.combo_years.setFixedSize(input_btn_size)
        self.combo_years.setStyleSheet(Styles.workers_combo())
        self.combo_years.activated.connect(self.change_list_month)
        self.combo_years.activated.connect(self.table_report.set_filter_table)
        self.combo_years.activated.connect(self.table_report.bild_table)
        # выбор месяца
        self.combo_month = QComboBox(self, objectName='combo_month')
        self.dict_month = self.get_list_dates()
        self.list_month = list(self.dict_month[self.combo_years.currentText()].values())
        self.combo_month.addItems(self.list_month)
        current_month = datetime.strptime(str(datetime.now().date()), "%Y-%m-%d").strftime("%B")
        self.combo_month.setCurrentText(current_month)
        self.combo_month.setFixedSize(input_btn_size)
        self.combo_month.setStyleSheet(Styles.workers_combo())
        self.combo_month.activated.connect(self.table_report.set_filter_table)
        self.combo_month.activated.connect(self.table_report.bild_table)
        # кнопка "обновить"
        self.btn_refresh = QPushButton('', self)
        self.btn_refresh.setIcon(QtGui.QIcon('refresh.png'))
        self.btn_refresh.setIconSize(QtCore.QSize(20, 20))
        self.btn_refresh.setFixedSize(main_btn_size)
        self.btn_refresh.setStyleSheet(Styles.workres_btn())
        self.btn_refresh.clicked.connect(self.table_report.check_table)
        # кнопка "копировать"
        self.btn_copy = QPushButton('', self)
        self.btn_copy.setIcon(QtGui.QIcon('copy_report.png'))
        self.btn_copy.setIconSize(QtCore.QSize(20, 20))
        self.btn_copy.setFixedSize(main_btn_size)
        self.btn_copy.setStyleSheet(Styles.workres_btn())
        self.btn_copy.clicked.connect(self.table_report.copy_report)
        # кнопка "вставить"
        self.btn_paste = QPushButton('', self)
        self.btn_paste.setIcon(QtGui.QIcon('paste_report.png'))
        self.btn_paste.setIconSize(QtCore.QSize(20, 20))
        self.btn_paste.setFixedSize(main_btn_size)
        self.btn_paste.setStyleSheet(Styles.workres_btn())
        self.btn_paste.clicked.connect(self.table_report.paste_report)
        # кнопка "очистить"
        self.btn_clear = QPushButton('', self)
        self.btn_clear.setIcon(QtGui.QIcon('delete_report.png'))
        self.btn_clear.setIconSize(QtCore.QSize(20, 20))
        self.btn_clear.setFixedSize(main_btn_size)
        self.btn_clear.setStyleSheet(Styles.workres_btn())
        self.btn_clear.clicked.connect(self.table_report.clear_report)
        # отработанные часы
        self.combo_hours = QComboBox(self)
        self.combo_hours.setFixedSize(input_btn_size)
        self.combo_hours.addItem(QIcon('time.png'), 'Отработано часов')
        [self.combo_hours.addItem(QIcon('time.png'), str(i)) for i in range(1, 25)]
        self.combo_hours.setStyleSheet(Styles.workers_combo())
        self.combo_hours.activated.connect(self.table_report.edit_report)
        # оценка
        self.combo_rating = QComboBox(self)
        self.combo_rating.setFixedSize(input_btn_size)
        self.combo_rating.addItem(QIcon('rating.png'), 'Оценка работы')
        [self.combo_rating.addItem(QIcon('rating_' + str(i) + '.png'), str(i)) for i in range(1, 4)]
        self.combo_rating.setStyleSheet(Styles.workers_combo())
        self.combo_rating.activated.connect(self.table_report.edit_report)
        # статус работника (полный день, отгул, отпуск, больничный)
        self.combo_status = QComboBox(self)
        self.combo_status.addItem('Полный день')
        self.combo_status.addItem(QIcon('day_off.png'), 'Отгул')
        self.combo_status.addItem(QIcon('vacation.png'), 'Отпуск')
        self.combo_status.addItem(QIcon('sick.png'), 'Больничный')
        self.combo_status.setFixedSize(input_btn_size)
        self.combo_status.setStyleSheet(Styles.workers_combo())
        self.combo_status.activated.connect(self.table_report.edit_report)
        # курсоры
        self.btn_refresh.setCursor(Qt.PointingHandCursor)
        self.combo_hours.setCursor(Qt.PointingHandCursor)
        self.combo_rating.setCursor(Qt.PointingHandCursor)
        self.combo_status.setCursor(Qt.PointingHandCursor)
        self.table_report.setCursor(Qt.PointingHandCursor)

        self.combo_hours.setToolTip('должность')
        self.combo_rating.setToolTip('цех/отдел')

        self.page_layout = QHBoxLayout()
        self.page_layout.setContentsMargins(0, 0, 0, 0)
        self.panel_layout = QVBoxLayout()
        self.panel_layout.setContentsMargins(0, 0, 0, 0)
        self.button_layout = QHBoxLayout()
        self.button_layout.setContentsMargins(0, 0, 0, 0)
        self.edit_layout = QVBoxLayout()
        self.edit_layout.setContentsMargins(0, 0, 0, 0)
        self.filter_layout = QVBoxLayout()
        self.filter_layout.setContentsMargins(0, 0, 0, 0)

        self.button_layout.addWidget(self.btn_refresh)
        self.button_layout.addWidget(self.btn_copy)
        self.button_layout.addWidget(self.btn_paste)
        self.button_layout.addWidget(self.btn_clear)

        self.filter_layout.addWidget(self.combo_years)
        self.filter_layout.addWidget(self.combo_month)
        self.filter_layout.addWidget(self.combo_role)
        self.filter_layout.addWidget(self.combo_gild)

        self.edit_layout.addWidget(self.combo_hours)
        self.edit_layout.addWidget(self.combo_rating)
        self.edit_layout.addWidget(self.combo_status)

        self.panel_layout.addLayout(self.button_layout)
        self.panel_layout.addLayout(self.edit_layout)
        self.panel_layout.addStretch(0)
        self.panel_layout.addLayout(self.filter_layout)
        self.panel_layout.setSpacing(10)

        self.page_layout.addWidget(self.table_report)
        self.page_layout.addLayout(self.panel_layout)

        self.setLayout(self.page_layout)

    def get_list_role(self):
        try:
            with connect(**DataBase.config()) as conn:
                with conn.cursor() as cursor:
                    cursor.execute(DataBase.sql_list_role())
                    list_role_out = ['Все сотрудники'] + [item for t in cursor.fetchall() for item in t if
                                                          isinstance(item, str)]
                    return list_role_out
        except:
            pass

    def get_list_gild(self):
        try:
            with connect(**DataBase.config()) as conn:
                with conn.cursor() as cursor:
                    cursor.execute(DataBase.sql_list_gild())
                    list_role_out = ['Все цеха/отделы'] + [item for t in cursor.fetchall() for item in t if
                                                          isinstance(item, str)]
                    return list_role_out
        except:
            pass

    def get_list_dates(self):
        try:
            with connect(**DataBase.config()) as conn:
                with conn.cursor() as cursor:
                    all_date_dict = dict()
                    cursor.execute(DataBase.sql_list_all_date())
                    date_list = cursor.fetchall()
                    year = 0
                    for date in date_list:
                        if year != date[0]:
                            all_date_dict[str(date[0])] = dict()
                            year = date[0]
                        all_date_dict[str(date[0])][str(date[1])] = str(date[2])
                        pass
                    return all_date_dict
        except:
            pass

    def change_list_month(self):
        self.combo_month.clear()
        self.list_month = list(self.get_list_dates()[self.combo_years.currentText()].values())
        self.combo_month.addItems(self.list_month)


# класс - таблица
class TableReport(QTableWidget):
    def __init__(self, wg):
        super().__init__(wg)
        self.filter_role = None
        self.filter_gild = None
        self.filter_month = None
        self.filter_year = None
        self.reports = None
        self.reports_dict = None
        self.buff_copy_hours = None
        self.buff_copy_rating = None
        self.buff_copy_status = None
        self.wg = wg
        self.verticalHeader().hide()
        self.horizontalHeader().setSectionsClickable(False)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setStyleSheet(Styles.table_table())
        self.horizontalHeader().setStyleSheet(Styles.table_header())
        self.setEditTriggers(QTableWidget.NoEditTriggers)  # запретить изменять поля
        self.cellClicked.connect(self.click_table)  # установить обработчик щелча мыши в таблице
        self.setMinimumSize(500, 500)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.add_month()
        self.set_filter_table()
        self.bild_table()

    def set_filter_table(self):
        if self.wg.findChild(QComboBox, name='combo_role') is None:
            self.filter_role = 'rs.role'
            self.filter_gild = 'gd.gild'
            self.filter_year = str(datetime.now().date().year)
            self.filter_month = str(datetime.now().date().month)
        else:
            months = {'Январь': '1', 'Февраль': '2', 'Март': '3', 'Апрель': '4', 'Май': '5', 'Июнь': '6',
                      'Июль': '7', 'Август': '8', 'Сентябрь': '9', 'Октябрь': '10', 'Ноябрь': '11', 'Декабрь': '12'}
            combo_role_text = self.wg.combo_role.currentText()
            self.filter_role = ("'" + combo_role_text + "'", 'rs.role')[combo_role_text == 'Все сотрудники']
            combo_gild_text = self.wg.combo_gild.currentText()
            self.filter_gild = ("'" + combo_gild_text + "'", 'gd.gild')[combo_gild_text == 'Все цеха/отделы']
            self.filter_year = self.wg.combo_years.currentText()
            self.filter_month = months[self.wg.combo_month.currentText()]

    def get_reports(self):
        try:
            with connect(**DataBase.config()) as conn:
                conn.autocommit = True
                with conn.cursor() as cursor:
                    cursor.execute(DataBase.sql_read_workers_report(self.filter_role, self.filter_year, self.filter_month, self.filter_gild))
                    self.reports = cursor.fetchall()
                    id_report = 0
                    new_reports_dict = dict()
                    for report in self.reports:
                        if id_report != report[7]:
                            new_reports_dict[report[7]] = dict()
                            id_report = report[7]
                        new_reports_dict[report[7]][report[8]] = {'hour': report[4], 'rating': report[5],
                                                                  'status': report[6]}
                    return new_reports_dict
        except:
            pass

    def get_dates(self, month):
        try:
            with connect(**DataBase.config()) as conn:
                conn.autocommit = True
                with conn.cursor() as cursor:
                    cursor.execute(DataBase.sql_read_date_report(), (month,))
                    return tuple(cursor.fetchall())
        except:
            pass

    def click_table(self, row, col):  # row - номер строки, col - номер столбца
        [self.item(i, 1).setForeground(QColor(200, 200, 200, 255)) for i in range(1, self.rowCount())]
        self.item(row, 1).setForeground(QColor(200, 100, 0, 255))

        self.check_table()
        if col > 2:
            id_worker = int(self.item(row, 2).text())
            id_date = int(self.item(0, col).text())
            self.wg.combo_hours.setCurrentIndex(self.reports_dict[id_worker][id_date]['hour'])
            self.wg.combo_rating.setCurrentIndex(self.reports_dict[id_worker][id_date]['rating'])
            self.wg.combo_status.setCurrentIndex(self.reports_dict[id_worker][id_date]['status'])

    def copy_report(self):
        row = self.currentRow()
        col = self.currentColumn()
        id_worker = int(self.item(row, 2).text())
        id_date = int(self.item(0, col).text())
        self.buff_copy_hours = self.reports_dict[id_worker][id_date]['hour']
        self.buff_copy_rating = self.reports_dict[id_worker][id_date]['rating']
        self.buff_copy_status = self.reports_dict[id_worker][id_date]['status']

    def paste_report(self):
        row = self.currentRow()
        col = self.currentColumn()
        hour = self.buff_copy_hours
        rating = self.buff_copy_rating
        status = self.buff_copy_status
        id_worker = int(self.item(row, 2).text())
        id_date = int(self.item(0, col).text())
        self.write_report(hour, rating, status, id_worker, id_date)

    def edit_report(self):
        row = self.currentRow()
        col = self.currentColumn()
        hour = self.wg.combo_hours.currentIndex()
        rating = self.wg.combo_rating.currentIndex()
        status = self.wg.combo_status.currentIndex()
        id_worker = int(self.item(row, 2).text())
        id_date = int(self.item(0, col).text())
        self.write_report(hour, rating, status, id_worker, id_date)

    def clear_report(self):
        row = self.currentRow()
        col = self.currentColumn()
        hour = 0
        rating = 0
        status = 0
        id_worker = int(self.item(row, 2).text())
        id_date = int(self.item(0, col).text())
        self.write_report(hour, rating, status, id_worker, id_date)

    def write_report(self, hour, rating, status, id_worker, id_date):
        try:
            with connect(**DataBase.config()) as conn:
                conn.autocommit = True
                with conn.cursor() as cursor:
                    cursor.execute(DataBase.sql_edit_table(), (hour, rating, status, id_worker, id_date))
            self.check_table()
        except:
            pass

    def check_table(self):
        new_reports_dict = self.get_reports()
        keys_equal = True
        if len(new_reports_dict) == len(self.reports_dict):
            for key in new_reports_dict:
                if key not in self.reports_dict:
                    keys_equal = False
                    break
        else:
            keys_equal = False
        if keys_equal:
            if new_reports_dict != self.reports_dict:
                self.updt_table(new_reports_dict)
                self.reports_dict = new_reports_dict
        else:
            self.bild_table()

    def updt_table(self, new_reports_dict):
        start_time = time.time()
        current_row = self.currentRow()
        current_col = self.currentColumn()
        i = 1
        for id_worker, data_worker in new_reports_dict.items():
            if id_worker in self.reports_dict:
                if data_worker != self.reports_dict[id_worker]:
                    j = 0
                    for id_day, data in data_worker.items():
                        if data != self.reports_dict[id_worker][id_day]:
                            self.fill_table(data, i, j)
                        j += 1

            else:
                self.bild_table()
                break
            i += 1

        print(str(datetime.now())[:19], "updt_table %s seconds ---" % str(time.time() - start_time)[:5])
        self.setCurrentCell(current_row, current_col)

    # обновление таблицы
    def bild_table(self):
        start_time = time.time()
        current_row = self.currentRow()
        current_col = self.currentColumn()
        scroll = self.verticalScrollBar()
        old_position_scroll = scroll.value() / (scroll.maximum() or 1)
        self.setColumnCount(0)
        self.setRowCount(0)
        dates = self.get_dates(self.filter_month)
        self.reports_dict = self.get_reports()

        fio_dict = {report[7]: report[0] + ' ' + report[1][:1] + '. ' + report[2][:1] + '.' for report in
                    self.reports}

        dates_dict = {1: 'пн', 2: 'вт', 3: 'ср', 4: 'чт', 5: 'пт', 6: 'сб', 7: 'вс'}
        i = 0
        for i in range(len(dates) + 3):
            self.setColumnCount(self.columnCount() + 1)
            self.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
            self.setHorizontalHeaderItem(i, QTableWidgetItem(str(i - 2) + '\n' + dates_dict[dates[i - 3][2]]))
        self.setRowCount(self.rowCount() + 1)

        i = 3
        for date in dates:
            self.setItem(0, i, QTableWidgetItem(str(date[0])))
            i += 1

        self.setRowHidden(0, True)  # скрываем нулевую строку с идентификаторами дней
        self.setColumnHidden(2, True)  # скрываем 3 столбец с идентификатором работника
        self.setHorizontalHeaderItem(0, QTableWidgetItem("№"))
        self.setHorizontalHeaderItem(1, QTableWidgetItem("ФИО"))
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        delegate = AlignDelegate(self)
        self.setItemDelegateForColumn(0, delegate)

        i = 1
        for worker_dict in self.reports_dict:
            self.setRowCount(self.rowCount() + 1)

            j, k = 0, 3
            for data in self.reports_dict[worker_dict].values():
                self.setItem(i, k, QTableWidgetItem())
                header_text = self.horizontalHeaderItem(k).text()
                if 'сб' in header_text or 'вс' in header_text:
                    self.item(i, k).setBackground(QColor(50, 50, 50, 255))
                self.fill_table(data, i, j)
                k += 1
                j += 1

            self.setRowHeight(i, 40)
            self.setItem(i, 0, QTableWidgetItem(str(i)))
            self.setItem(i, 1, QTableWidgetItem(fio_dict[worker_dict]))
            self.setItem(i, 2, QTableWidgetItem(str(worker_dict)))

            flags = self.item(i, 1).flags()
            flags &= ~QtCore.Qt.ItemIsEnabled
            self.item(i, 0).setFlags(flags)
            self.item(i, 1).setFlags(flags)
            i += 1

        self.setCurrentCell(current_row, current_col)
        scroll.setValue(round(old_position_scroll * scroll.maximum()))
        print(str(datetime.now())[:19], "bild_table %s seconds ---" % str(time.time() - start_time)[:5])

    def fill_table(self, data, i, j):
        data_hour = data['hour']
        data_status = data['status']
        data_rating = data['rating']

        status_string = {1: 'отг', 2: 'отп', 3: 'бол'}

        worker_hours = QtWidgets.QLabel(('', str(data_hour))[data_hour != 0])
        worker_hours.setFont(QFont("CalibriLight", 12, QtGui.QFont.Bold))
        worker_hours.setStyleSheet(Styles.color_hours(data_hour))

        worker_status = QtWidgets.QLabel(status_string.get(data_status, ''))
        worker_status.setFont(QFont("CalibriLight", 9, QtGui.QFont.Normal))
        worker_status.setStyleSheet(Styles.color_status(data_status))

        worker_rating = QtWidgets.QLabel(('', str(data_rating))[data_rating != 0])
        worker_rating.setFont(QFont("CalibriLight", 9, QtGui.QFont.Normal))
        worker_rating.setStyleSheet(Styles.color_rating(data_rating))

        cell_layout = QVBoxLayout()
        cell_layout.setContentsMargins(0, 0, 0, 0)

        upper_layout = QHBoxLayout()
        upper_layout.setContentsMargins(0, 0, 0, 0)
        upper_layout.addStretch(0)
        upper_layout.addWidget(worker_rating)

        center_layout = QHBoxLayout()
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.addStretch(0)
        center_layout.addWidget(worker_hours)
        center_layout.addStretch(0)

        lower_layout = QHBoxLayout()
        lower_layout.setContentsMargins(0, 0, 0, 0)
        lower_layout.addStretch(0)
        lower_layout.addWidget(worker_status)

        cell_layout.addLayout(upper_layout)
        cell_layout.addLayout(center_layout)
        cell_layout.addLayout(lower_layout)
        cell_layout.setSpacing(0)
        widget = QWidget()
        widget.setLayout(cell_layout)
        self.setCellWidget(i, j + 3, widget)

    # создаем в бд записи на месяц для сотрудников
    def add_month(self):
        try:
            with connect(**DataBase.config()) as conn:
                conn.autocommit = True
                with conn.cursor() as cursor:
                    cur_year = datetime.now().date().year
                    cur_month = datetime.now().date().month
                    cursor.execute(DataBase.sql_read_date_report(), (cur_month,))
                    date_list = cursor.fetchall()
                    if len(date_list) == 0:
                        month_range = calendar.monthrange(cur_year, cur_month)
                        first_number_day = month_range[0] + 1
                        first_day = datetime.today().replace(day=1)
                        for day in range(month_range[1]):
                            cursor.execute(DataBase.sql_insert_date_report(),
                                           (first_day + timedelta(days=day), first_number_day))
                            first_number_day = (first_number_day + 1, 1)[first_number_day > 6]
                        cursor.execute(DataBase.sql_read_date_report(), (cur_month,))
                        date_list = cursor.fetchall()
                        cursor.execute(DataBase.sql_read_workers_2())
                        worker_list = cursor.fetchall()
                        if len(worker_list) != 0:
                            for worker in worker_list:
                                for date in date_list:
                                    cursor.execute(DataBase.sql_insert_workers_report(),
                                                   (date[0], worker[0], 0, 0, 0))
        except:
            pass

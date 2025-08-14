import sys
from calendar import monthrange

import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox, QLabel, QFrame
from PyQt5.QtCore import Qt

class ExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Аренда")
        self.setGeometry(100, 100, 800, 600)

        #  Переменные для файлов arenda и payment
        self.arenda_file = None
        self.payment_file = None

        # Установить GUI
        layout = QVBoxLayout(self)

        self.lbl_arenda = QLabel("файл с информацией об аренде не загружен")
        self.lbl_arenda.setStyleSheet("color: red;")
        self.lbl_payment = QLabel("файл с информацией о платежах не загружен")
        self.lbl_payment.setStyleSheet("color: red;")
        self.lbl_result = QLabel("Результат анализа платежей")
        for lbl in (self.lbl_arenda, self.lbl_payment, self.lbl_result):
            lbl.setWordWrap(True)

        # Загрузить Аренда файл
        self.button_arenda = QPushButton("Показать Данные Файла Аренды")
        self.button_arenda.clicked.connect(self.load_arendaFile)

        # Загрузить выписку по платёжному счёту
        self.button_payment = QPushButton("Показать Выписку По Платёжному Счёту")
        self.button_payment.clicked.connect(self.load_paymentFile)

        # отслеживать платежи
        self.button_check = QPushButton("Отслеживать Платежи")
        self.button_check.setEnabled(False)
        self.button_check.clicked.connect(self.check_payment)

        # Визуальный разделитель
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)


        self.table = QTableWidget()
        layout.addWidget(self.button_arenda)
        layout.addWidget(self.lbl_arenda)
        layout.addWidget(self.button_payment)
        layout.addWidget(self.lbl_payment)
        layout.addWidget(sep)
        layout.addWidget(self.button_check)
        layout.addWidget(self.lbl_result, alignment=Qt.AlignTop)
        layout.addWidget(self.table)
        self.setLayout(layout)

    # создать dataframe с информацией из файла с информацией об аренде
    def load_arendaFile(self):

        # Запросите информацию о ГАРАЖАХ
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText("Загрузите файл с информацией о Гаражах")
        msg_box.setWindowTitle("Аренда Гаражов")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

        file_path, _ = QFileDialog.getOpenFileName(self,"Открывать файл","","файлы (*.xlsx)")

        if file_path:

            try:
                arenda_df = pd.read_excel(file_path)

                # добавить заголовки к df
                arenda_df.columns = ["Гараж", "Сумма","Дата"]

                # Отображение фрейма данных в layout
                arenda_df['Сумма'] = pd.to_numeric(arenda_df['Сумма'], errors='coerce')
                self.table.setRowCount(arenda_df.shape[0])
                self.table.setColumnCount(arenda_df.shape[1])
                self.table.setHorizontalHeaderLabels(arenda_df.columns)

                for i in range(arenda_df.shape[0]):
                    for j in range(arenda_df.shape[1]):
                        self.table.setItem(i, j, QTableWidgetItem(str(arenda_df.iat[i, j])))

                # сохранение данных в переменной arenda_file и обновление layout
                self.arenda_file = arenda_df
                self.lbl_arenda.setText(f"Информация об аренде загружена: {file_path}")
                self.lbl_arenda.setStyleSheet("color: green;")
                self.activate_check_button()
            except Exception as e:
                self.arenda_file = None
                self.lbl_arenda.setText("Аренда: Ошибка при извлечении информации")
                QMessageBox.critical(self, "Ошибка", f"Ошибка при извлечении информации Аренды:\n{e}")


    # создать df с информацией из файла с платежной информацией
    def load_paymentFile(self):

        # Запросите информацию о ПЛАТЕЖАХ
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText("Загрузите выписку по платёжному счёту")
        msg_box.setWindowTitle("Оплата аренды")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

        file_path, _ = QFileDialog.getOpenFileName(self,"Открывать выписку по платёжному счёту", "", "файлы (*.xlsx)")
        if file_path:

            try:
                df = pd.read_excel(file_path)
                payment_df = df.drop(df.columns[[1, 2, 3]], axis=1).dropna()

                # добавить заголовки к df
                payment_df.columns = ["Дата", "Сумма"]

                # фильтровать строки с текстом
                df_filtered = payment_df[~payment_df['Дата'].str.contains(r'[\u0400-\u04FF]', na=False) & ~payment_df['Сумма'].str.contains(r'[\u0400-\u04FF]', na=False)]

                # форматировать платежи
                df_filtered['Сумма'] = (
                    df_filtered['Сумма'].str.replace(' ', '', regex=False).str.replace(',', '.', regex=False).str.replace('+', '', regex=False))

                # преобразовать платежи в double
                df_filtered['Сумма'] = pd.to_numeric(df_filtered['Сумма'], errors='coerce')
                df_filtered = df_filtered.dropna(subset=['Сумма'])

                # Регулярное выражение для извлечения даты в формате dd.mm.yyyy
                pat = r'(?<!\d)(?P<d>\d{1,2})[.\-/]?(?P<m>\d{1,2})[.\-/]?(?P<y>\d{4})(?!\d)'

                # дата извлечения
                date = df_filtered['Дата'].astype(str).str.extract(pat)
                df_filtered['Дата'] = pd.to_datetime(
                    date['d'].str.zfill(2) + '.' + date['m'].str.zfill(2) + '.' + date['y'], dayfirst=True, errors='coerce')

                # удалить все данные, которые не удалось проанализировать
                df_filtered = df_filtered.dropna(subset=['Дата'])

                # Отображение таблицы в layout
                self.table.setRowCount(df_filtered.shape[0])
                self.table.setColumnCount(df_filtered.shape[1])
                self.table.setHorizontalHeaderLabels(df_filtered.columns)

                for i in range(df_filtered.shape[0]):
                    for j in range(df_filtered.shape[1]):
                        self.table.setItem(i, j, QTableWidgetItem(str(df_filtered.iat[i, j])))

                 # сохранение данных в переменной arenda_file и обновление layout
                self.payment_file = df_filtered
                self.lbl_payment.setText(f"Информация об оплатые загружена: {file_path}")
                self.lbl_payment.setStyleSheet("color: green;")
                self.activate_check_button()
            except Exception as e:
                self.payment_file = None
                self.lbl_payment.setText("Оплата: Ошибка при извлечении информации")
                QMessageBox.critical(self, "Ошибка", f"Ошибка при извлечении информации Оплаты:\n{e}")



    # Рассчитать сроки оплаты
    def payment_date(self, status_arenda):

        # составить договор df
        contracts = (status_arenda[['Гараж','Дата_x']].dropna().groupby('Гараж', as_index=False)['Дата_x'].min())

        # получить минимальную и максимальную дату
        valid_transfers = status_arenda['Дата_y'].dropna()
        if not valid_transfers.empty:
            per_min = valid_transfers.min().to_period('M')
            per_max = valid_transfers.max().to_period('M')
        else:
            # если нет перевода
            per_min = contracts['Дата_x'].min().to_period('M')
            per_max = contracts['Дата_x'].max().to_period('M')

        # рассчитать, если это конец месяца
        def end_of_month(ts: pd.Timestamp) -> bool:
            return ts.day == monthrange(ts.year, ts.month)[1]

        def due_date_for_month(start_date: pd.Timestamp, period: pd.Period) -> pd.Timestamp:
            y, m = period.year, period.month
            last_day = monthrange(y, m)[1]
            # если начало происходит в конце месяца, оплата будет произведена в конце месяца
            if end_of_month(start_date):
                return pd.Timestamp(y,m,last_day)
            # if not, same as the first day contract
            return pd.Timestamp(y, m, min(start_date.day, last_day))

        # составление календаря платежей
        periods = pd.period_range(per_min, per_max, freq='M')
        registers = []
        for _, row in contracts.iterrows():
            cod = row['Гараж']
            beginning = row['Дата_x']
            if pd.isna(beginning):
                continue
            from_month = beginning.to_period('M')
            months = periods[periods >= from_month]
            for p in months:
                registers.append({
                    'Гараж': cod,
                    'year_month': p,
                    'Дата оплаты (ожидаемая)': due_date_for_month(beginning, p)
                })
        cal = pd.DataFrame(registers)

        return cal.sort_values(['Гараж', 'Дата оплаты (ожидаемая)']).reset_index(drop=True)

    # платеж чеком
    def check_payment(self):
        try:
            arenda_df = self.arenda_file
            payment_df = self.payment_file
            # объединить рамки данных по арендной плате и платежам df
            status_arenda = arenda_df.merge(payment_df, on='Сумма', how='left', indicator=False)

            # последняя дата в файле платежей
            max_date = max(status_arenda['Дата_y'])

            # создание df с датами оплаты
            cal = self.payment_date(status_arenda)

            # дни до срока просрочки
            days_to_pay = 2

            # данные о платежах за первый месяц df
            payments_month = (status_arenda[['Гараж', 'Сумма', 'Дата_y']]
                              .dropna()
                              .assign(year_month=lambda x: x['Дата_y'].dt.to_period('M'))
                              .sort_values(['Гараж', 'Дата_y'])
                              .drop_duplicates(['Гараж', 'year_month'], keep='first'))

            # Статусная df со всеми данными, необходимыми для расчета статуса
            status = (cal.merge(payments_month, on=['Гараж', 'year_month'], how='left').rename(
                columns={'Дата_y': 'дата перевода'}))

            status['days_later'] = (status['дата перевода'] - status['Дата оплаты (ожидаемая)']).dt.days

            status['Статус'] = np.select(
                [
                    status['Дата оплаты (ожидаемая)'] > max_date,
                    status['days_later'] <= 0,
                    status['days_later'].between(1, days_to_pay)
                ],
                ['срок не наступил', 'получен', 'получен'],
                default='просрочен'
            )

            # вывести информацию о статусе df
            print(status.info(verbose=True))

            # очищенная версия фрейма df для отображения в GUI
            status_final = status.drop(['year_month', 'дата перевода', 'days_later'], axis=1)

            # Создание и заполнение данных GUI
            self.table.setRowCount(status_final.shape[0])
            self.table.setColumnCount(status_final.shape[1])
            self.table.setHorizontalHeaderLabels(status_final.columns)

            for i in range(status_final.shape[0]):
                for j in range(status_final.shape[1]):
                    self.table.setItem(i, j, QTableWidgetItem(str(status_final.iat[i, j])))
            self.lbl_result.setStyleSheet("color: blue;")
            QMessageBox.information(self, "", "Обработка данных успешно завершена")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"обработка данных не удалась:\n{e}")

    # Активирует button_track, если файлы payment_file и arenda_file существуют.
    def activate_check_button(self):
        activate = self.arenda_file is not None and self.payment_file is not None
        self.button_check.setEnabled(activate)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ExcelApp()
    win.show()
    sys.exit(app.exec_())

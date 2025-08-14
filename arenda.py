import sys
from pandas.tseries.offsets import MonthEnd
from calendar import monthrange

import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox

class ExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Аренда")
        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()

        # Загрузить Аренда файл
        self.button_arenda = QPushButton("Показать Данные Файла Аренды")
        self.button_arenda.clicked.connect(self.load_arendaFile)

        # Загрузить выписку по платёжному счёту
        self.button_payment = QPushButton("Показать Выписку По Платёжному Счёту")
        self.button_payment.clicked.connect(self.load_paymentFile)

        # отслеживать платежи
        self.button_track = QPushButton("Отслеживать Платежи")
        self.button_track.clicked.connect(self.check_payment)

        self.table = QTableWidget()
        layout.addWidget(self.button_arenda)
        layout.addWidget(self.button_payment)
        layout.addWidget(self.button_track)
        layout.addWidget(self.table)
        self.setLayout(layout)

    # create dataframe with info from file with the arenda information
    def load_arendaFile(self):

        # Ask for the information about the GARAGES
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText("Загрузите файл с информацией о Гаражах")
        msg_box.setWindowTitle("Аренда Гаражов")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

        file_path, _ = QFileDialog.getOpenFileName(self,"Открывать файл","","файлы (*.xlsx)")

        if file_path:
            arenda_df = pd.read_excel(file_path)

            # add headers to the df
            arenda_df.columns = ["Гараж", "Сумма","Дата"]
            arenda_df['Сумма'] = pd.to_numeric(arenda_df['Сумма'], errors='coerce')
            self.table.setRowCount(arenda_df.shape[0])
            self.table.setColumnCount(arenda_df.shape[1])
            self.table.setHorizontalHeaderLabels(arenda_df.columns)

            for i in range(arenda_df.shape[0]):
                for j in range(arenda_df.shape[1]):
                    self.table.setItem(i, j, QTableWidgetItem(str(arenda_df.iat[i, j])))
            return arenda_df


    # create df with info from file with payment information
    def load_paymentFile(self):

        # Ask for the information about the PAYMENTS
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText("Загрузите выписку по платёжному счёту")
        msg_box.setWindowTitle("Оплата аренды")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

        file_path, _ = QFileDialog.getOpenFileName(self,"Открывать выписку по платёжному счёту", "", "файлы (*.xlsx)")
        if file_path:
            df = pd.read_excel(file_path)
            payment_df = df.drop(df.columns[[1, 2, 3]], axis=1).dropna()

            # add headers to the df
            payment_df.columns = ["Дата", "Сумма"]

            # filter rows with text
            df_filtered = payment_df[~payment_df['Дата'].str.contains(r'[\u0400-\u04FF]', na=False) & ~payment_df['Сумма'].str.contains(r'[\u0400-\u04FF]', na=False)]

            # format the payments
            df_filtered['Сумма'] = (
                df_filtered['Сумма'].str.replace(' ', '', regex=False).str.replace(',', '.', regex=False).str.replace('+', '', regex=False))

            # convert payments to double
            df_filtered['Сумма'] = pd.to_numeric(df_filtered['Сумма'], errors='coerce')
            df_filtered = df_filtered.dropna(subset=['Сумма'])

            # regex for extracting date in format dd.mm.yyyy
            pat = r'(?<!\d)(?P<d>\d{1,2})[.\-/]?(?P<m>\d{1,2})[.\-/]?(?P<y>\d{4})(?!\d)'

            # extract date
            date = df_filtered['Дата'].astype(str).str.extract(pat)
            df_filtered['Дата'] = pd.to_datetime(
                date['d'].str.zfill(2) + '.' + date['m'].str.zfill(2) + '.' + date['y'], dayfirst=True, errors='coerce')

            # eliminate any data that could not be parsed
            df_filtered = df_filtered.dropna(subset=['Дата'])
            self.table.setRowCount(df_filtered.shape[0])
            self.table.setColumnCount(df_filtered.shape[1])
            self.table.setHorizontalHeaderLabels(df_filtered.columns)

            for i in range(df_filtered.shape[0]):
                for j in range(df_filtered.shape[1]):
                    self.table.setItem(i, j, QTableWidgetItem(str(df_filtered.iat[i, j])))
            return df_filtered

    # Calculate due payment dates
    def payment_date(self, status_arenda):

        # create a contract
        contracts = (status_arenda[['Гараж','Дата_x']].dropna().groupby('Гараж', as_index=False)['Дата_x'].min())

        # get minimum and maximum date
        valid_transfers = status_arenda['Дата_y'].dropna()
        if not valid_transfers.empty:
            per_min = valid_transfers.min().to_period('M')
            per_max = valid_transfers.max().to_period('M')
        else:
            # if there is no transfer
            per_min = contracts['Дата_x'].min().to_period('M')
            per_max = contracts['Дата_x'].max().to_period('M')

        # calculate if it is the end of the month
        def end_of_month(ts: pd.Timestamp) -> bool:
            return ts.day == monthrange(ts.year, ts.month)[1]

        def due_date_for_month(start_date: pd.Timestamp, period: pd.Period) -> pd.Timestamp:
            y, m = period.year, period.month
            last_day = monthrange(y, m)[1]
            # if the beginning is at the end of the month, payment will be at the end of the month
            if end_of_month(start_date):
                return pd.Timestamp(y,m,last_day)
            # if not, same as the first day contract
            return pd.Timestamp(y, m, min(start_date.day, last_day))

        # building the payment calendar
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

    # check payment
    def check_payment(self):
        arenda_df = self.load_arendaFile()
        payment_df = self.load_paymentFile()
        # merge the arenda and payment data frames
        status_arenda = arenda_df.merge(payment_df, on='Сумма', how='left', indicator=False)

        # latest date in the payment file
        max_date = max(status_arenda['Дата_y'])

        # generating dataframe with due payment dates
        cal = self.payment_date(status_arenda)

        # days before overdue date
        days_to_pay = 2

        # first month payment dataframe
        payments_month = (status_arenda[['Гараж','Сумма','Дата_y']]
                          .dropna()
                          .assign(year_month=lambda x: x['Дата_y'].dt.to_period('M'))
                          .sort_values(['Гараж','Дата_y'])
                          .drop_duplicates(['Гараж','year_month'], keep='first'))

        # Status dataframe with all needed for calculation of status
        status = (cal.merge(payments_month, on=['Гараж','year_month'], how='left').rename(columns={'Дата_y':'дата перевода'}))

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

        # print information about status dataframe
        print(status.info(verbose=True))

        # cleaned version of status dataframe for displaying in GUI
        status_final = status.drop(['year_month','дата перевода', 'days_later'],axis=1)

        # Creating and populating GUI dataframe
        self.table.setRowCount(status_final.shape[0])
        self.table.setColumnCount(status_final.shape[1])
        self.table.setHorizontalHeaderLabels(status_final.columns)

        for i in range(status_final.shape[0]):
            for j in range(status_final.shape[1]):
                self.table.setItem(i, j, QTableWidgetItem(str(status_final.iat[i, j])))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ExcelApp()
    win.show()
    sys.exit(app.exec_())

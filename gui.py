import os
import traceback
from datetime import date, datetime

from PyQt5.QtWidgets import QWidget, QPushButton, QVBoxLayout, QFileDialog, QDateEdit, QLineEdit, \
    QHBoxLayout, QFormLayout, QPlainTextEdit, QDialog, QDesktopWidget
from PyQt5.QtCore import QDir, QObject, pyqtSignal, QSettings, QThreadPool, QRunnable, pyqtSlot
from dateutil.relativedelta import relativedelta

from core import RESTAURANTS, process_data
from io import read_all_data, write_cancelled_df, write_all_invoices


class WorkerSignals(QObject):
    """
    Defines the signals available from a running worker thread.
    
    Custom signals can only be defined on objects derived from QObject. Since QRunnable
    is not derived from QObject we can't define the signals there directly. A custom QObject
    to hold the signals is the simplest solution.
    """
    finished = pyqtSignal()
    progress = pyqtSignal(str)


class Worker(QRunnable):

    def __init__(self, input_directory, from_date, to_date):
        super().__init__()
        self.input_directory = input_directory
        self.from_date = from_date
        self.to_date = to_date
        self.signals = WorkerSignals()

    @pyqtSlot()
    def run(self):
        # TODO: use logger object instead of passing around updater
        updater = lambda x: self.signals.progress.emit(x)
        try:
            df = read_all_data(updater, self.from_date, self.to_date, self.input_directory)

            suffix = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

            for restaurant in RESTAURANTS:
                serviced_df, cancelled_df = process_data(updater, self.from_date, self.to_date, df)

                if not cancelled_df.empty and not serviced_df.empty:
                    restaurant_base_path = os.path.join(self.input_directory, restaurant.value, suffix)
                    os.makedirs(restaurant_base_path, exist_ok=True)

                if cancelled_df.empty:
                    updater(f"No cancelled tours for {restaurant.value}")
                else:
                    write_cancelled_df(updater, restaurant_base_path, restaurant, cancelled_df)

                if serviced_df.empty:
                    updater(f"No tours found for {restaurant.value}")
                else:
                    write_all_invoices(updater, restaurant_base_path, restaurant, serviced_df)
        except Exception:
            updater(traceback.format_exc())
        finally:
            self.signals.finished.emit()


class QLoggingDialog(QDialog):

    signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.widget = QPlainTextEdit()
        self.widget.setReadOnly(True)
        self.signal.connect(self.widget.appendPlainText)

        layout = QVBoxLayout()
        layout.addWidget(self.widget)
        self.setLayout(layout)

        self.setWindowTitle("Logs")
        self.setMinimumSize(500, 300)

    def log(self, message):
        self.signal.emit(message)

    def clear_log(self):
        self.widget.clear()


class InvoiceGeneratorApp(QWidget):

    def __init__(self):
        super().__init__()

        self.settings = QSettings("Lucifer", "Invoice Generator")
        if self.settings.contains("input_directory"):
            self.existing_path = self.settings.value("input_directory")
        else:
            self.existing_path = ""

        self.threadpool = QThreadPool()

        self.init_ui()
        self.check_generate_button_state()

    def init_ui(self):
        self.setWindowTitle("Invoice Generator")
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

        self.input_dir_line_edit = QLineEdit()
        self.input_dir_line_edit.setText(self.existing_path)
        self.input_dir_line_edit.setReadOnly(True)
        self.input_dir_button = QPushButton("...")
        self.input_dir_button.clicked.connect(self.choose_input_directory)

        input_dir_layout = QHBoxLayout()
        input_dir_layout.addWidget(self.input_dir_line_edit)
        input_dir_layout.addWidget(self.input_dir_button)

        to_date = date.today()
        from_date = to_date + relativedelta(months=-1, day=1)

        self.from_date_selector = QDateEdit()
        self.from_date_selector.setDisplayFormat("dd/MM/yyyy")
        self.from_date_selector.setDate(from_date)
        self.from_date_selector.dateChanged.connect(self.choose_from_date)
        self.from_date_selector.setCalendarPopup(True)

        self.to_date_selector = QDateEdit()
        self.to_date_selector.setDisplayFormat("dd/MM/yyyy")
        self.to_date_selector.setDate(to_date)
        self.from_date_selector.dateChanged.connect(self.choose_to_date)
        self.to_date_selector.setCalendarPopup(True)

        self.generate_button = QPushButton("Generate Invoice")
        self.generate_button.clicked.connect(self.generate_invoice)
        self.generate_button.setEnabled(False)

        form_layout = QFormLayout()
        form_layout.addRow("Schedules:", input_dir_layout)
        form_layout.addRow("From Date:", self.from_date_selector)
        form_layout.addRow("To Date:", self.to_date_selector)

        layout = QVBoxLayout()
        layout.addLayout(form_layout)
        layout.addWidget(self.generate_button)

        self.setLayout(layout)

        self.logging_dialog = QLoggingDialog()

    def choose_from_date(self, from_date):
        self.to_date_selector.setMinimumDate(from_date)
        self.check_generate_button_state()

    def choose_to_date(self, to_date):
        self.check_generate_button_state()

    def choose_input_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Choose directory", self.existing_path or QDir.homePath())
        if directory:
            self.input_dir_line_edit.setText(directory)
            self.existing_path = directory
            self.settings.setValue("input_directory", directory)
            self.check_generate_button_state()

    def check_generate_button_state(self):
        input_dir_selected = bool(self.input_dir_line_edit.text())
        from_date_selected = not self.from_date_selector.date().isNull()
        to_date_selected = not self.to_date_selector.date().isNull()

        self.generate_button.setEnabled(input_dir_selected and from_date_selected and to_date_selected)

    def report_progress(self, message):
        self.logging_dialog.log(message)

    def disable_ui(self):
        self.input_dir_button.setEnabled(False)
        self.from_date_selector.setEnabled(False)
        self.to_date_selector.setEnabled(False)
        self.generate_button.setEnabled(False)

    def enable_ui(self):
        self.input_dir_button.setEnabled(True)
        self.from_date_selector.setEnabled(True)
        self.to_date_selector.setEnabled(True)
        self.generate_button.setEnabled(True)

    def generate_invoice(self):
        input_directory = self.input_dir_line_edit.text()
        from_date = self.from_date_selector.date().toPyDate()
        to_date = self.to_date_selector.date().toPyDate()

        self.disable_ui()

        self.logging_dialog.clear_log()
        self.logging_dialog.show()
        self.logging_dialog.raise_()

        worker = Worker(input_directory, from_date, to_date)
        worker.signals.progress.connect(self.report_progress)
        worker.signals.finished.connect(self.enable_ui)

        self.threadpool.start(worker)

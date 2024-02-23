import sys

from PyQt5.QtWidgets import QApplication

from gui import InvoiceGeneratorApp


def main():
    app = QApplication(sys.argv)

    window = InvoiceGeneratorApp()
    window.setMinimumWidth(500)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    # TODO: setup CI to package as exe
    main()

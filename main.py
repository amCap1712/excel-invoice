import sys

import pandas as pd
from PyQt5.QtWidgets import QApplication

from app.gui import InvoiceGeneratorApp

pd.options.mode.copy_on_write = True


def main():
    app = QApplication(sys.argv)

    window = InvoiceGeneratorApp()
    window.setMinimumWidth(500)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    # TODO: setup CI to package as exe
    main()

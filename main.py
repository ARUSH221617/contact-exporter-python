import sys
import os
import re
import pandas as pd

from PySide6.QtWidgets import *
from PySide6.QtCore import *


class VCardParser:

    def __init__(self, path):
        self.path = path
        self.contacts = []
        self.keys = set()

    def parse(self):

        current = {}

        with open(self.path, "r", encoding="utf-8", errors="ignore") as f:

            for line in f:

                line = line.strip()

                if line == "BEGIN:VCARD":
                    current = {}

                elif line == "END:VCARD":

                    if current:
                        self.contacts.append(current)

                else:

                    m = re.match(r"^([^:;]+)(?:;[^:]*)?:(.*)$", line)

                    if m:

                        key, value = m.groups()
                        key = key.split(";")[0].upper()

                        current[key] = value
                        self.keys.add(key)

        return self.contacts, sorted(self.keys)


class ExcelExporter:

    @staticmethod
    def export(data, fields, path):

        rows = []

        for c in data:

            row = {}

            for f in fields:
                row[f] = c.get(f, "")

            rows.append(row)

        df = pd.DataFrame(rows)

        df.to_excel(path, index=False)


class Worker(QThread):

    finished = Signal(object, object)

    def __init__(self, path):

        super().__init__()
        self.path = path

    def run(self):

        parser = VCardParser(self.path)

        contacts, fields = parser.parse()

        self.finished.emit(contacts, fields)


class MainWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle("VCF → Excel Converter (Pro)")
        self.resize(900, 600)

        self.contacts = []
        self.fields = []

        self._build_ui()

    def _build_ui(self):

        central = QWidget()
        self.setCentralWidget(central)

        layout = QVBoxLayout()

        top = QHBoxLayout()

        self.file_label = QLabel("No file selected")

        browse = QPushButton("Select VCF")
        browse.clicked.connect(self.select_file)

        top.addWidget(self.file_label)
        top.addWidget(browse)

        layout.addLayout(top)

        self.field_list = QListWidget()
        self.field_list.setSelectionMode(QAbstractItemView.MultiSelection)

        layout.addWidget(QLabel("Fields"))
        layout.addWidget(self.field_list)

        btns = QHBoxLayout()

        select_all = QPushButton("Select All")
        select_all.clicked.connect(self.select_all)

        clear = QPushButton("Clear")

        clear.clicked.connect(self.clear_all)

        export = QPushButton("Export Excel")
        export.clicked.connect(self.export_excel)

        btns.addWidget(select_all)
        btns.addWidget(clear)
        btns.addWidget(export)

        layout.addLayout(btns)

        self.table = QTableWidget()

        layout.addWidget(QLabel("Preview"))
        layout.addWidget(self.table)

        central.setLayout(layout)

    def select_file(self):

        path, _ = QFileDialog.getOpenFileName(self, "Select VCF", "", "VCF Files (*.vcf)")

        if not path:
            return

        self.file_label.setText(os.path.basename(path))

        self.worker = Worker(path)
        self.worker.finished.connect(self.loaded)

        self.worker.start()

    def loaded(self, contacts, fields):

        self.contacts = contacts
        self.fields = fields

        self.field_list.clear()

        for f in fields:

            item = QListWidgetItem(f)

            item.setSelected(True)

            self.field_list.addItem(item)

        self.preview()

    def preview(self):

        if not self.contacts:
            return

        sample = self.contacts[:20]

        keys = list(sample[0].keys())

        self.table.setColumnCount(len(keys))
        self.table.setRowCount(len(sample))

        self.table.setHorizontalHeaderLabels(keys)

        for r, contact in enumerate(sample):

            for c, k in enumerate(keys):

                self.table.setItem(r, c, QTableWidgetItem(contact.get(k, "")))

    def select_all(self):

        for i in range(self.field_list.count()):
            self.field_list.item(i).setSelected(True)

    def clear_all(self):

        for i in range(self.field_list.count()):
            self.field_list.item(i).setSelected(False)

    def export_excel(self):

        items = self.field_list.selectedItems()

        if not items:
            QMessageBox.warning(self, "Error", "Select fields first")
            return

        fields = [i.text() for i in items]

        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel (*.xlsx)")

        if not path:
            return

        ExcelExporter.export(self.contacts, fields, path)

        QMessageBox.information(self, "Done", "Export completed")


app = QApplication(sys.argv)

window = MainWindow()
window.show()

sys.exit(app.exec())

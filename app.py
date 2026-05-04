import os, sys
import sqlite3
import phonenumbers
import quopri
import openpyxl

from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

# --------------------------------
# GET BASE DIR LOCATION
# --------------------------------

def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)  # کنار EXE
    return os.path.dirname(os.path.abspath(__file__))  # حین development

# --------------------------------
# TEXT DECODER
# --------------------------------

def decode_vcard_text(value, encoding=None, charset=None):
    if not value:
        return ""

    charset = charset or "utf-8"

    try:
        if encoding and encoding.upper() == "QUOTED-PRINTABLE":
            decoded = quopri.decodestring(value.encode("ascii", errors="ignore"))
            return decoded.decode(charset, errors="ignore").strip()

        if encoding and encoding.upper() == "BASE64":
            return "[BASE64 DATA]"

        if "=" in value:
            decoded = quopri.decodestring(value.encode("ascii", errors="ignore"))
            return decoded.decode(charset, errors="ignore").strip()

    except Exception:
        pass

    try:
        return value.encode("latin-1").decode(charset, errors="ignore").strip()
    except Exception:
        return value.strip()


# --------------------------------
# PROPERTY PARSER
# --------------------------------

def parse_property(line):
    """
    Parse a vCard property line into:
      (group, key, params_dict, value)

    Example:
      item1.TEL;TYPE=CELL;ENCODING=QUOTED-PRINTABLE:+1234
      -> ("item1", "TEL", {"TYPE": "CELL", "ENCODING": "QUOTED-PRINTABLE"}, "+1234")
    """
    colon_idx = line.find(":")
    if colon_idx == -1:
        return None

    prop_part = line[:colon_idx]
    value = line[colon_idx + 1:]

    # split group prefix (item1.TEL -> group=item1, rest=TEL;...)
    group = ""
    if "." in prop_part.split(";")[0]:
        dot_idx = prop_part.index(".")
        group = prop_part[:dot_idx]
        prop_part = prop_part[dot_idx + 1:]

    segments = prop_part.split(";")
    key = segments[0].upper().strip()

    params = {}
    for seg in segments[1:]:
        if "=" in seg:
            pk, pv = seg.split("=", 1)
            params[pk.upper().strip()] = pv.strip()
        else:
            # bare param like BASE64 or QUOTED-PRINTABLE
            params[seg.upper().strip()] = seg.upper().strip()

    return group, key, params, value


# --------------------------------
# VCF LINE UNFOLDING (RFC6350)
# --------------------------------

def read_vcf_lines(path):
    with open(path, "r", encoding="utf8", errors="ignore") as f:
        raw = f.readlines()

    lines = []
    for line in raw:
        line = line.rstrip("\r\n")

        # RFC 6350 whitespace folding: lines starting with space/tab
        if line.startswith((" ", "\t")) and lines:
            lines[-1] += line[1:]  # strip exactly one leading whitespace char

        # vCard 2.1 QP soft line break: previous line ends with "="
        elif lines and lines[-1].endswith("="):
            lines[-1] = lines[-1][:-1] + line.strip()

        else:
            lines.append(line)

    return lines


# --------------------------------
# PHONE INTEL
# --------------------------------

class PhoneIntel:

    @staticmethod
    def analyze(number):
        try:
            parsed = phonenumbers.parse(number, None)
            country = phonenumbers.region_code_for_number(parsed)
            line = phonenumbers.number_type(parsed)
            formatted = phonenumbers.format_number(
                parsed,
                phonenumbers.PhoneNumberFormat.E164
            )
            return formatted, country, str(line)
        except Exception:
            return number, "", ""


# --------------------------------
# DATABASE
# --------------------------------

class ContactDB:
    def __init__(self):
        path = os.path.join(get_base_dir(), "contacts.db")
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.create()

    def create(self):
        cur = self.conn.cursor()

        cur.execute("""
        CREATE TABLE IF NOT EXISTS contacts(
            id         INTEGER PRIMARY KEY,
            name       TEXT,
            phone      TEXT,
            email      TEXT,
            country    TEXT,
            line_type  TEXT,
            phone_type TEXT,
            address    TEXT
        )
        """)

        # migrate old db: add columns if missing
        existing = {row[1] for row in cur.execute("PRAGMA table_info(contacts)")}

        if "phone_type" not in existing:
            cur.execute("ALTER TABLE contacts ADD COLUMN phone_type TEXT DEFAULT ''")

        if "address" not in existing:
            cur.execute("ALTER TABLE contacts ADD COLUMN address TEXT DEFAULT ''")

        self.conn.commit()


    def insert(self, name, phone, email, country, line, phone_type="", address=""):
        cur = self.conn.cursor()
        cur.execute(
            """INSERT INTO contacts
               (name, phone, email, country, line_type, phone_type, address)
               VALUES (?,?,?,?,?,?,?)""",
            (name, phone, email, country, line, phone_type, address)
        )

    def delete_contact(self, contact_id):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM contacts WHERE id=?", (contact_id,))
        self.conn.commit()

    def fetch_all(self):
        cur = self.conn.cursor()
        cur.execute(
            "SELECT id,name,phone,email,country,line_type,phone_type,address FROM contacts"
        )
        return cur.fetchall()


# --------------------------------
# EXPORT DIALOG
# --------------------------------

class ExportDialog(QDialog):

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Select Columns")
        self.columns = ["id", "name", "phone", "email", "country",
                        "line_type", "phone_type", "address"]
        layout = QVBoxLayout()
        self.checks = []

        for c in self.columns:
            chk = QCheckBox(c)
            chk.setChecked(True)
            self.checks.append(chk)
            layout.addWidget(chk)

        btn = QPushButton("Export Excel")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn)
        self.setLayout(layout)

    def selected(self):
        return [c for c, chk in zip(self.columns, self.checks) if chk.isChecked()]


# --------------------------------
# EXPORT ENGINE
# --------------------------------

def export_excel(db, columns):
    path, _ = QFileDialog.getSaveFileName(
        None, "Save Excel", "", "Excel File (*.xlsx)"
    )
    if not path:
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(columns)

    cur = db.conn.cursor()
    query = "SELECT {} FROM contacts".format(",".join(columns))
    for row in cur.execute(query):
        ws.append(list(row))

    wb.save(path)


# --------------------------------
# NAME BUILDER
# --------------------------------

def build_name(fn_value, n_value, encoding=None, charset=None):
    """
    Use FN if available, otherwise reconstruct from N field.
    N format: Last;First;Middle;Prefix;Suffix
    """
    name = decode_vcard_text(fn_value, encoding, charset).strip()
    if name:
        return name

    if n_value:
        parts = n_value.split(";")
        # order: Prefix First Middle Last Suffix
        order = [3, 1, 2, 0, 4]
        assembled = " ".join(
            decode_vcard_text(parts[i], encoding, charset).strip()
            for i in order
            if i < len(parts) and parts[i].strip()
        )
        return assembled.strip()

    return ""


# --------------------------------
# ADDRESS BUILDER
# --------------------------------

def build_address(value, encoding=None, charset=None):
    """
    ADR format: POBox;Extended;Street;City;State;ZIP;Country
    """
    parts = value.split(";")
    labels = ["POBox", "Extended", "Street", "City", "State", "ZIP", "Country"]
    result = []
    for i, label in enumerate(labels):
        if i < len(parts):
            v = decode_vcard_text(parts[i], encoding, charset).strip()
            if v:
                result.append(v)
    return ", ".join(result)


# --------------------------------
# VCF PARSER
# --------------------------------

class VCFParser:

    def __init__(self, files, db):
        self.files = files
        self.db = db
        self.errors = []

    def parse(self):
        for file in self.files:
            self._parse_file(file)
        self.db.conn.commit()

    def _empty_contact(self):
        return {
            "FN": ("", {}),
            "N": ("", {}),
            "TEL": [],        # list of (value, params)
            "EMAIL": [],      # list of (value, params)
            "ADR": [],        # list of (value, params)
        }

    def _parse_file(self, path):
        lines = read_vcf_lines(path)
        contact = None
        in_card = False

        for line in lines:
            if line == "BEGIN:VCARD":
                contact = self._empty_contact()
                in_card = True
                continue

            if line == "END:VCARD":
                if contact is not None:
                    self._flush(contact)
                contact = None
                in_card = False
                continue

            if not in_card or not line:
                continue

            parsed = parse_property(line)
            if not parsed:
                continue

            _group, key, params, value = parsed

            if key == "TEL":
                contact["TEL"].append((value, params))

            elif key == "EMAIL":
                contact["EMAIL"].append((value, params))

            elif key == "ADR":
                contact["ADR"].append((value, params))

            elif key in ("FN", "N"):
                contact[key] = (value, params)

        # detect truncated vCard
        if in_card and contact is not None:
            self.errors.append(f"{path}: truncated vCard (missing END:VCARD)")

    def _flush(self, contact):
        fn_val, fn_params = contact["FN"]
        n_val, n_params = contact["N"]

        # use FN params for encoding/charset hints, fallback to N params
        enc = fn_params.get("ENCODING") or n_params.get("ENCODING")
        cset = fn_params.get("CHARSET") or n_params.get("CHARSET")

        name = build_name(fn_val, n_val, enc, cset)

        # --- emails: decode, deduplicate, join ---
        seen_emails = set()
        email_list = []
        for e_val, e_params in contact["EMAIL"]:
            e_enc = e_params.get("ENCODING")
            e_cset = e_params.get("CHARSET")
            decoded_email = decode_vcard_text(e_val, e_enc, e_cset).strip().lower()
            if decoded_email and decoded_email not in seen_emails:
                seen_emails.add(decoded_email)
                email_list.append(decoded_email)
        emails = ",".join(email_list)

        # --- address: take first ADR ---
        address = ""
        if contact["ADR"]:
            a_val, a_params = contact["ADR"][0]
            a_enc = a_params.get("ENCODING")
            a_cset = a_params.get("CHARSET")
            address = build_address(a_val, a_enc, a_cset)

        # --- phones: decode, normalize, deduplicate ---
        phones = contact["TEL"] if contact["TEL"] else [("", {})]
        seen_phones = set()

        for t_val, t_params in phones:
            t_enc = t_params.get("ENCODING")
            t_cset = t_params.get("CHARSET")
            phone = decode_vcard_text(t_val, t_enc, t_cset).strip()

            phone_type = t_params.get("TYPE", "")

            phone_e164, country, line_type = PhoneIntel.analyze(phone)

            # deduplicate by normalized E164 (or raw if parse failed)
            dedup_key = phone_e164.strip() or phone.strip()
            if dedup_key and dedup_key in seen_phones:
                continue
            if dedup_key:
                seen_phones.add(dedup_key)

            self.db.insert(
                name,
                phone_e164,
                emails,
                country,
                line_type,
                phone_type,
                address
            )


# --------------------------------
# TABLE MODEL
# --------------------------------

class ContactTableModel(QAbstractTableModel):

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.headers = ["ID", "Name", "Phone", "Email",
                        "Country", "Line", "Phone Type", "Address"]
        self.cache = []
        self.load()

    def load(self):
        self.cache = self.db.fetch_all()

    def rowCount(self, parent=QModelIndex()):
        return len(self.cache)

    def columnCount(self, parent=QModelIndex()):
        return len(self.headers)

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self.cache[index.row()][index.column()])

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.headers[section]


# --------------------------------
# TABLE VIEW DELETE KEY
# --------------------------------

class ContactTableView(QTableView):

    def __init__(self, parent):
        super().__init__()
        self.main = parent
        self.setSelectionBehavior(QTableView.SelectRows)
        self.setSelectionMode(QTableView.ExtendedSelection)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            rows = self.selectionModel().selectedRows()
            if not rows:
                return

            confirm = QMessageBox.question(
                self,
                "Delete",
                f"Delete {len(rows)} contacts?",
                QMessageBox.Yes | QMessageBox.No
            )
            if confirm != QMessageBox.Yes:
                return

            for r in rows:
                source_row = self.main.proxy.mapToSource(r).row()
                cid = self.main.model.cache[source_row][0]
                self.main.db.delete_contact(cid)

            self.main.refresh()
        else:
            super().keyPressEvent(event)


# --------------------------------
# MAIN WINDOW
# --------------------------------

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.db = ContactDB()
        self.setWindowTitle("ARUSH - VCF Contact Manager")
        self.resize(1100, 700)
        self.setWindowIcon(QIcon("icon.png"))
        self.build()

    def build(self):
        root = QWidget()
        layout = QVBoxLayout()

        search_layout = QHBoxLayout()

        self.search = QLineEdit()
        self.search.setPlaceholderText("Search contacts...")

        self.column_filter = QComboBox()
        self.column_filter.addItem("All")
        self.column_filter.addItems(["ID", "Name", "Phone", "Email",
                                     "Country", "Line", "Phone Type", "Address"])

        search_layout.addWidget(QLabel("Search"))
        search_layout.addWidget(self.search)
        search_layout.addWidget(self.column_filter)

        layout.addLayout(search_layout)

        self.table = ContactTableView(self)
        self.model = ContactTableModel(self.db)

        self.proxy = QSortFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)

        self.table.setModel(self.proxy)
        layout.addWidget(self.table)

        self.search.textChanged.connect(self.apply_filter)
        self.column_filter.currentIndexChanged.connect(self.apply_filter)

        btns = QHBoxLayout()

        load = QPushButton("Load VCF")
        load.clicked.connect(self.load_vcf)

        export = QPushButton("Export Excel")
        export.clicked.connect(self.export_excel)

        btns.addWidget(load)
        btns.addWidget(export)

        layout.addLayout(btns)

        root.setLayout(layout)
        self.setCentralWidget(root)

    def apply_filter(self):
        text = self.search.text()
        col = self.column_filter.currentIndex() - 1

        if col < 0:
            self.proxy.setFilterKeyColumn(-1)
        else:
            self.proxy.setFilterKeyColumn(col)

        self.proxy.setFilterFixedString(text)

    def load_vcf(self):
        files, _ = QFileDialog.getOpenFileNames(self, "", "", "*.vcf")
        if not files:
            return

        parser = VCFParser(files, self.db)
        parser.parse()

        if parser.errors:
            QMessageBox.warning(
                self,
                "Parse Warnings",
                "\n".join(parser.errors)
            )

        self.refresh()

    def refresh(self):
        self.model.load()
        self.model.layoutChanged.emit()

    def export_excel(self):
        dlg = ExportDialog(self)
        if dlg.exec():
            cols = dlg.selected()
            export_excel(self.db, cols)


# --------------------------------
# RUN
# --------------------------------

app = QApplication(sys.argv)
win = MainWindow()
win.show()
sys.exit(app.exec())

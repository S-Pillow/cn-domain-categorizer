#!/usr/bin/env python3
# CN Domain Categorizer GUI
# Version: 2025-05-14  (handles xj.cn, gs.cn, etc.; popup text selectable)

import sys, re, datetime, pandas as pd, tldextract, idna
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QProgressBar, QMessageBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

COLUMN = "Domain Name"

# ─────────────────── bucket definitions ────────────────────────────────────
BUCKETS = [
    "IDN.IDN", "IDN.XN--FIQS8S", "ASCII.XN--FIQS8S",
    "ASCII.CN", "IDN.CN", ".COM.CN", ".NET.CN", ".ORG.CN"
]
ALL_BUCKETS = BUCKETS + ["UNCLASSIFIED"]

PUNY = re.compile(r"^xn--", re.I)
PUNY_TLDS = {"xn--fiqs8s", "xn--fiqz9s"}             # .中国 / .中國
GENERIC_SLD = {"com.cn": ".COM.CN", "net.cn": ".NET.CN", "org.cn": ".ORG.CN"}


def is_idn(label: str) -> bool:
    return (not label.isascii()) or bool(PUNY.match(label))


def classify(domain: str) -> str:
    ext = tldextract.extract(domain.lower())
    sld, suffix = ext.domain, ext.suffix   # suffix may be cn, com.cn, xj.cn, etc.

    # 1. Generic second-level zones under .cn
    if suffix in GENERIC_SLD:
        return GENERIC_SLD[suffix]

    # 2. Puny-code Chinese TLDs (.中国 / .中國)
    if suffix in PUNY_TLDS:
        return "IDN.XN--FIQS8S" if is_idn(sld) else "ASCII.XN--FIQS8S"

    # 3. Any suffix that ends with ".cn" (plain 'cn' *or* province SLDs like xj.cn)
    if suffix == "cn" or suffix.endswith(".cn"):
        return "IDN.CN" if is_idn(sld) else "ASCII.CN"

    # 4. Edge-case: fully-Unicode domain + Unicode TLD
    try:
        ace = idna.encode(domain).decode()
        ext2 = tldextract.extract(ace)
        if ext2.suffix in PUNY_TLDS and is_idn(ext2.domain):
            return "IDN.IDN"
    except idna.IDNAError:
        pass

    return "UNCLASSIFIED"


def multi_sheet_write(df: pd.DataFrame, dest: Path):
    with pd.ExcelWriter(dest, engine="openpyxl") as writer:
        for bucket in ALL_BUCKETS:
            sub = df[df["bucket"] == bucket]
            if not sub.empty:
                sub.to_excel(writer, sheet_name=bucket[:31], index=False)


# ─────────────────── background worker ─────────────────────────────────────
class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(Path, str, dict)   # outfile, error, counts

    def __init__(self, infile: Path, outfile: Path):
        super().__init__()
        self.infile, self.outfile = infile, outfile

    def run(self):
        try:
            df = (pd.read_csv if self.infile.suffix.lower() == ".csv" else pd.read_excel)(self.infile)
            if COLUMN not in df.columns:
                raise ValueError(f"Input must contain a '{COLUMN}' column.")

            total = len(df)
            df["bucket"] = ""
            for i, dom in enumerate(df[COLUMN].astype(str).str.strip()):
                df.at[i, "bucket"] = classify(dom)
                if i % 50 == 0:
                    self.progress.emit(int(i / max(total, 1) * 100))
            self.progress.emit(100)

            multi_sheet_write(df, self.outfile)

            counts = df["bucket"].value_counts().to_dict()
            counts["TOTAL"] = total
            self.finished.emit(self.outfile, "", counts)

        except Exception as exc:
            self.finished.emit(self.outfile, str(exc), {})


# ─────────────────── main window ───────────────────────────────────────────
class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CN Domain Categorizer")
        self.setFixedWidth(540)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        # input
        row = QHBoxLayout()
        self.in_edit = QLineEdit(); self.in_edit.setPlaceholderText(f"Spreadsheet with '{COLUMN}' column …")
        row.addWidget(QLabel("Input file:")); row.addWidget(self.in_edit)
        btn = QPushButton("Browse…"); btn.clicked.connect(self._pick_in); row.addWidget(btn)
        layout.addLayout(row)

        # output
        row = QHBoxLayout()
        self.out_edit = QLineEdit(); self.out_edit.setPlaceholderText("Auto-generated if blank")
        row.addWidget(QLabel("Output file:")); row.addWidget(self.out_edit)
        btn = QPushButton("Browse…"); btn.clicked.connect(self._pick_out); row.addWidget(btn)
        layout.addLayout(row)

        self.bar = QProgressBar(); self.bar.setAlignment(Qt.AlignCenter)
        self.run_btn = QPushButton("Run"); self.run_btn.clicked.connect(self._run)
        layout.addWidget(self.bar); layout.addWidget(self.run_btn)

    # file pickers ----------------------------------------------------------
    def _pick_in(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select spreadsheet", str(Path.home()), "Spreadsheets (*.xlsx *.xls *.csv)"
        )
        if path:
            self.in_edit.setText(path); self._suggest_out(Path(path))

    def _pick_out(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save workbook", str(Path.home()), "Excel (*.xlsx)"
        )
        if path: self.out_edit.setText(path)

    def _suggest_out(self, infile: Path):
        today = datetime.date.today().strftime("%Y%m%d")
        name = f"sorted_CN_domains_{today}.xlsx"
        target = Path.home()/ "Downloads" if (Path.home()/ "Downloads").exists() else Path.home()
        self.out_edit.setText(str(target / name))

    # run job ---------------------------------------------------------------
    def _run(self):
        in_path = Path(self.in_edit.text().strip())
        if not in_path.exists():
            QMessageBox.warning(self, "Error", "Select a valid input file."); return

        out_path = Path(self.out_edit.text().strip()) if self.out_edit.text().strip() else None
        if out_path is None:
            self._suggest_out(in_path); out_path = Path(self.out_edit.text())

        self.run_btn.setEnabled(False)
        self.worker = Worker(in_path, out_path)
        self.worker.progress.connect(self.bar.setValue)
        self.worker.finished.connect(self._done)
        self.worker.start()

    # completion popup ------------------------------------------------------
    def _done(self, outfile: Path, error: str, counts: dict):
        self.run_btn.setEnabled(True); self.bar.setValue(0)
        if error:
            QMessageBox.critical(self, "Failed", error); return

        total = counts.get("TOTAL", 0)
        lines = [
            f"Workbook saved to:\n{outfile}\n",
            f"Total processed: {total:,}\n"
        ]
        breakdown = [f"{b:<22} {counts.get(b,0):,}" for b in ALL_BUCKETS if counts.get(b,0)]
        if breakdown:
            lines.append("Break-down\n" + "─"*10)
            lines.extend(breakdown)
        summary = "\n".join(lines)

        box = QMessageBox(self)
        box.setWindowTitle("Done")
        box.setText(summary)
        box.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        box.exec_()

        if sys.platform.startswith("win"):
            import subprocess, os
            subprocess.Popen(rf'explorer /select,"{os.path.normpath(outfile)}"')

# ─────────────────── bootstrap ─────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = App(); win.show()
    sys.exit(app.exec_())

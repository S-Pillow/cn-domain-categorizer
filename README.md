# CN Domain Categorizer GUI 🀄️

A tiny PyQt 5 application that takes a list of Chinese‐namespace domains and
spits out a multi-sheet Excel workbook, classifying each name into eight
buckets (plus *UNCLASSIFIED* catch-all):

* **IDN.IDN**  
* **IDN.XN--FIQS8S** – IDN label under `xn--fiqs8s` (“.中国 / .中國”)  
* **ASCII.XN--FIQS8S**  
* **ASCII.CN** – includes province SLDs like `xj.cn`, `qh.cn`, …  
* **IDN.CN**  
* **.COM.CN / .NET.CN / .ORG.CN**

![screenshot](docs/screenshot.png)
*(optional — add one!)*

---

## Features

| ✔ | What it does |
|---|--------------|
| GUI file picker | Accepts **.xlsx, .xls, or .csv** input (column `Domain Name`). |
| Smart rules     | Detects Punycode, province second-level zones, generic SLDs. |
| Excel output    | One sheet per bucket, generated with *openpyxl*. |
| Progress bar    | Updates every 50 rows; shows per-bucket counts at the end. |
| Copyable report | Completion popup text is selectable for quick copy-paste. |

---

## Installation

```bash
# Windows – same interpreter Thonny uses
python -m pip install PyQt5 pandas openpyxl tldextract idna

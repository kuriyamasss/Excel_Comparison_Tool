# Excel Comparison Tool

A local Excel/CSV comparison tool for inventory or part-number datasets.  
Built with **Python + Flask**, providing a browser-based interface for comparing two tables (old vs. new) and exporting intersection and differences.

Supports `.xlsx`, `.xls`, `.csv`, multi-sheet Excel files, header auto-detection, duplicate key handling, and multilingual UI (Simplified Chinese & Vietnamese).

---

## 🚀 Features

- Upload **Old Table** and **New Table** (Excel or CSV)
- Auto-detect or manually specify:
  - Sheet (工作表)
  - Header row (表头)
- Select comparison key (e.g., 料号 / Part Number)
- Choose duplicate-key strategy:
  - Keep first
  - Keep last
  - Error on duplicates
- Export an Excel file containing:
  - `intersection`（两表共有行）
  - `only_in_old`（仅旧表有）
  - `only_in_new`（仅新表有）
- Loading spinner & progress indicator
- Multilingual UI: 简体中文 (default) / Tiếng Việt
- Packaged into a **single-file Windows EXE**

---

## 📦 How to Use (EXE Version)

1. Download the latest EXE from **Releases**.
2. Double-click the `.exe`.
3. The tool will automatically open your default browser at:
    http://127.0.0.1:5000/
4. Upload the old/new tables → select sheet & header → choose key → generate output.
5. Download the generated comparison report.

> Note: Windows SmartScreen may warn about unsigned executables. Choose “Run anyway” if used in a trusted environment.

---

## Run from Source (Python)

### 1. Clone the repository
```bash
git clone https://github.com/YOUR_NAME/Excel_Comparison_Tool.git
cd Excel_Comparison_Tool
```

### 2. (Recommended) Create virtual environment
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
# source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```
or:
```bash
pip install flask pandas openpyxl
```
### 4. Run the application
```bash
python compare_tool.py
```
Visit:
```bash
http://127.0.0.1:5000/
```

## Example Output
Generated Excel will contain multiple sheets:
| Sheet Name   | Meaning                      |
| ------------ | ---------------------------- |
| intersection | Rows present in both tables  |
| only_in_old  | Rows unique to the old table |
| only_in_new  | Rows unique to the new table |

File naming example:
```text
compare_20251204_145230_key_料号_inter12_old5_new7.xlsx
```

## Packaging to EXE (PyInstaller)

Test (onedir)::
```bash
pyinstaller --onedir --clean compare_tool.py
```

Final build (onefile):
```bash
pyinstaller --onefile --clean ^
  --hidden-import=openpyxl --hidden-import=pandas ^
  --add-data "venv\Lib\site-packages\openpyxl;openpyxl" ^
  compare_tool.py
```
EXE will appear under dist/.

## Project Structure

```cpp
Excel_Comparison_Tool/
│── compare_tool.py
│── README.md
│── requirements.txt
│── static/
│── templates/
└── venv/ (ignored)

```

## Notes & Limitations

· Very large datasets (> 100k–500k rows) depend on available RAM.

· Out-of-memory conditions may occur with Excel; CSV recommended for larger workloads.

· Unsigned EXE may trigger antivirus or SmartScreen warnings.

## License

This project is released under the MIT License.

## Contact

Author: kuriyamasss
For issues, suggestions, or feature requests, please submit an Issue on GitHub.
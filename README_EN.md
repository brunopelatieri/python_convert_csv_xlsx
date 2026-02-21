# 📋 Chatwoot CSV → Excel Converter

> Converts **Chatwoot** contact exports from CSV to a formatted Excel spreadsheet (`.xlsx`), automatically fixing broken special characters like accents and cedillas.

---

## 🔍 The Problem

When exporting contacts from Chatwoot, the generated CSV file often displays corrupted special characters:

| ❌ Broken | ✅ Fixed |
|---|---|
| `AndrÃ© Camily` | `André Camily` |
| `ElisÃ¢ngela` | `Elisângela` |
| `JoÃ£o da Silva` | `João da Silva` |
| `ApareciÃ§a` | `Aparecição` |

This happens due to an **encoding conflict** — the file is saved in UTF-8 by Chatwoot, but many programs (like Excel) open it assuming Latin-1/ISO-8859-1, or vice-versa.

---

## ✅ What the Script Does

- 🔎 **Auto-detects** the correct CSV encoding (tests utf-8, utf-8-sig, latin-1, iso-8859-1, cp1252)
- 🔧 **Fixes** corrupted special characters (accents, ç, ã, ê, etc.)
- 📊 **Converts** the CSV to `.xlsx` with professional formatting:
  - Header row with blue background and bold white text
  - Column widths automatically adjusted to content
  - Standardized Arial font
- 📁 **Names** the output file automatically (same name as the CSV, `.xlsx` extension)

---

## 🚀 How to Use

### 1. Prerequisites

Make sure you have **Python 3.7+** installed:

```bash
python --version
# or
python3 --version
```

> Don't have Python? Download it at [python.org](https://www.python.org/downloads/)

### 2. Install dependencies

```bash
pip install pandas openpyxl
```

### 3. Download the script

Clone the repository or download `convert.py` directly.

```bash
git clone https://github.com/your-username/chatwoot-csv-excel.git
cd chatwoot-csv-excel
```

### 4. Export contacts from Chatwoot

In Chatwoot, go to: **Contacts → Import/Export → Export Contacts**

The file will be downloaded as a `.csv`.

### 5. Run the script

**Basic usage** (generates `.xlsx` with the same name as the CSV):
```bash
python convert.py contacts.csv
```

**Specifying a custom output filename:**
```bash
python convert.py contacts.csv my_spreadsheet.xlsx
```

---

## 💻 Examples by Operating System

### Windows (Command Prompt)

```cmd
# Navigate to the folder containing the script and CSV
cd C:\Users\YourName\Downloads

# Run
python convert.py contacts.csv
```

### macOS / Linux (Terminal)

```bash
# Navigate to the folder containing the script and CSV
cd ~/Downloads

# Run
python3 convert.py contacts.csv
```

> 💡 **Tip:** Place the script in the same folder as the CSV file to keep things simple — no need to type long file paths.

---

## 📂 Repository Structure

```
chatwoot-csv-excel/
│
├── convert.py     # Main script
├── README.md                 # This documentation
└── contacts_example.csv  # Sample file for testing
```

---

## ⚙️ How It Works (Technical Details)

The script follows three main steps:

**1. Encoding Detection**
Tests a prioritized list of encodings until the file can be read without errors:
```
utf-8 → utf-8-sig → latin-1 → iso-8859-1 → cp1252
```

**2. Character Correction**
Applies a `latin-1 → utf-8` re-encoding column by column to reverse the mojibake (the technical term for garbled characters caused by encoding mismatches). Columns that are already correct are automatically skipped.

**3. Excel Export**
Uses `pandas` + `openpyxl` to generate the `.xlsx` with:
- Sheet named `Contacts`
- Formatted header row (blue background `#2B5ED6`, white bold text)
- Column widths proportional to content (capped at 40 characters)

---

## 🐛 Troubleshooting

**`python: command not found`**
> Use `python3` instead of `python`, or verify that Python is installed and added to your system PATH.

**`ModuleNotFoundError: No module named 'pandas'`**
> Run: `pip install pandas openpyxl`

**Characters still appear broken after conversion**
> The file may use a less common encoding. Open the CSV in a text editor (such as VS Code or Notepad++) and check the encoding displayed in the status bar. Adjust the `encodings` list in the script accordingly.

**Error opening `.xlsx` in Excel**
> Make sure the file is not already open in another program during the conversion.

---

## 🛠️ Technologies Used

| Library | Min. Version | Purpose |
|---|---|---|
| Python | 3.7+ | Base language |
| pandas | 1.3+ | CSV reading and data manipulation |
| openpyxl | 3.0+ | Excel file generation and formatting |

---

## 🤝 Contributing

Contributions are welcome! Feel free to:

1. **Fork** the repository
2. Create a **branch** for your feature: `git checkout -b my-improvement`
3. **Commit** your changes: `git commit -m 'Add support for X'`
4. **Push** to the branch: `git push origin my-improvement`
5. Open a **Pull Request**

---

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

---

## 📬 Issues & Feedback

Have a question or suggestion? Open an [issue](../../issues) in the repository.

---

## 👤 Author

**Bruno Pelatieri Goulart**
- 🌐 [brunogoulart.com.br](https://brunogoulart.com.br)
- 🤖 [bru.ia.br](https://bru.ia.br)

---

*Built to make life easier for Chatwoot and CSV users* 🚀

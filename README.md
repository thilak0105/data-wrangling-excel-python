# 📊 Python Excel Wrangling

A comprehensive Jupyter notebook demonstrating various Python libraries for reading, writing, and manipulating Excel files. This project explores multiple approaches to Excel data handling, from legacy formats (.xls) to modern standards (.xlsx).

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Jupyter](https://img.shields.io/badge/Jupyter-Notebook-orange.svg)](https://jupyter.org/)

## 🎯 Overview

This repository contains hands-on examples of working with Excel files in Python using multiple libraries:

- **xlrd/xlwt** - Reading and writing legacy `.xls` files
- **openpyxl** - Modern `.xlsx` file manipulation
- **python-calamine** - Fast Rust-based Excel reading
- **xlutils** - Utilities for working with xlrd/xlwt

The notebook includes real-world datasets including corruption perception indices and health life expectancy data from WHO, demonstrating practical data wrangling techniques.

## 🚀 Features

- ✅ Read legacy `.xls` files using xlrd
- ✅ Read/write modern `.xlsx` files with openpyxl
- ✅ High-performance Excel reading with python-calamine
- ✅ Practical examples with real-world datasets
- ✅ Cross-library comparison and use cases
- ✅ Data extraction and row-level processing

## 📋 Prerequisites

- Python 3.10 or higher
- Jupyter Notebook or JupyterLab
- Basic understanding of Python and pandas

## 🔧 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/python-excel-wrangling.git
cd python-excel-wrangling
```

### 2. Create Virtual Environment (Recommended)

```bash
# Using venv
python -m venv venv

# Activate on macOS/Linux
source venv/bin/activate

# Activate on Windows
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Launch Jupyter Notebook

```bash
jupyter notebook Excel.ipynb
```

## 📦 Dependencies

```
xlrd>=2.0.0
xlwt>=1.3.0
xlutils>=2.0.0
openpyxl>=3.1.0
python-calamine>=0.2.0
jupyter>=1.0.0
```

## 📁 Project Structure

```
python-excel-wrangling/
│
├── Excel.ipynb              # Main notebook with all examples
├── README.md                # This file
├── requirements.txt         # Python dependencies
├── LICENSE                  # MIT License
│
└── data/                    # Sample datasets (you'll need to add these)
    ├── corruption_perception_index.xls
    └── data-text.xlsx
```

## 💡 Usage Examples

### Reading Legacy .xls Files

```python
import xlrd

# Open workbook
xls_file = "data/corruption_perception_index.xls"
workbook = xlrd.open_workbook(xls_file)
sheet = workbook.sheet_by_index(0)

# Read data
for row_idx in range(sheet.nrows):
    print(sheet.row_values(row_idx))
```

### Reading Modern .xlsx Files with python-calamine

```python
from python_calamine import CalamineWorkbook

# Fast Excel reading
wb = CalamineWorkbook.from_path("data/data-text.xlsx")
sheet_names = wb.sheet_names
data = wb.get_sheet_by_name(sheet_names[0]).to_python()

for row in data:
    print(row)
```

### Working with openpyxl

```python
from openpyxl import load_workbook

# Load and manipulate .xlsx files
wb = load_workbook("data/data-text.xlsx")
ws = wb.active

# Iterate through rows
for row in ws.iter_rows(values_only=True):
    print(row)
```

## 🎓 Learning Objectives

This notebook is designed to help you:

1. **Understand Excel file formats** - Differences between .xls (BIFF) and .xlsx (OOXML)
2. **Choose the right library** - When to use xlrd vs openpyxl vs python-calamine
3. **Handle real-world data** - Working with messy, multi-sheet Excel files
4. **Performance optimization** - Comparing library speeds for large datasets
5. **Data extraction patterns** - Best practices for reading and processing Excel data

## 📊 Datasets

The notebook works with the following datasets:

### 1. Corruption Perception Index (.xls)
- **Format**: Legacy Excel format
- **Library Used**: xlrd
- **Purpose**: Demonstrates reading older Excel files

### 2. WHO Health Data (.xlsx)
- **Fields**: Indicator, Publication Status, Year, Region, Income Group, Country, Sex, Life Expectancy
- **Rows**: ~8000+ records
- **Library Used**: python-calamine, openpyxl
- **Purpose**: Real-world health statistics processing

> **Note**: You'll need to provide your own datasets or update the file paths in the notebook to point to your data.

## 🔍 Library Comparison

| Library | Format | Speed | Features | Use Case |
|---------|--------|-------|----------|----------|
| **xlrd** | .xls | Fast | Read-only | Legacy Excel files |
| **xlwt** | .xls | Medium | Write-only | Creating .xls files |
| **openpyxl** | .xlsx | Medium | Read/Write | Modern Excel with formatting |
| **python-calamine** | .xlsx, .xls, .xlsb | Very Fast | Read-only | High-performance data extraction |

## ⚡ Performance Tips

1. **For large .xlsx files**: Use `python-calamine` for fastest reading
2. **For .xls files**: Stick with `xlrd` (v2.0+ only supports .xls)
3. **For writing**: Use `openpyxl` for .xlsx, `xlwt` for .xls
4. **For formatting**: `openpyxl` provides the most comprehensive styling options

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Contribution Ideas
- Add examples with pandas integration
- Include write operations for all libraries
- Add performance benchmarks
- Create visualizations of the datasets
- Add error handling examples

## 🐛 Known Issues

- Hardcoded file paths need to be updated for your local environment
- Sample data rows (8160-8165) contain duplicates
- No output sanitization in the notebook

## 📝 TODO

- [ ] Add `requirements.txt` file
- [ ] Parameterize file paths
- [ ] Add pandas integration examples
- [ ] Include write operations
- [ ] Add unit tests
- [ ] Create performance benchmarks
- [ ] Add data validation examples
- [ ] Include error handling patterns

## 📚 Resources

- [xlrd Documentation](https://xlrd.readthedocs.io/)
- [openpyxl Documentation](https://openpyxl.readthedocs.io/)
- [python-calamine GitHub](https://github.com/dimastbk/python-calamine)
- [Working with Excel Files in Python - Real Python](https://realpython.com/openpyxl-excel-spreadsheets-python/)

## 👨‍💻 Author

**Thilak**  
M.Sc. Data Science Student  
Amrita Vishwa Vidyapeetham  

[![GitHub](https://img.shields.io/badge/GitHub-Profile-black?logo=github)](https://github.com/yourusername)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?logo=linkedin)](https://linkedin.com/in/yourprofile)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Data Wrangling Course - Amrita School of Physical Sciences
- WHO for health statistics datasets
- Transparency International for corruption index data
- Python community for excellent Excel handling libraries

---

<div align="center">

**⭐ Star this repo if you found it helpful!**

Made with ❤️ for the Data Science community

</div>

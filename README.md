Swisstek Payslip Generator

A desktop application that automates monthly payslip generation for Swisstek Aluminium Limited employees — reads raw payroll data from Excel/CSV and produces individual, formula-driven payslips in Excel format for every employee in seconds.


 Features

- 🖥️ **Modern desktop GUI** — dark-mode interface built with CustomTkinter
- 📊 **Flexible input** — reads payroll data from both `.xlsx` and `.csv` sources
- 📄 **Template-driven output** — generates one payslip per employee from a formatted Excel template, preserving all built-in formulas (EPF, ETF, Gross Salary, Net Salary auto-calculate)
- 📈 **Live progress tracking** — real-time progress bar and status updates during generation
- 📁 **One-click output access** — automatically opens the output folder when generation completes
- ⚡ **Bulk generation** — processes an entire company payroll sheet (any number of employees) in one run

 Tech Stack

- **Python 3** — core application logic
- **pandas** — payroll data parsing (Excel/CSV)
- **openpyxl** — Excel template reading and writing
- **CustomTkinter** — desktop GUI

 Project Structure
 PayslipGenerator
├── input/              # Raw payroll data (xlsx/csv)
├── output/             # Generated payslips
├── templates/           # Payslip Excel template
├── assets/              # Logo, icons
├── src/
│   ├── models/          # Employee data model
│   ├── services/        # Excel reading & payslip generation logic
│   ├── ui/               # CustomTkinter GUI
│   └── config.py
├── main.py               # Application entry point
└── requirements.txt


 Setup

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

 How It Works

1. Launch the app and browse to select the payroll data file, the payslip template, and an output folder.
2. The app reads and maps 30+ payroll fields per employee (earnings, deductions, incentives, statutory contributions).
3. Click **Generate Payslips** — a formatted, formula-driven `.xlsx` payslip is created for every employee.
4. The output folder opens automatically once generation is complete.

 Roadmap

- [ ] PDF export per payslip
- [ ] Email delivery to employees
- [ ] Standalone `.exe` build (no Python installation required)


Built as part of an internship project at Swisstek Aluminium Limited.


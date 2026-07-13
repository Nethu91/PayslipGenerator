# PayslipGenerator
A Python-based payroll payslip generator that automatically reads employee payroll data from Excel and generates formatted individual payslips using a customizable Excel template.




 Features

- Read employee payroll data from Excel (.xlsx)
- Automatically generate individual payslips
- Bulk payslip generation for all employees
- Customizable Excel payslip template
- No manual calculations required
- Simple and easy-to-maintain code structure
- Fast payroll processing



 Technologies Used

- Python 3
- Pandas
- OpenPyXL
- ReportLab (Optional)
- Pillow


 Project Structure


PayslipGenerator/
│
├── input/
│   └── Payslips.xlsx
│
├── output/
│   └── Generated Payslips
│
├── templates/
│   └── Payslip_Template.xlsx
│
├── src/
│   ├── models/
│   ├── services/
│   └── app.py
│
├── main.py
├── requirements.txt
└── README.md




 Installation

Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/PayslipGenerator.git
```

Go to the project directory

```bash
cd PayslipGenerator
```

Create a virtual environment

```bash
python -m venv venv
```

Activate the virtual environment

Windows

```bash
venv\Scripts\activate
```

Install dependencies

```bash
pip install -r requirements.txt
```

---

 Usage

Place the payroll Excel file inside the **input** folder.

Run the application

```bash
python main.py
```

Generated payslips will be saved inside the **output** folder.

---

 Sample Workflow

```
Payroll Excel
        │
        ▼
Read Employee Data
        │
        ▼
Generate Individual Payslips
        │
        ▼
Save Payslips to Output Folder
```

---

 Future Improvements

- PDF Payslip Export
- GUI Application
- Drag & Drop Excel Upload
- Email Payslips Automatically
- Company Logo Support
- Multi-Sheet Payroll Support
- Employee Search
- Progress Bar





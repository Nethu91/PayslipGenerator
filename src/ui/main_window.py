import threading
import subprocess
import sys
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pythoncom

from src.services.excel_reader import ExcelReader
from src.services.payslip_generator import PayslipGenerator
from src.services.pdf_exporter import PdfExporter

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class PayslipApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Swisstek Payslip Generator")
        self.geometry("560x420")
        self.resizable(False, False)

        self.payroll_path = None
        self.template_path = None
        self.output_path = None
        self.employees = []

        self._build_ui()

    # ---------------- UI LAYOUT ----------------
    def _build_ui(self):
        ctk.CTkLabel(
            self, text="SWISSTEK PAYSLIP GENERATOR",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=(20, 15))

        self._file_row("Payroll Excel", self.browse_payroll, "payroll_label")
        self._file_row("Template File", self.browse_template, "template_label")
        self._file_row("Output Folder", self.browse_output, "output_label")

        self.export_pdf_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            self, text="Also export as PDF", variable=self.export_pdf_var
        ).pack(pady=(0, 10))

        self.employee_count_label = ctk.CTkLabel(self, text="Employees : -")
        self.employee_count_label.pack(pady=(15, 5))

        self.progress_bar = ctk.CTkProgressBar(self, width=440)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=(0, 15))

        self.generate_btn = ctk.CTkButton(
            self, text="Generate Payslips", height=40,
            command=self.on_generate, state="disabled"
        )
        self.generate_btn.pack(pady=(0, 10))

        self.status_label = ctk.CTkLabel(self, text="Status : Waiting for files")
        self.status_label.pack(pady=(5, 10))

    def _file_row(self, label_text, command, attr_name):
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(pady=6, padx=20, fill="x")

        ctk.CTkLabel(row, text=label_text, width=110, anchor="w").pack(side="left")
        label = ctk.CTkLabel(row, text="Not selected", anchor="w", text_color="gray")
        label.pack(side="left", fill="x", expand=True)
        setattr(self, attr_name, label)

        ctk.CTkButton(row, text="Browse", width=80, command=command).pack(side="right")

    # ---------------- BROWSE HANDLERS ----------------
    def browse_payroll(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel or CSV files", "*.xlsx *.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        self.payroll_path = Path(path)
        self.payroll_label.configure(text=self.payroll_path.name, text_color="white")
        self._try_load_employees()

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.template_path = Path(path)
        self.template_label.configure(text=self.template_path.name, text_color="white")
        self._check_ready()

    def browse_output(self):
        path = filedialog.askdirectory()
        if not path:
            return
        self.output_path = Path(path)
        self.output_label.configure(text=str(self.output_path), text_color="white")
        self._check_ready()

    # ---------------- LOGIC ----------------
    def _try_load_employees(self):
        try:
            reader = ExcelReader(input_folder=self.payroll_path.parent)
            self.employees = reader.load_employees(self.payroll_path.name)
            self.pay_period = reader.get_pay_period(self.payroll_path.name)
            self.employee_count_label.configure(text=f"Employees : {len(self.employees)}")
            self.status_label.configure(text=f"Status : Loaded ({self.pay_period})")
        except Exception as e:
            self.employees = []
            messagebox.showerror("Error reading payroll file", str(e))
            self.status_label.configure(text="Status : Error reading payroll file")
        self._check_ready()

    def _check_ready(self):
        ready = bool(self.payroll_path and self.template_path and self.output_path and self.employees)
        self.generate_btn.configure(state="normal" if ready else "disabled")

    def on_generate(self):
        self.generate_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Status : Generating...")
        threading.Thread(target=self._generate_worker, daemon=True).start()

    def _generate_worker(self):
        pythoncom.CoInitialize()
        pdf_exporter = PdfExporter() if self.export_pdf_var.get() else None
        try:
            generator = PayslipGenerator(
                template_path=self.template_path,
                output_folder=self.output_path,
            )
            total = len(self.employees)
            for i, emp in enumerate(self.employees, start=1):
                generator.generate(emp, self.pay_period)

                if pdf_exporter:
                    xlsx_file = self.output_path / f"{emp.epf}_{emp.name}.xlsx"
                    pdf_exporter.export(xlsx_file)

                self.progress_bar.set(i / total)
                self.status_label.configure(text=f"Status : Generating {i}/{total} - {emp.name}")

            self.status_label.configure(text=f"Status : Done - {total} payslips generated")
            messagebox.showinfo("Success", f"{total} payslips generated successfully.")
            self._open_output_folder()
        except Exception as e:
            messagebox.showerror("Generation failed", str(e))
            self.status_label.configure(text="Status : Error")
        finally:
            if pdf_exporter:
                pdf_exporter.close()
            pythoncom.CoUninitialize()
            self.generate_btn.configure(state="normal")

    def _open_output_folder(self):
        if sys.platform == "win32":
            subprocess.Popen(f'explorer "{self.output_path}"')


def run():
    app = PayslipApp()
    app.mainloop()
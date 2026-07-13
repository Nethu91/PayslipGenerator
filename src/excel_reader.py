from pathlib import Path
import pandas as pd


class ExcelReader:
    def __init__(self):
        self.input_folder = Path("input")

    def load_excel(self, filename):
        file_path = self.input_folder / filename

        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        df = pd.read_excel(file_path, header=1)

        return df
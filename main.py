import os
import tkinter as tk
from tkinter import filedialog
import csv
from tkinter import ttk
from datetime import datetime
import pandas as pd

class ExcelConverter:
    def __init__(self):
        """
        Inicializa la aplicación y muestra la ventana de selección de directorio.
        """
        self.main_window = tk.Tk()
        self.main_window.geometry("300x100")
        self.main_window.title("Select Directory")
        self.select_button = tk.Button(self.main_window, text="Select Directory", command=self.select_directory)
        self.select_button.pack(pady=20)
        self.main_window.mainloop()

    @staticmethod
    def convert_sheet_to_csv(excel_file, sheet_name):
        """
        Convierte una hoja de cálculo Excel a un archivo CSV.

        Args:
            excel_file (str): Ruta al archivo Excel.
            sheet_name (str): Nombre de la hoja de cálculo a convertir.

        Returns:
            str: Mensaje indicando el resultado de la conversión.
        """
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            csv_file = f"{sheet_name}.csv"
            with open(csv_file, 'w', newline='', encoding='iso-8859-1', errors='replace') as file:
                writer = csv.writer(file)
                valid_columns = [col for col in df.columns if not col.startswith('Unnamed')]
                writer.writerow(valid_columns)
                for _, row in df.iterrows():
                    formatted_row = [value.strftime('%d/%m/%Y') if isinstance(value, datetime) else value for value in row]
                    writer.writerow(formatted_row)
            return f'Sheet "{sheet_name}" has been converted to {csv_file}'
        except pd.errors.EmptyDataError:
            return f'Sheet "{sheet_name}" is empty.'
        except Exception as e:
            return f'Error converting sheet "{sheet_name}": {str(e)}'

    @staticmethod
    def show_result(results):
        """
        Muestra los resultados de la conversión en una ventana emergente.

        Args:
            results (str): Mensaje de resultados a mostrar.
        """
        result_window = tk.Tk()
        result_window.title("Conversion Result")
        frame = ttk.Frame(result_window)
        frame.pack(fill=tk.BOTH, expand=True)
        text_box = tk.Text(frame, wrap=tk.WORD)
        text_box.insert("1.0", results)
        text_box.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_box.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_box.config(yscrollcommand=scrollbar.set)
        close_button = ttk.Button(result_window, text="Close", command=result_window.quit)
        close_button.pack()
        result_window.mainloop()

    @classmethod
    def select_directory(cls):
        """
        Muestra el cuadro de diálogo para seleccionar un directorio y realiza la conversión de archivos en el directorio.

        """
        directory = filedialog.askdirectory()
        if directory:
            results = []
            for file_name in os.listdir(directory):
                full_path = os.path.join(directory, file_name)
                file_name, file_extension = os.path.splitext(file_name)
                if file_extension.lower() in ['.xlsx', '.xls']:
                    xls = pd.ExcelFile(full_path)
                    sheets = xls.sheet_names
                    for sheet in sheets:
                        result = cls.convert_sheet_to_csv(full_path, sheet)
                        results.append(result)
            if results:
                cls.show_result("\n".join(results))
            else:
                cls.show_result("No valid files found in the directory.")
        else:
            cls.show_result("No directory selected.")

if __name__ == "__main__":
    excel_converter = ExcelConverter()

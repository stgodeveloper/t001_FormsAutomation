import logging
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Configurable parameters
LOGS_DIRECTORY = "_logs"  # Directory to store logs
LOGS_FILE = os.path.join(LOGS_DIRECTORY, "input_validation.log")  # Log file path
INPUT_DIRECTORY = "input"
INPUT_FILE = "input_file.xlsx"
INPUT_TEMPLATE_FILE = "input_file_template.xlsx"

# Variables to track validation results
input_file_exists = False
input_file_format = False
input_file_first_row = False

def configure_logging():
    """
    Configures logging for the application.
    """
    if not os.path.exists(LOGS_DIRECTORY):
        os.makedirs(LOGS_DIRECTORY)
    
    # Clear any existing logging configuration
    logging.basicConfig(
        filename=LOGS_FILE,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        force=True  # Force reconfiguration of logging
    )
    logging.info("Logging configured successfully.")

def show_popup_centered(title, message, popup_type="info"):
    """
    Displays a centered popup without extra icons or overlapping windows.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    if popup_type == "info":
        messagebox.showinfo(title, message, parent=root)
    elif popup_type == "error":
        messagebox.showerror(title, message, parent=root)
    elif popup_type == "question":
        return messagebox.askokcancel(title, message, parent=root)
    return None

def validate_input_file():
    """
    Validates the input file based on the following rules:
    1. The file must exist.
    2. The file must follow the same format as the template.
    3. The file must have at least one row of data.
    Returns True if all validations pass, otherwise False.
    """
    global input_file_exists, input_file_format, input_file_first_row

    try:
        input_file_path = os.path.join(INPUT_DIRECTORY, INPUT_FILE)
        if not os.path.exists(input_file_path):
            logging.error(f"The input file '{INPUT_FILE}' does not exist in the '{INPUT_DIRECTORY}' directory.")
            input_file_exists = False
            return False
        input_file_exists = True

        template_path = os.path.join(INPUT_DIRECTORY, INPUT_TEMPLATE_FILE)
        template_columns = pd.read_excel(template_path).columns.tolist()
        input_columns = pd.read_excel(input_file_path).columns.tolist()

        input_file_format = template_columns == input_columns
        if not input_file_format:
            logging.error("The input file does not follow the correct format.")

        input_data = pd.read_excel(input_file_path)
        input_file_first_row = len(input_data) > 0
        if not input_file_first_row:
            logging.error("The input file has no rows of data.")
    
        return input_file_exists and input_file_format and input_file_first_row
    except Exception as e:
        logging.exception("Unexpected error in validation")
        show_popup_centered("Error", "Ha ocurrido un error inesperado", "error")
        return False

def show_validation_failed_popup():
    """
    Shows a popup with the reasons why the validation failed.
    Returns False after the user clicks "Ok".
    """
    reasons = []
    if not input_file_exists:
        reasons.append("- El archivo 'input_file.xlsx' no existe.")
    if not input_file_format:
        reasons.append("- El archivo 'input_file.xlsx' no sigue el formato correcto.")
    if not input_file_first_row:
        reasons.append("- El archivo 'input_file.xlsx' no tiene filas de datos.")

    message = "Se ha cancelado la operación por los siguientes motivos:\n\n" + "\n".join(reasons)
    show_popup_centered("Validación Fallida", message, "info")
    return False

def input_validation_module():
    """
    Main function for the input validation module.
    Returns True if all validations pass and the user confirms, otherwise False.
    """
    response = show_popup_centered("Confirmación", "Se iniciará la carga de datos a Forms. ¿Desea continuar?", "question")
    if not response:
        show_popup_centered("Operación Cancelada", "Se ha cancelado la operación.", "info")
        return False

    if not validate_input_file():
        return show_validation_failed_popup()

    input_file_path = os.path.join(INPUT_DIRECTORY, INPUT_FILE)
    input_data = pd.read_excel(input_file_path)
    transaction_count = len(input_data)
    show_popup_centered("Carga de Transacciones", f"Se cargarán un total de {transaction_count} transacciones.", "info")
    logging.info("Input file validation passed. Returning True.")
    return True

if __name__ == "__main__":
    configure_logging()
    result = input_validation_module()
    logging.info(f"Module finished. Final result: {result}")
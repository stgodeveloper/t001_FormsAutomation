import shutil
import pandas as pd
import logging
import os
from openpyxl import load_workbook

logging.info("---- Starting module 'worktray_creation' ----")

# Configurable parameters
INPUT_DIRECTORY = "input"
WORKTRAY_TEMPLATE_FILE = "worktray_template.xlsx"
INPUT_FILE = "input_file.xlsx"
PROCESS_DATA_DIRECTORY = "process_data"
WORKTRAY_OUTPUT_FILE = "worktray.xlsx"

# Required columns in the template
REQUIRED_COLUMNS = [
    "Nombre", "Producto", "Monto", "Fecha de Solicitud", 
    "Datos correctos", "Ingreso exitoso a Forms", "Observaciones"
]

def create_worktray():
    """
    Creates the worktray by copying the template and updating it with input data.
    Sets "Datos correctos" and "Ingreso exitoso a Forms" to "PENDING" by default.
    """
    try:
        # Paths to the template and input files
        template_path = os.path.join(INPUT_DIRECTORY, WORKTRAY_TEMPLATE_FILE)
        input_file_path = os.path.join(INPUT_DIRECTORY, INPUT_FILE)
        
        # Create the "process_data" directory if it doesn't exist
        if not os.path.exists(PROCESS_DATA_DIRECTORY):
            os.makedirs(PROCESS_DATA_DIRECTORY)
            logging.info(f"Created directory: {PROCESS_DATA_DIRECTORY}")
        
        # Copy the template to the output file
        output_path = os.path.join(PROCESS_DATA_DIRECTORY, WORKTRAY_OUTPUT_FILE)
        shutil.copy(template_path, output_path)
        logging.info(f"Template copied to: {output_path}")
        
        # Read the input file
        input_data = pd.read_excel(input_file_path)
        
        # Validate required columns in the input file
        input_columns = ["Nombre", "Producto", "Monto", "Fecha de Solicitud"]
        if not all(column in input_data.columns for column in input_columns):
            logging.error("The input file is missing required columns.")
            return pd.DataFrame()
        
        # Load the copied template
        worktray_wb = load_workbook(output_path)
        worktray_ws = worktray_wb.active
        
        # Append input data to the worktray
        for index, row in input_data.iterrows():
            worktray_ws.append([
                row["Nombre"], row["Producto"], row["Monto"], row["Fecha de Solicitud"], 
                "PENDING", "PENDING", ""
            ])
        
        # Save the updated worktray
        worktray_wb.save(output_path)
        logging.info(f"Worktray successfully updated and saved to: {output_path}")
        
        return input_data
    
    except Exception as e:
        logging.error(f"Error creating the worktray: {str(e)}")
        return pd.DataFrame()

# Execute the function if the script is run directly
if __name__ == "__main__":
    # Configure logging
    if not os.path.exists("_logs"):
        os.makedirs("_logs")

    logging.basicConfig(
        filename="_logs/worktray_creation.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )
    
    # Create the worktray
    worktray = create_worktray()
    
    if worktray.empty:
        logging.error("Worktray creation failed. Check the logs for details.")
    else:
        logging.info("Worktray creation completed successfully.")
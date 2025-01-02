import os
from dotenv import load_dotenv

import pandas as pd

# from openpyxl import load_workbook

load_dotenv()

# ENV file (.env)
FOLDER = "D:\OneDrive - MPS\BMC\Temp"
FILENAME = "OUTPUT_MH_SUM_update_2024.xlsm"
SHEET_NAME = "DataBase1"

folder_path = FOLDER or os.getenv("FOLDER")
xlsx_filename = FILENAME or os.getenv("FILENAME")
sheet_name = SHEET_NAME or os.getenv("SHEET_NAME")

filename = f"{folder_path}\{xlsx_filename}"
# print(f"{folder_path}\{xlsx_filename}")

# wb = load_workbook(filename, read_only=True, keep_vba=False)
# ws = wb[sheet_name]

data = pd.read_excel(
    filename,
    sheet_name=sheet_name,
    usecols="A:B, E:F, I",
    converters={"YEAR": str, "MONTH": str},
).query("YEAR == '2024'" and "`JOB NO.` == 'BMC-538'")

# data1 = data.query("YEAR == '2024'")

print(data)

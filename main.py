import os
import sys
from dotenv import load_dotenv

import pandas as pd

load_dotenv()

# ENV file (.env)

folder_path = os.getenv("FOLDER") #or FOLDER
xlsx_filename = os.getenv("FILENAME") #or FILENAME
sheet_name = os.getenv("SHEET_NAME") #or SHEET_NAME

filename = f"{folder_path}\\{xlsx_filename}"

job_no = sys.argv[1:][0]

data = pd.read_excel(
    filename,
    sheet_name=sheet_name,
    usecols="A:B, E:F, I",
    converters={"YEAR": int, "MONTH": int},
).query("YEAR == 2024").query("`JOB NO.` == @job_no")

# ).query("`JOB NO.` == @job_no")
# data1 = data.query("YEAR == '2024'")

print(data)

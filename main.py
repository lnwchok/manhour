import os
import sys
import datetime
from dotenv import load_dotenv

import pandas as pd

load_dotenv()

# ENV file (.env)

folder_path = os.getenv("FOLDER") #or FOLDER
xlsx_filename = os.getenv("FILENAME") #or FILENAME
sheet_name = os.getenv("SHEET_NAME") #or SHEET_NAME

filename = f"{folder_path}\\{xlsx_filename}"

job_no = sys.argv[1:][0]

def generate_period():
    x = datetime.datetime.now()
    return f"YEAR == {x.year - 1} & MONTH == {x.month + 5}"

period = generate_period()

pre_data = pd.read_excel(
    filename,
    sheet_name=sheet_name,
    usecols="A:B, E:F, I",
    converters={"YEAR": int, "MONTH": int},
).query("`JOB NO.` == @job_no").query(period)

data = pre_data.pivot_table(values="HOURS", index=["STAFF NAME"], columns=["MONTH"], aggfunc="sum")


print(data)

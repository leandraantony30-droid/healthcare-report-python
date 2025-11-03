

import pandas as pd
import numpy as np

df = pd.read_excel(r"C:\Users\admin\Downloads\healthcare_dataset.xlsx")
blood = df.groupby(["Gender", "Medical Condition"]).size().reset_index(name="Count")
blood_type_count = df.groupby("Blood Type").size().reset_index(name="Count")

# Save both DataFrames to the same sheet
with pd.ExcelWriter(r"C:\Users\admin\Downloads\healthcar.xlsx", engine='xlsxwriter') as writer:
    blood.to_excel(writer, sheet_name='Combined', index=False, startrow=0)
    blood_type_count.to_excel(writer, sheet_name='Combined', index=False, startrow=len(blood) + 3)


import pandas as pd

xl = pd.ExcelFile("backend/uploads/5c873ac1-cdea-48a7-ba72-b364711c4fda_PROCESSED_R122HDCD.xlsx")
df = xl.parse("MODIFY")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
print(df.head(20).to_string())

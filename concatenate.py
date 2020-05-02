import pandas as pd
import time

excel_names = ["SampleOutput3.xlsx", "SampleOutput1.xlsx", "SampleOutput2.xlsx"]

excels = [pd.ExcelFile(name) for name in excel_names]

frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

combined = pd.concat(frames)

combined.to_excel("SampleOutput.xlsx", header=False, index=False)
print('3 excel files concatenated')
time.sleep(2)

import pandas as pd
import os

# file_path=__path__(r"C:\Users\xin\Desktop\py-stu\*.xlsx")
# df=pd.read_excel(file_path)
dfs = []

# read all excel's sheet append to dfs
for fname in os.listdir("./"):
    if fname.endswith(".xlsx") and fname != "final.xlsx":
            df = pd.read_excel(
                  fname,
                  header=None,
                  sheet_name=None
            )
            dfs.extend(df.values())
# contact 
result = pd.concat(dfs)

# output excel
result.to.to_excel("./final.xlsx",index=False)
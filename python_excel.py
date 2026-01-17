import pandas as pd

print(pd.__version__)

df = pd.read_excel(
    r"E:\mywork\education\授课\2518YPython\20260110_Python项目\源码\python_excel\data\demo1.xlsx"
)
# print(df)
# print(df.keys())
df["_source_file"] = "source_file"  # 新增一列
print(df)

df.to_excel("new_demo.xlsx", index=False)  # 保存文件

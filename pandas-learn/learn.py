import pandas as pd

# pd.set_option('display.width', 2000)  # 设置字符显示宽度
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)  # 设置显示最大行
df = pd.read_excel('C1600C390810.xls', 'Sheet1', header=None, skiprows=[0, 1], usecols=[0, 3, 4],
                   names=['部件名称', '重量', '数量'])
print(df)

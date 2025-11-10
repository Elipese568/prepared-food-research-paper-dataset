import pandas as pd
import matplotlib.pyplot as plt
# 设置字体（Windows下使用SimHei，macOS下使用Arial Unicode MS等）
plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置为黑体，适用于Windows
plt.rcParams['axes.unicode_minus'] = False  # 防止负号显示为乱码
# 读取Excel文件
file_path = "商品数据.xlsx"  # 请替换为你的文件路径
file = open(file_path, 'br')
df = pd.read_excel(file, engine="openpyxl")

# 假设“配料”这一列是“Ingredient”，你可以修改成实际列名
ingredient_column = "配料"  # 根据实际列名调整

# 合并所有配料表内容
all_ingredients = df[ingredient_column].dropna().str.replace('(', '、').str.replace(')', '、').str.replace('[', '、').str.replace(']', '、').str.replace('。','、').str.replace(':', '、').str.replace('：', '、').str.replace('、、、', '、').str.replace('、、', '、').str.split('、', expand=True).stack()

# 去除多余的空格，并计算各个材料的出现频率
ingredient_counts = all_ingredients.str.strip().value_counts()

# 打印配料种类及其占比
print(ingredient_counts)
print(all_ingredients.to_list())

# 绘制饼状图
plt.figure(figsize=(8, 6))
plt.pie(ingredient_counts, labels=ingredient_counts.index, autopct='%1.1f%%', startangle=140)
plt.title("配料材料种类及占比")
plt.axis('equal')  # 保证饼状图是圆形
plt.savefig("配料材料占比.png", dpi=500)
plt.show()



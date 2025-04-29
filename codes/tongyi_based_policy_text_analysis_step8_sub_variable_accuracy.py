import pandas as pd

# 定义文件路径
file1 = r"D:/PMC-analysis-using-LLM-main/results/sub_variable_scores_labor.xlsx"
file2 = r"D:/PMC-analysis-using-LLM-main/results/sub_variable_scores_standard.xlsx"
output_file = r"D:/PMC-analysis-using-LLM-main/results/sub_variable_scores_accuracy.xlsx"


# 读取Excel文件并处理合并单元格（填充主变量列）
def read_and_process(file):
    df = pd.read_excel(file)
    # 填充主变量列的缺失值（处理合并单元格）
    df['主变量'] = df['主变量'].ffill()
    return df


try:
    df1 = read_and_process(file1)
    df2 = read_and_process(file2)
except Exception as e:
    print(f"读取文件时发生错误: {e}")
    exit()

# 检查数据列是否一致（排除前两列，只检查数据列）
data_columns1 = df1.columns[2:]  # 从第3列开始的所有列（C列及之后）
data_columns2 = df2.columns[2:]
if not data_columns1.equals(data_columns2):
    print("错误：数据列（C列及之后）的列名或顺序不一致")
    exit()

# 提取需要对比的数据部分（排除前两列：主变量、子变量）
df1_data = df1.iloc[:, 2:]  # 选择第3列到最后一列
df2_data = df2.iloc[:, 2:]

# 检查数据行数是否一致
if len(df1_data) != len(df2_data):
    print("错误：两个表格的数据行数不一致")
    exit()

# 计算相同数据的个数和总数据个数
correct_count = (df1_data == df2_data).sum().sum()
total_count = df1_data.size  # 仅计算数据列的单元格总数
accuracy = correct_count / total_count

# 创建结果数据框
result_df = pd.DataFrame({
    '准确率': [accuracy],
    '相同数据个数': [correct_count],
    '数据总数': [total_count]
})

# 保存结果到新Excel
try:
    with pd.ExcelWriter(output_file) as writer:
        result_df.to_excel(writer, sheet_name='准确率统计', index=False)
    print(f"准确率计算完成，结果已保存到 {output_file}")
    print(f"准确率：{accuracy:.4f}（{correct_count}/{total_count}）")
except Exception as e:
    print(f"保存文件时发生错误: {e}")

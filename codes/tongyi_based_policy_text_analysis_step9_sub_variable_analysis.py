import pandas as pd

# 定义文件路径 Define file path
file_path_labor = r"D:\PMC-analysis-using-LLM-main\results\sub_variable_scores_labor.xlsx"
file_path_standard = r"D:\PMC-analysis-using-LLM-main\results\sub_variable_scores_standard.xlsx"
output_path = r"D:\PMC-analysis-using-LLM-main\results\sub_variable_scores_analysis.xlsx"

# 读取两个表格 Read two tables
df_labor = pd.read_excel(file_path_labor, header=0)
df_standard = pd.read_excel(file_path_standard, header=0)

# 确定数据区域：行从第1行开始（索引0），列从第3列开始（索引2）到最后一列 Determine the data area: The rows start from the 1st row (index 0), and the columns start from the 3rd column (index 2) to the last column.
data_labor = df_labor.iloc[0:, 2:]  # 数据部分（数值部分） Data part (numerical part)
data_standard = df_standard.iloc[0:, 2:]

# 获取主变量和子变量作为行标识 Get the main variable and sub - variable as row identifiers
row_labels = df_labor.iloc[0:, :2].copy()  # 主变量和子变量列 Main variable and sub-variable columns

# ------------------------ 按列分析（单个政策准确率） ---------Analysis by column (accuracy of a single policy)--------------
# 计算每列的准确率：相同值的数量 / 行数 Calculate the accuracy of each column: the number of identical values / the number of rows
column_accuracy = (data_labor == data_standard).mean(axis=0)  # 按列计算均值（布尔值转换为0/1） Calculate the mean by column (convert boolean values to 0/1)
column_accuracy = column_accuracy.round(4)  # 保留四位小数用于后续百分比处理 Retain four decimal places for subsequent percentage processing

# ------------------------ 按行分析（每个子指标准确率） ---Analyze line by line (accuracy of each sub - criterion)------------
# 计算每行的准确率：相同值的数量 / 列数 Calculate the accuracy of each row: the number of identical values / the number of columns
row_accuracy = (data_labor == data_standard).mean(axis=1)  # 按行计算均值 Calculate the mean value row by row
row_accuracy = row_accuracy.round(4)  # 保留四位小数用于后续百分比处理 Retain four decimal places for subsequent percentage processing

# 合并主变量、子变量和行准确率 Merge main variables, sub - variables and row accuracy
row_result = pd.concat([row_labels, row_accuracy], axis=1)
row_result.columns = ['主变量', '子变量', '准确率']

# 将准确率转换为百分比格式（保留两位小数） Convert the accuracy rate to percentage format (retaining two decimal places)
column_accuracy = column_accuracy.apply(lambda x: f"{x * 100:.2f}%")
row_result['准确率'] = row_result['准确率'].apply(lambda x: f"{x * 100:.2f}%")

# 创建Excel写入器 Create an Excel writer
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # 写入列准确率（按列分析结果） Write column accuracy (analysis results by column)
    column_accuracy.to_frame(name='准确率').to_excel(writer, sheet_name='列准确率', index=True)

    # 写入行准确率（按行分析结果） Write line accuracy (analyze results line by line)
    row_result.to_excel(writer, sheet_name='行准确率', index=False)

print("准确率计算完成，结果已保存到指定路径！")

import os
import xlrd
import xlwt
import json
import re
from langchain_community.llms import Tongyi
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain


wb = xlrd.open_workbook("D:/PMC-analysis-using-LLM-main/results/extraction_results_by_Tongyi.xls")
sheet = wb.sheet_by_index(0)
rows = sheet.nrows
# print(sheet.cell(1, 3).value)
release_agency = []
implementation_agency = []
functions = []
measures = []
policy_coverage = []
for i in range(1, rows):
    release_agency += json.loads(sheet.cell(i,1).value.replace('\'', '\"').replace('，', ','))
    # print(sheet.cell(i,2).value)
    implementation_agency += json.loads(sheet.cell(i,2).value)
    # print(sheet.cell(i,3).value)
    functions += json.loads(sheet.cell(i,3).value.replace('\'', '\"').replace('，', ','))
    measures += json.loads(sheet.cell(i,4).value.replace('\'', '\"').replace('，', ','))
    policy_coverage += json.loads(sheet.cell(i, 5).value.replace('\'', '\"').replace('，', ','))

# print(len(str(implementation_agency)))

os.environ["DASHSCOPE_API_KEY"] = "your API key"
template = """Question: {question}

Answer: 按照要求回答这个问题"""
#Answer: According to the requirements, answer this question.

prompt = PromptTemplate(
    template=template,
    input_variables=["question"])
#
# # print(prompt)
#
llm = Tongyi()
llm.model_name = 'qwen-plus'
llm_chain = LLMChain(prompt=prompt, llm=llm)
#
#
prompt_new = ('接下来给出的列表元素帮我将它们去重并归类，请尽量精炼，类别数量少一点，不超过6个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别6"]，'             
          '请输出完整的结果，不要用省略号。列表如下：')

prompt1 = ('接下来给出的列表帮我将它们去重并归类，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')
# prompt_new = ("Please deduplicate and categorize the list elements provided next. Keep the categories as concise as possible with no more than 6 categories, and name each category. Only reply with the categories in list format, i.e., [\"Category 1\", \"Category 2\", ..., \"Category 6\"]. Please output the complete result without ellipses. The list is as follows:")
# prompt1 = ("Please deduplicate and categorize the list provided next. The number of categories should not exceed 8, and name each category. Only reply with the categories in list format, i.e., [\"Category 1\", \"Category 2\", ..., \"Category 8\"]. Please output the complete result without ellipses. The list is as follows:")

prompt1 = prompt1 + str(release_agency)
res1 = llm_chain.invoke(prompt1)
# result1 = re.findall(r'{.+}',res1['text'].replace('\n', '').replace(' ', ''))[0]
print(res1['text'])
# result1_json = json.loads(result1)

prompt2 = ('接下来给出的列表帮我将它们去重并归类，请尽量精炼，类别数量少一点，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')
#prompt2 = ("Please deduplicate and categorize the list provided next. Keep the categories as concise as possible with no more than 8 categories, and name each category. Only reply with the categories in list format, i.e., [\"Category 1\", \"Category 2\", ..., \"Category 8\"]. Please output the complete result without ellipses. The list is as follows:")

prompt2 = prompt2 + str(implementation_agency)
res2 = llm_chain.invoke(prompt2)
print(res2['text'])
prompt2_2 = prompt_new + res2['text']
res2_2 = llm_chain.invoke(prompt2_2)
print(res2_2['text'])

# prompt2_2 = prompt2 + str(res2_1['text'])
# res2_2 = llm_chain.invoke(prompt2_2)
# print(res2_2['text'])
# print(res2['text'].replace('\n', '').replace(' ', ''))
# result2 = re.findall(r'{.+}',res2['text'].replace('\n', '').replace(' ', ''))[0]
# print(result2)
# result2_json = json.loads(result2)


prompt3 = ('接下来给出的列表帮我将它们去重并归类，请尽量精炼，类别数量少一点，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')
# prompt3 = ("Please deduplicate and categorize the list provided next. Keep the categories as concise as possible with no more than 8 categories, and name each category. Only reply with the categories in list format, i.e., [\"Category 1\", \"Category 2\", ..., \"Category 8\"]. Please output the complete result without ellipses. The list is as follows:")

prompt3 = prompt3 + str(functions)
res3 = llm_chain.invoke(prompt3)
print(res3['text'])
prompt3_2 = prompt_new + res3['text']
res3_2 = llm_chain.invoke(prompt3_2)
print(res3_2['text'])
# result3 = re.findall(r'{.+}',res3['text'].replace('\n', '').replace(' ', ''))[0]
# # print(result3)
# result3_json = json.loads(result3)

prompt4 = ('接下来给出的列表帮我将它们去重并归类，请尽量精炼，类别数量少一点，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')
#prompt4 = ("Please deduplicate and categorize the list provided next. Keep the categories as concise as possible with no more than 8 categories, and name each category. Only reply with the categories in list format, i.e., [\"Category 1\", \"Category 2\", ..., \"Category 8\"]. Please output the complete result without ellipses. The list is as follows:")

prompt4 = prompt4 + str(measures)
res4 = llm_chain.invoke(prompt4)
print(res4['text'])
prompt4_2 = prompt_new + res4['text']
res4_2 = llm_chain.invoke(prompt4_2)
print(res4_2['text'])
# result4 = re.findall(r'{.+}',res4['text'].replace('\n', '').replace(' ', ''))[0]
# # print(result4)
# result4_json = json.loads(result4)

prompt5 = ('接下来给出的列表帮我将它们去重并归类，请尽量精炼，类别数量少一点，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')
#prompt5 = ("Please deduplicate and categorize the list provided next. Keep the categories as concise as possible with no more than 8 categories, and name each category. Only reply with the categories in list format, i.e., [\"Category 1\", \"Category 2\", ..., \"Category 8\"]. Please output the complete result without ellipses. The list is as follows:")

prompt5 = prompt5 + str(policy_coverage)
res5 = llm_chain.invoke(prompt5)
print(res5['text'])
prompt5_2 = prompt_new + res5['text']
res5_2 = llm_chain.invoke(prompt5_2)
print(res5_2['text'])

book = xlwt.Workbook(encoding='utf-8')
sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
sheet.write(0, 0, '主属性')#main variables
sheet.write(0, 1, '子属性')#sub variables
sheet.write(1, 0, '发布机构')#Issuing agencies
sheet.write(1, 1, res1['text'])
sheet.write(2, 0, '执行机构')#Implementing agencies
sheet.write(2, 1, res2_2['text'])
sheet.write(3, 0, '功能')#Function
sheet.write(3, 1, res3_2['text'])
sheet.write(4, 0, '措施')#Measures
sheet.write(4, 1, res4_2['text'])
sheet.write(5, 0, '覆盖人群')#Covered population
sheet.write(5, 1, res5_2['text'])
book.save("D:/PMC-analysis-using-LLM-main/results/classification_results_by_Tongyi.xls")

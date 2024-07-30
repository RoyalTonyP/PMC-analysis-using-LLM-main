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
prompt1 = prompt1 + str(release_agency)
res1 = llm_chain.invoke(prompt1)
# result1 = re.findall(r'{.+}',res1['text'].replace('\n', '').replace(' ', ''))[0]
print(res1['text'])
# result1_json = json.loads(result1)

prompt2 = ('接下来给出的列表帮我将它们去重并归类，请尽量精炼，类别数量少一点，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')

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
prompt3 = prompt3 + str(functions)
res3 = llm_chain.invoke(prompt3)
print(res3['text'])
prompt3_2 = prompt_new + res3['text']
res3_2 = llm_chain.invoke(prompt3_2)
print(res3_2['text'])
# result3 = re.findall(r'{.+}',res3['text'].replace('\n', '').replace(' ', ''))[0]
# # print(result3)
# result3_json = json.loads(result3)
#
prompt4 = ('接下来给出的列表帮我将它们去重并归类，请尽量精炼，类别数量少一点，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')
prompt4 = prompt4 + str(measures)
res4 = llm_chain.invoke(prompt4)
print(res4['text'])
prompt4_2 = prompt_new + res4['text']
res4_2 = llm_chain.invoke(prompt4_2)
print(res4_2['text'])
# result4 = re.findall(r'{.+}',res4['text'].replace('\n', '').replace(' ', ''))[0]
# # print(result4)
# result4_json = json.loads(result4)
#
prompt5 = ('接下来给出的列表帮我将它们去重并归类，请尽量精炼，类别数量少一点，类别数量不超过8个，并给每个类起一个名字，'
           '只回复类别，用列表形式回复，即["类别1", "类别2", ...,"类别8"]'
          '请输出完整的结果，不要用省略号。列表如下：')
prompt5 = prompt5 + str(policy_coverage)
res5 = llm_chain.invoke(prompt5)
print(res5['text'])
prompt5_2 = prompt_new + res5['text']
res5_2 = llm_chain.invoke(prompt5_2)
print(res5_2['text'])

book = xlwt.Workbook(encoding='utf-8')
sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
sheet.write(0, 0, '主属性')
sheet.write(0, 1, '子属性')
sheet.write(1, 0, '发布机构')
sheet.write(1, 1, res1['text'])
sheet.write(2, 0, '执行机构')
sheet.write(2, 1, res2_2['text'])
sheet.write(3, 0, '功能')
sheet.write(3, 1, res3_2['text'])
sheet.write(4, 0, '措施')
sheet.write(4, 1, res4_2['text'])
sheet.write(5, 0, '覆盖人群')
sheet.write(5, 1, res5_2['text'])
book.save("D:/PMC-analysis-using-LLM-main/results/classification_results_by_Tongyi.xls")
# result5 = re.findall(r'{.+}',res5['text'].replace('\n', '').replace(' ', ''))[0]
# # print(result5)
# result5_json = json.loads(result5)
# book = xlwt.Workbook(encoding='utf-8')
# sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
# sheet.write(0, 0, '主属性')
# sheet.write(0, 1, '子属性')
# sheet.write(0, 2, '属性元素')
# row = 1
#
# sheet.write(row, 0, '发布机构')
# for i in result1_json.keys():
#     sheet.write(row, 1, i)
#     sheet.write(row, 2, str(result1_json[i]))
#     row += 1
#
# sheet.write(row, 0, '执行机构')
# for i in result2_json.keys():
#     sheet.write(row, 1, i)
#     sheet.write(row, 2, str(result2_json[i]))
#     row += 1
#
# sheet.write(row, 0, '功能')
# for i in result3_json.keys():
#     sheet.write(row, 1, i)
#     sheet.write(row, 2, str(result3_json[i]))
#     row += 1
#
# sheet.write(row, 0, '措施')
# for i in result4_json.keys():
#     sheet.write(row, 1, i)
#     sheet.write(row, 2, str(result4_json[i]))
#     row += 1
#
# sheet.write(row, 0, '覆盖人群')
# for i in result5_json.keys():
#     sheet.write(row, 1, i)
#     sheet.write(row, 2, str(result5_json[i]))
#     row += 1
#
# book.save("D:/成都理工大学重要文件夹/Text Analysis and Evaluation of China's financial inclusion Policy based on text mining/中间结果/classification_results_by_Tongyi.xls")
# extraction_temp += res['text']


        # # ==========4.提取政策范围===========
        # prompt4 = ('对于给定的政策文本，帮我判断政策的涉及的方面，方面包含：第一，经济；第二，社会；第三，政治；第四：技术。'
        #            '回复的格式是字典格式，即{"经济":1或0, "社会":1或0, "政治":1或0, "技术":1或0}，这里的1表示涉及该方面，0表示不涉及，对一个政策而言，有可能涉及多个方面，则可以有多个方面取值为1。'
        #            '只回复该字典，不要回复其他内容，字典请在一行中输出，不要加入换行符。文本如下：')
        # prompt4 += doc.page_content
        # res = llm_chain.invoke(prompt4)
        # print(res['text'])
        # policy_area_temp.append(json.loads(res['text']))
        #
        # # ==========5.提取政策工具类型===========
        # prompt5 = ('对于给定的政策文本，帮我判断政策的工具类型，包含以下类型：'
        #            '第一类，供给型政策工具，即指政府在人才培养、资金支持、技术支持、公共服务等方面直接投入资源，推动特定领域或行业的发展；'
        #            '第二类，需求型政策工具，指政府通过政府采购、贸易政策、用户补贴、应用示范和价格指导等方式，减少市场的不确定性，培育并扩大特定市场，从需求侧拉动产业的发展；'
        #            '第三类，环境型政策工具，指政府通过目标规划、金融支持、法规规范、标准管理、税收优惠等方式，为特定领域或行业的发展提供有利的政策环境、金融环境和法律环境，间接促进其发展。'
        #            '回复的格式是字典格式，即{"供给":1或0, "需求":1或0, "环境":1或0}，这里的1表示属于该类型，0表示不属于该类型，对一个政策的工具类型而言，有可能包含涉及多个方面，则可以有多个方面取值为1。'
        #            '只回复该字典，不要回复其他内容，字典请在一行中输出，不要加入换行符。文本如下：')
        # prompt5 += doc.page_content
        # res = llm_chain.invoke(prompt5)
        # print(res['text'])
        # policy_tool_temp.append(json.loads(res['text']))

        # ==========6.提取政策执行机构===========
        # prompt6 = ('对于给定的政策文本片段，帮我提取政策的执行机构（政策执行机构指的是这个政策片段指定的需要执行这个政策片段的部门，注意如果有重复的部门需要去重，如果机构名称省略了具体的城市或省份，请根据上下文给出完整的部门名称，如‘省人民政府’补充为‘XX省人民政府’），用列表[机构1,机构2,...]的格式存放；'
        #           '只回复列表，不要其他的文字。文本如下：')
        # prompt6 += doc.page_content
        # res = llm_chain.invoke(prompt6)
        # print(res['text'])
        # implementation_agency_temp += res['text']

        # ==========7.进行政策文本摘要===========

    # print('===========================================')
    # # ==========4.提取政策范围汇总===========
    # result4 = policy_area_temp[0]
    # for key in policy_area_temp[0].keys():
    #     result4[key] = 1 if sum([policy_area_temp[index][key] for index in range(len(policy_area_temp))]) > 0 else 0
    # print(result4)
    # policy_area.append(result4)
    # # ==========5.提取政策工具类型汇总===========
    # result5 = policy_tool_temp[0]
    # for key in policy_tool_temp[0].keys():
    #     result5[key] = 1 if sum([policy_tool_temp[index][key] for index in range(len(policy_tool_temp))]) > 0 else 0
    # print(result5)
    # policy_tool.append(result5)
    # print('===========================================')
    # ==========6.提取政策执行机构汇总===========
    # prompt6_2 = ('这些列表当中存放了政策的执行部门名称，帮我将给出的多个列表合并成一个列表，去掉其中的重复的元素，去掉不属于部门名称的元素，结果只返回列表，不要其他文字。列表如下：')
    # prompt6_2 += implementation_agency_temp
    # res = llm_chain.invoke(prompt6_2)
    # print(res['text'])
    # implementation_agency.append(res['text'])

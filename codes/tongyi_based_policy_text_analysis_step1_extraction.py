import os
os.environ["DASHSCOPE_API_KEY"] = "your API key"

from langchain_community.llms import Tongyi
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from langchain_community.document_loaders import UnstructuredFileLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
import json
import xlwt


template = """Question: {question}

Answer: 按照要求回答这个问题"""
#Answer: Answer the question according to the requirements.

prompt = PromptTemplate(
    template=template,
    input_variables=["question"])

# print(prompt)

llm = Tongyi()



llm_chain = LLMChain(prompt=prompt, llm=llm)

# question = "我给通义千问传入的prompt最大长度是可以是多少？"
# question = "What is the maximum length of the prompt that I can input to Tongyi Qianwen?"
# res = llm_chain.invoke(question)
#
# print(res)
# 当前目录 Current directory
base_dir = "D:/PMC-analysis-using-LLM-main/datasets/"
# 获取当前目录下的所有文件 
#Get all files in the current directory
files = [os.path.join(base_dir, file) for file in os.listdir(base_dir)]
release_agency = []
implementation_agency = []
function_and_measures = []
policy_coverage = []
book = xlwt.Workbook(encoding='utf-8')
sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
sheet.write(0, 0, '政策文件名')#Policy document name
sheet.write(0, 1, '政策发布机构')#Policy issuing authority
sheet.write(0, 2, '政策执行机构')#Policy implementation agency
sheet.write(0, 3, '政策功能')#Policy function
sheet.write(0, 4, '政策措施')#Policy measures
sheet.write(0, 5, '政策覆盖对象')#Policy coverage targets
row = 1
# 遍历文件列表，输出文件名 
#Traverse the file list and output the file names
for file in files:
    sheet.write(row, 0, file)
    loader = UnstructuredFileLoader(file)
    documents = loader.load()
    start = documents[0].page_content[:500]
    # ==========1.获取发布机构===========
    prompt1 = ('对于给定的政策文本的开头，帮我提炼出政策的发布机构，'
              '回复的格式是列表格式，列表中的元素表示发布机构，只回复列表，即["机构1", "机构2", ...]格式，若没有给出发布机构，请回复空列表，即[]，不要回复其他内容，不要在回复中添加额外的换行符。文本如下：')
    #For the beginning of the given policy text, help me extract the issuing agency of the policy.The reply format is in a list format. The elements in the list represent the issuing institutions. Only reply with the list, that is, in the format of ["Institution 1", "Institution 2", ...]. If the issuing institution is not provided, please reply with an empty list, that is, []. Do not reply with any other content, and do not add extra line breaks in the reply. The text is as follows: 
    prompt1 += start
    res = llm_chain.invoke(prompt1)
    print(res['text'])
    sheet.write(row, 1, res['text'])
    # release_agency.append(json.loads(res['text']))




    # ==========4-6.提取政策范围、工具类型和执行机构===========
    text_spliter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=10)
    split_docs = text_spliter.split_documents(documents)
    policy_area_temp = []
    policy_tool_temp = []
    implementation_agency_temp = ''
    extraction_temp = ''
    for doc in split_docs:
        # ==========6.提取政策执行机构===========
        prompt6 = ('对于给定的政策文本片段，帮我提取政策的执行机构（政策执行机构指的是这个政策片段指定的需要执行这个政策片段的部门，注意如果有重复的部门需要去重，如果机构名称省略了具体的城市或省份，请根据上下文给出完整的部门名称，如‘省人民政府’补充为‘XX省人民政府’），'
                   '用列表["机构1", "机构2", ...]的格式存放，若没有给出具体的执行机构，请返回空列表，即[]；'
                  '只回复列表，不要其他的文字，不要在回复中添加额外的换行符。文本如下：')
        #For a given policy text snippet, help me extract the policy implementing institutions (the policy implementing institutions refer to the departments specified in this policy snippet that are required to implement this policy snippet. Note that if there are duplicate departments, duplicates need to be removed. If the specific city or province is omitted in the institution name, please provide the complete department name according to the context, for example, "Provincial People's Government" should be supplemented as "People's Government of XX Province"). Store them in the format of a list ["Institution 1", "Institution 2", ...]. If no specific implementing institutions are given, please return an empty list, that is, []; Only reply with the list, without any other words, and do not add extra line breaks in the reply. The text is as follows: 
        prompt6 += doc.page_content
        res = llm_chain.invoke(prompt6)
        print(res['text'])
        implementation_agency_temp += res['text']

        # ==========7.进行政策文本摘要===========
        prompt7 = ('对于给定的政策文本片段，帮我进行摘要，200字以内；'
                  '只回复摘要结果，不要其他的文字。文本如下：')
        #For a given policy text snippet, help me summarize it within 200 words. Only reply with the summary result and no other words. The text is as follows: 
        prompt7 += doc.page_content
        res = llm_chain.invoke(prompt7)
        # print(res['text'])
        extraction_temp += res['text']

    # ==========6.提取政策执行机构汇总===========
    prompt6_2 = ('这些列表当中存放了政策的执行部门名称，帮我将给出的多个列表合并成一个列表，去掉其中的重复的元素，去掉不属于部门名称的元素，结果只返回列表，即["机构1", "机构2", ...]格式，若没有给出具体的执行机构，请返回空列表，即[]；'
                 '不要其他文字，不要在回复中添加额外的换行符。列表如下：')
    #These lists contain the names of the policy implementation departments. Help me merge the given multiple lists into one list, remove the duplicate elements in it, and remove the elements that do not belong to the department names. The result should only return a list, that is, in the format of ["Institution 1", "Institution 2", ...]. If no specific implementing institutions are given, please return an empty list, that is, []; Do not include other words, and do not add extra line breaks in the reply. The lists are as follows: 
    prompt6_2 += implementation_agency_temp
    res = llm_chain.invoke(prompt6_2)
    print(res['text'])
    sheet.write(row, 2, res['text'])
    # implementation_agency.append(json.loads(res['text']))
    # ==========7.根据政策文本摘要提取政策功能和措施===========
    prompt7_2 = ('下面给出的是政策文本摘要，帮我从摘要当中进一步提炼政策功能和政策措施，功能和措施都用短语来进行概括，短语尽量精炼，单个短语长度不要太长，短语的总数量不要太多。'
               '提炼格式为字典格式，即：{"功能":["功能1", "功能2", "功能3", ...], "措施":["措施1", "措施2", "措施3", ...]}；'
                 '若政策文本中不涉及功能或者措施，列表中的元素可以为空'
               '只回复该字典，不要回复其他内容，字典请在一行中输出，不要加入换行符。文本如下：')
    #The following is the summary of the policy text. Help me further extract the policy functions and policy measures from the summary. Both the functions and measures should be summarized in phrases. The phrases should be as concise as possible, not too long individually, and the total number of phrases should not be excessive. The extraction format is in dictionary form, that is: {"functions":["Function 1", "Function 2", "Function 3", ...], "measures":["Measure 1", "Measure 2", "Measure 3", ...]}; If the policy text does not involve functions or measures, the elements in the list can be empty. Only reply with this dictionary, and do not reply with any other content. Please output the dictionary on a single line without adding line breaks. The text is as follows: 
    prompt7_2 += extraction_temp
    res = llm_chain.invoke(prompt7_2)
    print(res['text'])
    function_and_measures_dict = json.loads(res['text'])
    sheet.write(row, 3, str(function_and_measures_dict['功能']))
    sheet.write(row, 4, str(function_and_measures_dict['措施']))
    # function_and_measures.append(json.loads(res['text']))


    # ==========9.根据政策文本摘要提炼政策主要为哪些对象解决金融服务困难===========
    prompt9 = ('下面给出的是关于普惠金融的政策文本摘要，帮我提取出该政策主要是为哪些对象解决金融服务困难的问题。'
               '回复的格式是列表格式，列表中的每个元素为提取出的对象，即["对象1", "对象2", ...]，如果没有涉及到任何的对象，请回复空列表，即[]。'
               '只回复该列表，不要回复其他内容。文本如下：')
    #The following is the summary of the policy text on inclusive finance. Help me extract the objects for which this policy mainly addresses the difficulties in obtaining financial services. The reply format is in a list format, and each element in the list is the extracted object, that is, ["Object 1", "Object 2", ...]. If no objects are involved, please reply with an empty list, that is, []. Only reply with this list and do not reply with any other content. The text is as follows: 
    prompt9 += extraction_temp
    res = llm_chain.invoke(prompt9)
    print(res['text'])
    sheet.write(row, 5, res['text'])
    # policy_coverage.append(json.loads(res['text']))
    row += 1
book.save("D:/PMC-analysis-using-LLM-main/results/extraction_results_by_Tongyi.xls")




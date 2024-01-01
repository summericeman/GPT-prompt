
import os
import win32com.client as win32
import re

def read_doc_file(file_name):
    # 确保文件存在
    file_path = os.path.abspath(file_name)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"无法找到文件：{file_path}")

    # 初始化 Word COM 对象
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    # 打开文档
    doc = word.Documents.Open(file_path)
    doc.Activate()

    # 提取内容并检查是否邻接
    content = []
    start_collecting = False
    prev_line = ""
    for para in doc.Paragraphs:
        line = para.Range.Text.strip()
        if "供方：" in prev_line and "需方：" in line:
            start_collecting = True
            if "签字盖章：" not in prev_line:  # 检查是否包含"签字盖章："
                content.append(prev_line)  # 添加“供方”的行
        if start_collecting:
            if "签字盖章：" not in line:  # 检查是否包含"签字盖章："
                content.append(line)
        prev_line = line

    # 关闭文档
    doc.Close(False)

    return content

def splitOutputIntoArray(output_info):
    # Split the output into lines
    lines = output_info.split('\n')

    # Initialize the info array
    info_array = []
    temp_array = []

    # Function to check if a line is empty (no Chinese characters)
    def is_empty_line(line):
        return not re.search("[\u4e00-\u9fff]", line)

    # Function to clean up special characters
    def clean_line(line):
        # Replace non-breaking spaces and other unwanted characters
        line = line.replace('\xa0', ' ').replace('\r', '').replace('\x07', '')
        return line

    for line in lines:
        line = clean_line(line)
        if is_empty_line(line):
            # Empty line found, add the current temp array to info_array if it's not empty
            if temp_array:
                info_array.append(temp_array)
                temp_array = []
        else:
            # Add line to the current temp array
            temp_array.append(line)

    # Add the last temp array if not empty
    if temp_array:
        info_array.append(temp_array)

    return info_array

def getResultStr(info_array):
    result_str = ""
    for sub_array in info_array:
        if len(sub_array) > 1:
            result_str += sub_array[1] + '\n'
    return result_str

def getCustomerInfoResult(docName):
    doc_contents = read_doc_file(docName)
    output_info = '\n'.join(doc_contents)
    info_array = splitOutputIntoArray(output_info)
    result_string = getResultStr(info_array)
    return result_string

# 主函数
if __name__ == "__main__":
    # 设置文件名
    docName = 'SX-3265.doc'

    # 调用 getCustomerInfoResult 函数并打印结果
    result_string = getCustomerInfoResult(docName)
    print(result_string)

#this code is from getCustomerInfo.py

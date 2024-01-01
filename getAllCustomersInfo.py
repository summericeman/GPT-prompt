import os
import getCustomerInfo  # 确保 getCustomerInfo.py 在同一目录下或在Python搜索路径中

def check_existing_doc_files(directory):
    existing_docs = []
    for subdir in os.listdir(directory):
        subdir_path = os.path.join(directory, subdir)
        if os.path.isdir(subdir_path):
            expected_doc_file = os.path.join(subdir_path, f"{subdir}.doc")
            if os.path.isfile(expected_doc_file):
                existing_docs.append(expected_doc_file)  # 现在保存完整的文件路径
    return existing_docs

def processResult(existing_docs):
    with open("result.txt", "w", encoding="utf-8") as result_file:
        for doc_file in existing_docs:
            print(f"正在处理文件: {doc_file}")  # 这会打印出当前正在处理的文件路径
            result_string = getCustomerInfo.getCustomerInfoResult(doc_file)
            result_file.write(result_string + '\n')
    print("处理完成，结果已保存在 result.txt 文件中。")

# 主函数
if __name__ == "__main__":
    contract_directory = "合同"  # 替换为正确的路径
    existing_docs = check_existing_doc_files(contract_directory)
    processResult(existing_docs)

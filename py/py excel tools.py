import os
import docx

def find_string_in_docx(file_path, target_string):
    """
    在指定的 DOCX 文件中查找特定字符串。

    参数:
    file_path (str): DOCX 文件的路径。
    target_string (str): 需要查找的特定字符串。

    返回:
    bool: 如果找到特定字符串返回 True，否则返回 False。
    """
    try:
        # 打开 DOCX 文件
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            # 在段落中查找特定字符串
            if target_string in para.text:
                return True
        return False
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return False

# 以下是测试代码
if __name__ == "__main__":
    file_path = "example.docx"  # 替换为实际的 DOCX 文件路径
    target_string = "ABCDEF"
    result = find_string_in_docx(file_path, target_string)
    if result:
        print(f"在文件中找到了字符串 '{target_string}'。")
    else:
        print(f"在文件中未找到字符串 '{target_string}'。")
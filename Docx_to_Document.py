import json
from docx import Document


def extract_content_from_docx(file_path):
    # 打开并读取Word文档
    doc = Document(file_path)

    # 初始化存储内容的字典和跟踪当前标题的变量
    content_dict = {}
    current_section = None
    current_subsection = None
    current_subsubsection = None

    # 遍历文档中的每个段落
    for paragraph in doc.paragraphs:
        # 检查段落是否为标题
        if paragraph.style.name.startswith('Heading'):
            # 获取标题的级别（例如，Heading 1, Heading 2, Heading 3）
            level = int(paragraph.style.name.split()[-1])

            # 根据标题级别更新当前的章节、子章节或子子章节
            if level == 1:
                # 一级标题（最高级别）
                current_section = paragraph.text  # 存储当前的标题名
                content_dict[current_section] = {}  # 将当前标题存入字典
                current_subsection = None
                current_subsubsection = None
            elif level == 2 and current_section:
                # 二级标题
                current_subsection = paragraph.text
                content_dict[current_section][current_subsection] = {}
                current_subsubsection = None
            elif level == 3 and current_section and current_subsection:
                # 三级标题
                current_subsubsection = paragraph.text
                content_dict[current_section][current_subsection][current_subsubsection] = {}
            continue

        # 将段落文本按换行符分割，并添加到相应的章节或子章节
        if current_section:
            lines = paragraph.text.split('\n')  # 获取正文每行的内容
            for line in lines:  # 正文内容每行依次处理
                if line.strip():  # 跳过空行
                    if '：' in line:
                        para_key, para_value = line.split('：', 1)
                        para_key = para_key.strip()  # 冒号前面的部分作为键
                        para_value = para_value.strip()  # 冒号后面的部分作为值
                    else:
                        para_key = f'段落 {len(content_dict[current_section]) + 1}'  # 使用默认键
                        para_value = line.strip()  # 整个段落作为值

                    if current_subsubsection:
                        # 如果有三级标题，将内容添加到对应的子子章节
                        content_dict[current_section][current_subsection][current_subsubsection][para_key] = para_value
                    elif current_subsection:
                        # 如果有二级标题，将内容添加到对应的子章节
                        content_dict[current_section][current_subsection][para_key] = para_value
                    else:
                        # 如果没有子标题，将内容添加到对应的章节
                        content_dict[current_section][para_key] = para_value

    return content_dict


# 示例用法
file_path = 'D:/test.docx'
content = extract_content_from_docx(file_path)
# print(json.dumps(content, ensure_ascii=False, indent=4))
print(content)

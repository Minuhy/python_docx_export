if __name__ == '__main__':
    pass

import docx

from 测试例程.文件路径 import docx_file

# 打开docx文档
docx_document = docx.Document(docx_file)

all_text = ''
for paragraph in docx_document.paragraphs:
    all_text += paragraph.text
print('所有文本：', all_text)

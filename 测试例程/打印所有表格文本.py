if __name__ == '__main__':
    pass

import docx

from 测试例程.文件路径 import docx_file

# 打开docx文档
docx_document = docx.Document(docx_file)

all_table_text = ''
for table in docx_document.tables:
    for cell in getattr(table, '_cells'):
        all_table_text += cell.text + ' '  # 单元格之间用空格隔开
print('所有表格文本：', all_table_text)

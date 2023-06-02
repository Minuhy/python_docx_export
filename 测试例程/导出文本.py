if __name__ == '__main__':
    pass

from main import get_new_path
from 测试例程.打印所有文本 import all_text
from 测试例程.打印所有表格文本 import all_table_text

text_file_path = get_new_path(0, 'docx文本.txt')
with open(text_file_path, 'w', encoding='utf-8') as f:
    f.write(all_text)
    f.write(all_table_text)

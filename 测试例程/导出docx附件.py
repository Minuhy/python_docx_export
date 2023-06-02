if __name__ == '__main__':
    pass

import os
import docx

from 测试例程.文件路径 import docx_file

# 打开docx文档
docx_document = docx.Document(docx_file)

# 文件导出列表
export_files = []

# 遍历所有附件
index = 0
docx_related_parts = docx_document.part.related_parts
for part in docx_related_parts:
    part = docx_related_parts[part]
    part_name = str(part.partname)  # 附件路径（partname）
    if part_name.startswith('/word/media/') or part_name.startswith('/word/embeddings/'):  # 只导出这两个目录下的
        # 构建导出路径
        index += 1
        save_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前py脚本路径
        index_str = str(index).rjust(2, '0')
        save_path = save_dir + '\\' + index_str + ' - ' + os.path.basename(part.partname)  # 拼接路径
        print('导出路径：', save_path)

        # 写入文件
        with open(save_path, 'wb') as f:
            f.write(part.blob)
        # 记录文件
        export_files.append(save_path)
print('导出的所有文件：', export_files)

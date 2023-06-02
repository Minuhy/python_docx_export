if __name__ == '__main__':
    pass

import os
from oletools import oleobj
from 测试例程.导出docx附件 import export_files

for file in export_files:
    if file.endswith('.bin') and 'ole' in file.lower():  # .bin 作为后缀且文件名中有ole，则被认为是OLE文件
        res = oleobj.main([file])
        if res == 1:  # 1为成功提取
            os.remove(file)  # 删除OLE文件，仅保留原始附件
        else:
            print(file, '提取OLE失败')

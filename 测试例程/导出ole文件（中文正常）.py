if __name__ == '__main__':
    pass

import os
from oletools import oleobj
from 测试例程.导出docx附件 import export_files
from 测试例程.导出docx附件 import index


def re_decode(s, encoding='gbk'):
    """
    重新解码，解决oleobj对中文乱码的问题
    :param s: 原始字符串
    :param encoding: 新的解码编码
    :return: 新的字符串
    """
    i81 = s.encode('iso-8859-1')
    return i81.decode(encoding)


for file_path in export_files:  # 遍历导出的文件

    # 不符合 .bin 作为后缀且文件名中有ole，则不被认为是OLE文件，跳过
    if not (file_path.endswith('.bin') and 'ole' in file_path.lower()):
        continue

    # 准备导出 OLE 文件
    has_error = False
    export_ole_files = []
    for ole in oleobj.find_ole(file_path, None):  # 找OLE文件（oleobj支持对压缩包里的OLE处理）
        if ole is None:  # 没有找到 OLE 文件，跳过
            continue

        for path_parts in ole.listdir():  # 遍历OLE中的文件

            # 判断是不是[1]Ole10Native，使用列表推导式忽略大小写，不是的话就不要继续了
            if '\x01ole10native'.casefold() not in [path_part.casefold() for path_part in path_parts]:
                continue

            stream = None
            try:
                # 使用 Ole File 打开 OLE 文件
                stream = ole.openstream(path_parts)
                opkg = oleobj.OleNativeStream(stream)
            except IOError:
                print('不是OLE文件：', path_parts)
                if stream is not None:  # 关闭文件流
                    stream.close()
                continue

            # 打印信息
            if opkg.is_link:
                print('是链接而不是文件，跳过')
                continue

            ole_filename = re_decode(opkg.filename)
            ole_src_path = re_decode(opkg.src_path)
            ole_temp_path = re_decode(opkg.temp_path)
            print('文件名：{0}，原路径：{1}，缓存路径：{2}'.format(ole_filename, ole_src_path, ole_temp_path))

            # 生成新的文件名（这部分与上一部分导出docx的文件的构建路径类似，可以封装为函数）
            index += 1
            save_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前py脚本路径
            index_str = str(index).rjust(2, '0')
            filename = save_dir + '\\' + index_str + ' - ' + ole_filename  # 拼接路径
            print('导出路径：', filename)

            # 转存
            try:
                print('导出OLE中的文件：', filename)
                with open(filename, 'wb') as writer:
                    n_dumped = 0
                    next_size = min(oleobj.DUMP_CHUNK_SIZE, opkg.actual_size)
                    while next_size:
                        data = stream.read(next_size)
                        writer.write(data)
                        n_dumped += len(data)
                        if len(data) != next_size:
                            print('想要读取 {0}, 实际取得 {1}'.format(next_size, len(data)))
                            break
                        next_size = min(oleobj.DUMP_CHUNK_SIZE, opkg.actual_size - n_dumped)
                export_ole_files.append(filename)
            except Exception as exc:
                has_error = True
                print('在转存时出现错误：{0} {1}'.format(filename, exc))
            finally:
                stream.close()
    if export_ole_files:  # 如果有解出ole包
        export_files.remove(file_path)
        export_files += export_ole_files
        # 删除 bin 文件
        if not has_error:
            os.remove(file_path)
            print('已删除OLE打包文件，仅保留原文件', file_path)
print('最终导出文件：', export_files)

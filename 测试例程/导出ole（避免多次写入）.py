if __name__ == '__main__':
    pass

import os
import docx
from oletools import oleobj
from 测试例程.文件路径 import docx_file

# 打开docx文档
docx_document = docx.Document(docx_file)


def get_new_path(i: int, file: str, out_dir: str = __file__):
    """
    根据序号和文件名拿到新的路径名，
    默认输出目录为当前文件所在文件夹
    :param i: 序号
    :param file: 原始文件名（路径）
    :param out_dir: 输出目录，默认为当前文件所在文件夹
    :return: 新文件名（格式为“序号 - 原文件名”）
    """
    self_dir = os.path.dirname(os.path.abspath(out_dir))  # 获取导出路径
    i = str(i).rjust(2, '0')  # 不够前面添0
    return self_dir + '\\' + i + ' - ' + os.path.basename(file)  # 拼接路径


def re_decode(s: str, encoding: str = 'gbk'):
    """
    重新解码，解决oleobj对中文乱码的问题
    :param s: 原始字符串
    :param encoding: 新的解码编码，默认为 GBK
    :return: 新的字符串
    """
    i81 = s.encode('iso-8859-1')
    return i81.decode(encoding)


# 文件导出列表
export_files = []
# 遍历所有附件
docx_related_parts = docx_document.part.related_parts
for part in docx_related_parts:
    part = docx_related_parts[part]
    part_name = str(part.partname)  # 附件路径（partname）
    if part_name.startswith('/word/media/') or part_name.startswith('/word/embeddings/'):  # 只导出这两个目录下的
        # 构建导出路径
        save_path = get_new_path(len(export_files) + 1, part.partname)

        # ole 文件判断
        # 不符合 .bin 作为后缀且文件名中有ole，则不被认为是OLE文件
        if not (save_path.endswith('.bin') and 'ole' in save_path.lower()):
            # 直接写入文件
            print('DOCX 导出路径：', save_path)
            with open(save_path, 'wb') as f:
                f.write(part.blob)
            export_files.append(save_path)  # 记录文件
        else:
            # 将字节数组传递给oleobj处理
            for ole in oleobj.find_ole(save_path, part.blob):
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
                        print('非OLE文件：', path_parts)
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

                    # 生成新的文件名
                    filename = get_new_path(len(export_files) + 1, ole_filename)
                    print('OLE 导出路径：', filename)

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
                                    print('预计读取 {0}, 实际取得 {1}'.format(next_size, len(data)))
                                    break
                                next_size = min(oleobj.DUMP_CHUNK_SIZE, opkg.actual_size - n_dumped)
                        export_files.append(filename)  # 记录导出的文件
                    except Exception as exc:
                        print('在转存时出现错误：{0} {1}'.format(filename, exc))
                    finally:
                        stream.close()
print('所有导出文件：', export_files)

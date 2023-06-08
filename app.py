import logging
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
from tkinter.messagebox import askyesno, showinfo, showwarning

import export
from gui import ApplicationGUI

v = 'v 0.1.0.1'
logger = None


def log_create():
    # APP根目录
    app_path = os.path.dirname(os.path.abspath(__file__))
    # 日志
    func_logger = logging.getLogger()

    # 日志目录路径（一天一个文件）
    dt = datetime.now()
    log_file_path = os.path.abspath(app_path + "/logs/app" + dt.strftime('%j').rjust(4, '0') + ".log")

    # 处理器
    console_handler = logging.StreamHandler()  # 控制台处理器
    file_handler = logging.FileHandler(log_file_path, mode='a+', encoding="UTF-8")  # 文件处理器

    # 日志格式
    formatter = logging.Formatter(fmt="%(asctime)s |-%(levelname)s in %(name)s@%(funcName)s - %(message)s",
                                  datefmt="%Y-%m-%d %H:%M:%S")
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    # 日志等级
    func_logger.setLevel(logging.DEBUG)

    # 将处理器添加至日志器中
    func_logger.addHandler(console_handler)
    func_logger.addHandler(file_handler)

    print('日志位置：', log_file_path)
    return func_logger


def file_check(file_path: str):
    if os.path.basename(file_path).startswith('~$'):
        logger.debug('未添加：缓存文件 %s', file_path)
        return False
    if not os.path.basename(file_path).endswith('.docx'):
        logger.debug('未添加：非docx文件 %s', file_path)
        return False
    if not os.path.isfile(file_path):
        logger.debug('未添加：不是文件 %s', file_path)
        return False
    return True


class Application(ApplicationGUI):
    def __init__(self, master=None, version=''):
        super().__init__(master)
        self.master = master
        self.version = version

        # 待处理文件列表
        self.file_list = []
        # 文件添加索引
        self.file_index = 0

        self.main_listbox = None
        self.create_list(self.frame_list)
        self.set_combobox_text()
        self.set_checkbutton_text()
        self.register_command()

        self.entry_tips_val.set('(>_<)')

        self.lb_brand.configure(text='批量导出docx文本、图片和附件程序\n' + self.version + '\n@B站：敏Ymm')

    def set_checkbutton_text(self):
        self.che_text.set(1)
        self.che_image.set(1)
        self.che_table.set(1)
        self.che_combine.set(1)
        self.che_attachment.set(1)
        self.che_info.set(1)

    def set_combobox_text(self):
        self.cb_save_position['value'] = ('原文件目录中以原文件命名的子文件夹中', '原文件所在文件夹', '自定义文件夹')
        self.combobox_save_path.set(self.cb_save_position['value'][0])
        self.cb_save_position.current(0)

        self.cb_name_part1['value'] = ('|自增编号|', '|连接符|', '|原文件名|', '|后缀名|')
        self.cb_name_part2['value'] = ('|连接符|', '|原文件名|', '|后缀名|', '|自增编号|')
        self.cb_name_part3['value'] = ('|原文件名|', '|后缀名|', '|自增编号|', '|连接符|')
        self.cb_name_part4['value'] = ('|后缀名|', '|自增编号|', '|连接符|', '|原文件名|')

        self.combobox_name1.set(self.cb_name_part1['value'][0])
        self.combobox_name2.set(self.cb_name_part2['value'][0])
        self.combobox_name3.set(self.cb_name_part3['value'][0])
        self.combobox_name4.set(self.cb_name_part4['value'][0])
        self.cb_name_part1.current(0)
        self.cb_name_part2.current(0)
        self.cb_name_part3.current(0)
        self.cb_name_part4.current(0)

    def create_list(self, frame):
        # 一个列表
        list_frame = tk.Frame(frame)
        scroll_h_bar = tk.Scrollbar(list_frame, orient=tk.HORIZONTAL)  # 水平滚动条组件
        scroll_v_bar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, )  # 垂直滚动条组件
        self.main_listbox = tk.Listbox(list_frame,
                                       width=60, height=20,
                                       selectmode=tk.MULTIPLE,
                                       yscrollcommand=scroll_v_bar.set,
                                       xscrollcommand=scroll_h_bar.set)
        scroll_v_bar.pack(side=tk.RIGHT, fill=tk.Y)  # 设置垂直滚动条显示的位置
        scroll_v_bar.config(command=self.main_listbox.yview)  # 设置Scrollbar组件的command选项为该组件的yview()方法
        scroll_h_bar.pack(side=tk.BOTTOM, fill=tk.X)  # 设置水平滚动条显示的位置
        scroll_h_bar.config(command=self.main_listbox.xview)  # 设置Scrollbar组件的command选项为该组件的xview()方法
        self.main_listbox.place(relx=0, rely=0, relheight=0.95, relwidth=0.952, bordermode='ignore')
        list_frame.place(relx=0, rely=0, relheight=1, relwidth=1, bordermode='ignore')

    def register_command(self):
        # 选择文件
        self.btn_import_files.bind('<Button>', self.choose_file)
        # 选择文件夹
        self.btn_import_dir.bind('<Button>', self.choose_dir)
        # 删除列表选中项
        self.btn_delete_list_items.bind('<Button>', self.remove_list_item)
        # 选择文件夹
        self.btn_choose_position.bind('<Button>', self.choose_export_dir)
        # 导出按钮
        self.btn_export.bind('<Button>', self.export)
        # ---------------------------------------------------------------------------
        self.bind_cb_evt(self.cb_name_part1, self.name_tips)
        self.bind_cb_evt(self.cb_name_part2, self.name_tips)
        self.bind_cb_evt(self.cb_name_part3, self.name_tips)
        self.bind_cb_evt(self.cb_name_part4, self.name_tips)
        # ---------------------------------------------------------------------------
        self.bind_cb_evt(self.cb_save_position, self.dir_tips)
        self.bind_cb_evt(self.entry_save_position, self.dir_tips)
        # ---------------------------------------------------------------------------
        self.cb_delete_raw_file.bind('<Button>', self.delete_tips)
        # ---------------------------------------------------------------------------
        self.cb_export_type_text.bind('<Button>', self.export_tips_text)
        self.cb_export_type_image.bind('<Button>', self.export_tips_image)
        self.cb_export_type_attachment.bind('<Button>', self.export_tips_attachment)
        self.cb_export_type_table.bind('<Button>', self.export_tips_table)
        self.cb_export_type_combine_text_table.bind('<Button>', self.export_tips_combine_text_table)
        self.cb_export_type_info.bind('<Button>', self.export_tips_info)

    @staticmethod
    def bind_cb_evt(cb, evt):
        cb.bind('<Button>', evt)
        cb.bind('<<ComboboxSelected>>', evt)
        cb.bind('<space>', evt)
        cb.bind('<Return>', evt)
        cb.bind('<Key>', evt)

    def export_tips_image(self, evt):
        if not evt:
            return
        # 这个值是点击之前的值，
        if not self.che_image.get():
            self.entry_tips_val.set('导出Word中的图片')
        else:
            self.entry_tips_val.set('不导出Word中的图片')

    def export_tips_attachment(self, evt):
        if not evt:
            return
        # 这个值是点击之前的值，
        if not self.che_attachment.get():
            self.entry_tips_val.set('导出Word中的附件')
        else:
            self.entry_tips_val.set('不导出Word中的附件')

    def export_tips_table(self, evt):
        if not evt:
            return
        # 这个值是点击之前的值，
        if not self.che_table.get():
            self.entry_tips_val.set('导出Word表格中的文字')
        else:
            self.entry_tips_val.set('不导出Word表格中的文字')

    def export_tips_info(self, evt):
        if not evt:
            return
        # 这个值是点击之前的值，
        if not self.che_info.get():
            self.entry_tips_val.set('导出Word文档的信息（作者、修改时间啥的）')
        else:
            self.entry_tips_val.set('不导出Word文档的信息')

    def export_tips_combine_text_table(self, evt):
        if not evt:
            return
        # 这个值是点击之前的值，
        if not self.che_combine.get():
            self.entry_tips_val.set('合并导出Word中的普通文字和表格文字')
        else:
            self.entry_tips_val.set('分别导出Word中的普通文字和表格文字')

    def export_tips_text(self, evt):
        if not evt:
            return
        # 这个值是点击之前的值，
        if not self.che_text.get():
            self.entry_tips_val.set('导出Word中的普通文字')
        else:
            self.entry_tips_val.set('不导出Word中的普通文字')

    def delete_tips(self, evt):
        if not evt:
            return
        # 这个值是点击之前的值，
        if not self.che_delete_raw.get():
            self.entry_tips_val.set('导出成功后立即删除原文件')
        else:
            self.entry_tips_val.set('导出成功后保留原文件')

    def dir_tips(self, evt):
        if not evt:
            return
        way = self.combobox_save_path.get()
        if way == self.cb_save_position['value'][0]:
            # 原文件子目录
            self.entry_tips_val.set('会在原文件所在目录创建一个同名子文件夹以存储')
        elif way == self.cb_save_position['value'][1]:
            # 原文件目录
            self.entry_tips_val.set('保存位置与原文件在同一文件夹')
        else:
            export_dir = self.entry_save_position_val.get()
            # 指定目录
            self.entry_tips_val.set('保存位置：' + export_dir)

    def name_tips(self, evt):
        if not evt:
            return
        name = self.combobox_name1.get()
        name += self.combobox_name2.get()
        name += self.combobox_name3.get()
        name += self.combobox_name4.get()

        name = name.replace('|自增编号|', '编号') \
            .replace('|连接符|', ' - ') \
            .replace('|原文件名|', '原文件名') \
            .replace('|后缀名|', '.后缀') \
            .strip()

        self.entry_tips_val.set('文件名格式：%s' % name)

    def choose_export_dir(self, evt):
        if not evt:
            return
        directory = filedialog.askdirectory()
        if directory:
            self.entry_save_position_val.set(directory)
        self.dir_tips(None)

    def remove_list_item(self, evt):
        if not evt:
            return
        if len(self.main_listbox.curselection()) == 0:
            showwarning('错误', '没有选中任何条目')
            self.entry_tips_val.set('没有从列表中移除任何文件')
        else:
            success = 0
            remove_index_list = []
            for index in self.main_listbox.curselection():
                remove_index_list.append(index)
                self.main_listbox.select_clear(index)

            # 需要从后面开始删除
            remove_index_list.sort(reverse=True)

            remove_file_list = []
            for index in remove_index_list:
                item_text = self.main_listbox.get(index)
                if item_text:
                    file_path = str(item_text).split('|-', 1)[1]
                    remove_file_list.append(file_path)
                    self.main_listbox.delete(index)

            for file_path in remove_file_list:
                self.file_list.remove(file_path)
                success += 1

            self.entry_tips_val.set('已从列表移除{0}个文件'.format(success))

    def add_file_list(self, files: list):
        success = 0
        print(self.file_list)

        # 序号前面补 0
        bit = len(str(len(files)))
        if bit < 4:
            bit = 4

        for file in files:
            # 已添加，不需要再添加
            if file in self.file_list:
                continue
            self.file_index += 1
            logger.debug('添加文件：%s', file)
            self.main_listbox.insert('end', str(self.file_index).rjust(bit, '0') + '|-' + file)  # 从最后一个位置开始加入值
            self.file_list.append(file)
            success += 1
        showinfo('添加结果', '成功添加{0}个docx文件'.format(success))
        self.entry_tips_val.set('成功添加{0}个docx文件，失败{1}个'.format(success, len(files) - success))

    def choose_file(self, evt):
        if not evt:
            return
        files_tuple = filedialog \
            .askopenfilename(title='请选择docx文件', filetypes=[('Word', '.docx')],
                             defaultextension='.docx',
                             multiple=True)
        if files_tuple:
            file_list = []
            for file in files_tuple:
                if file_check(file):
                    file_list.append(file)  # 添加到列表中
            self.add_file_list(file_list)

    def choose_dir(self, evt):
        if not evt:
            return
        is_choose_son = askyesno('选择文件夹', '选择文件夹时是否选择子文件夹内的文件？')
        directory = filedialog.askdirectory()
        if directory:
            file_list = []
            logger.debug('添加子文件夹中的内容：%s 选择目录：%s', is_choose_son, directory)
            for root, dirs, files in os.walk(directory):  # 遍历目录
                if directory != root and not is_choose_son:  # 跳过子文件夹
                    continue
                logger.debug('添加目录：%s', root)
                for file in files:  # 遍历文件
                    file_path = os.path.join(root, file)  # 拼接路径
                    if file_check(file_path):
                        file_list.append(file_path)  # 添加到列表中
            self.add_file_list(file_list)

    def export(self, evt):
        if not evt:
            return
        export_dir_choose = self.combobox_save_path.get()
        logger.debug('导出位置：%s', export_dir_choose)

        try:
            export_dir_choose = self.cb_save_position['value'].index(export_dir_choose) + 1
        except ValueError:
            showwarning('出错', '保存方式参数错误')
            return

        logger.debug('导出方式：%s', export_dir_choose)

        export_dir = self.entry_save_position_val.get()
        logger.debug('导出目录：%s', export_dir)

        if export_dir_choose == self.cb_save_position['value'][2]:
            if not (export_dir and os.path.isdir(export_dir)):
                showwarning('导出时遇到问题', '保存位置未设置文件夹，请设置保存文件夹或设置其他保存位置')
                return

        name = self.combobox_name1.get()
        name += self.combobox_name2.get()
        name += self.combobox_name3.get()
        name += self.combobox_name4.get()
        logger.debug('名字规则：%s', name)
        if not name.endswith('|后缀名|'):
            if not askyesno('提示', '文件名最好以“|后缀名|”结尾，否则可能导致无法识别，是否继续？'):
                return

        export_type = []
        if self.che_text.get():
            export_type.append('文本')
        if self.che_table.get():
            export_type.append('表格')
        if self.che_image.get():
            export_type.append('图片')
        if self.che_attachment.get():
            export_type.append('附件')
        if self.che_combine.get():
            export_type.append('合并')
        if self.che_info.get():
            export_type.append('信息')

        logger.debug('导出类型：%s', str(export_type))

        is_delete = self.che_delete_raw.get()
        logger.debug('导出后删除原文件：%s', str(is_delete))

        parameter = {
            '保存方式': export_dir_choose,
            '保存目录': export_dir,
            '文件名格式': name,
            '导出类型': export_type,
            '导出后删除原文件': bool(is_delete),
            '导出文件': self.file_list
        }

        res = self.start_export(parameter)
        logger.debug('导出结果：%s', str(res))
        self.entry_tips_val.set(res)

    def start_export(self, parameter):
        export.export_dialog = export.Export(master=self.master, parameter=parameter)
        self.master.wait_window(export.export_dialog)
        return export.export_dialog.result

def run():
    root = tk.Tk()
    Application(master=root, version=v)

    def close_window():
        ans = askyesno(title='提示', message='是否关闭窗口？')
        if ans:
            root.destroy()
        else:
            return

    root.protocol('WM_DELETE_WINDOW', close_window)

    root.title('导出docx')
    root.iconbitmap('images/icon.ico')
    root.mainloop()


if __name__ == '__main__':
    logger = log_create()
    run()

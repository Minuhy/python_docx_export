import logging
import os
import tkinter as tk
from tkinter import ttk
from datetime import datetime
from tkinter import filedialog
from tkinter.messagebox import askokcancel, askyesno, showinfo, showwarning

from gui import ApplicationGUI

v = 'v 0.1.0.1'

# APP根目录
app_path = os.path.dirname(os.path.abspath(__file__))
# 日志
logger = logging.getLogger()

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
logger.setLevel(logging.DEBUG)

# 将处理器添加至日志器中
logger.addHandler(console_handler)
logger.addHandler(file_handler)

print('日志位置：', log_file_path)


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

    def create_list(self, frame):
        # 一个列表
        list_frame = tk.Frame(frame)
        scroll_h_bar = tk.Scrollbar(list_frame, orient=tk.HORIZONTAL)  # 水平滚动条组件
        scroll_v_bar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)  # 垂直滚动条组件
        self.main_listbox = tk.Listbox(list_frame,
                                       width=60, height=20,
                                       selectmode=tk.MULTIPLE,
                                       yscrollcommand=scroll_v_bar.set,
                                       xscrollcommand=scroll_h_bar.set)
        scroll_v_bar.pack(side=tk.RIGHT, fill=tk.Y)  # 设置垂直滚动条显示的位置
        scroll_v_bar.config(command=self.main_listbox.yview)  # 设置Scrollbar组件的command选项为该组件的yview()方法
        scroll_h_bar.pack(side=tk.BOTTOM, fill=tk.X)  # 设置水平滚动条显示的位置
        scroll_h_bar.config(command=self.main_listbox.xview)  # 设置Scrollbar组件的command选项为该组件的xview()方法
        self.main_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        list_frame.place(relx=0, rely=0, relheight=1, relwidth=1, bordermode='ignore')

    def create_left(self, frame):
        """
        左边布局，文件列表
        :return: None
        """

    def create_right(self, frame):
        """
        右边布局
        :return: None
        """
        main_right_frame = tk.Frame(frame)

        # 导出设置=====================================================================================================
        export_setting_frame = tk.LabelFrame(main_right_frame, text='导出设置')
        # 导出下拉框
        tk.Label(export_setting_frame, text="保存位置：").grid(row=0, column=0)
        # 创建下拉菜单
        combobox = ttk.Combobox(export_setting_frame, width=32, state="readonly")
        # 设置下拉菜单中的值
        combobox['value'] = ('原文件目录中以原文件命名的子文件夹中', '原文件所在文件夹', '自定义文件夹')
        # 使用 grid() 来控制控件的位置
        combobox.grid(row=0, column=1)
        # 通过 current() 设置下拉菜单选项的默认值
        combobox.current(0)

        # 选择导出目录按钮
        tk.Button(export_setting_frame,
                  text="选择导出文件夹",
                  width=15, height=1,
                  command=self.choose_export_dir) \
            .grid(row=0, column=2)

        # 导出目录输入框
        tk.Entry(export_setting_frame,
                 width=60,
                 textvariable=self.entry_export_dir_val) \
            .grid(row=1, column=0, columnspan=3)

        # 导出下拉框
        tk.Label(export_setting_frame, text="导出文件名设置：") \
            .grid(row=2, column=0)

        # 创建下拉菜单
        combobox = ttk.Combobox(export_setting_frame, width=8)
        # 设置下拉菜单中的值
        combobox['value'] = ('|自增编号|', '|原文件名|', '|后缀名|', '|连接符|')
        # 使用 grid() 来控制控件的位置
        combobox.grid(row=3, column=0)
        # 通过 current() 设置下拉菜单选项的默认值
        combobox.current(0)

        # 创建下拉菜单
        combobox = ttk.Combobox(export_setting_frame, width=8)
        # 设置下拉菜单中的值
        combobox['value'] = ('|自增编号|', '|原文件名|', '|后缀名|', '|连接符|')
        # 使用 grid() 来控制控件的位置
        combobox.grid(row=3, column=1)
        # 通过 current() 设置下拉菜单选项的默认值
        combobox.current(0)

        # 创建下拉菜单
        combobox = ttk.Combobox(export_setting_frame, width=8)
        # 设置下拉菜单中的值
        combobox['value'] = ('|自增编号|', '|原文件名|', '|后缀名|', '|连接符|')
        # 使用 grid() 来控制控件的位置
        combobox.grid(row=3, column=2)
        # 通过 current() 设置下拉菜单选项的默认值
        combobox.current(0)

        # 创建下拉菜单
        combobox = ttk.Combobox(export_setting_frame, width=8)
        # 设置下拉菜单中的值
        combobox['value'] = ('|自增编号|', '|原文件名|', '|后缀名|', '|连接符|')
        # 使用 grid() 来控制控件的位置
        combobox.grid(row=3, column=3)
        # 通过 current() 设置下拉菜单选项的默认值
        combobox.current(0)

        check_box_frame = tk.Frame(export_setting_frame)
        # 复选框控件，使用variable参数来接收变量
        check1 = tk.Checkbutton(check_box_frame, text="内容文本", variable=self.check_export_text, onvalue=1,
                                offvalue=0)
        check2 = tk.Checkbutton(check_box_frame, text="表格文本", variable=self.check_export_table, onvalue=1,
                                offvalue=0)
        check3 = tk.Checkbutton(check_box_frame, text="图片", variable=self.check_export_image, onvalue=1,
                                offvalue=0)
        check4 = tk.Checkbutton(check_box_frame, text="附件", variable=self.check_export_attachment, onvalue=1,
                                offvalue=0)
        check5 = tk.Checkbutton(check_box_frame, text="合并文本表格", variable=self.check_export_attachment, onvalue=1,
                                offvalue=0)

        # 选择第一个为默认选项

        check1.select()
        check2.select()
        check3.select()
        check4.select()
        check5.select()

        check1.grid(row=0, column=0)
        check2.grid(row=0, column=1)
        check3.grid(row=0, column=2)
        check4.grid(row=0, column=3)
        check5.grid(row=0, column=4)

        check_box_frame.grid(row=4, column=0, columnspan=3)

        export_setting_frame.pack()

        # 导出按钮=====================================================================================================

        # 导出按钮
        tk.Button(main_right_frame,
                  text="导出",
                  fg="red",
                  width=20, height=1,
                  command=self.export) \
            .pack()

        tk.Label(main_right_frame, text='批量导出docx文本、图片和附件程序\n' + self.version, font=('宋体', 18, 'bold italic'),
                 bg='#7CCD7C',
                 # 设置标签内容区大小
                 width=34, height=2,
                 # 设置填充区距离、边框宽度和其样式（凹陷式）
                 padx=10, pady=15, borderwidth=10, relief='sunken').pack(side=tk.BOTTOM)

        main_right_frame.grid(row=0, column=1, sticky='n')

    def choose_export_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_export_dir_val.set(directory)

    def export(self):
        # 没有选择任何项目的情况下
        if (check_var1.get() == 0 and check_var1.get() == 0 and check_var1.get() == 0):
            s = '您还没选择任语言'
        else:
            s1 = "Python" if check_var1.get() == 1 else ""
            s2 = "C语言" if check_var1.get() == 1 else ""
            s3 = "Java" if check_var1.get() == 1 else ""
            s = "您选择了%s %s %s" % (s1, s2, s3)

    def remove_list_item(self):
        if len(self.main_listbox.curselection()) == 0:
            showwarning('错误', '没有选中任何条目')
        else:
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
        showinfo('添加结果', '添加成功{0}个'.format(success))

    def choose_file(self):
        files_tuple = filedialog.askopenfilename(title='请选择docx文件', filetypes=[('Word', '.docx')],
                                                 defaultextension='.docx',
                                                 multiple=True)
        if files_tuple:
            file_list = []
            for file in files_tuple:
                if file_check(file):
                    file_list.append(file)  # 添加到列表中
            self.add_file_list(file_list)

    def choose_dir(self):
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


def run():
    root = tk.Tk()
    Application(master=root, version=v)

    def closeWindow():
        ans = askyesno(title='提示', message='是否关闭窗口？')
        if ans:
            root.destroy()
        else:
            return

    root.protocol('WM_DELETE_WINDOW', closeWindow)

    root.title('导出docx')
    root.iconbitmap('images/icon.ico')
    root.mainloop()


if __name__ == '__main__':
    run()

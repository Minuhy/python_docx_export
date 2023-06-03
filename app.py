import logging
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
from tkinter.messagebox import askokcancel, askyesno, showinfo

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


class Application(tk.Frame):
    def __init__(self, master=None, version=''):
        super().__init__(master)
        self.master = master
        self.pack()
        self.file_list = []

        label = tk.Label(self, text='批量导出docx文本、图片和附件程序\n' + version, font=('宋体', 18, 'bold italic'), bg='#7CCD7C',
                         # 设置标签内容区大小
                         width=34, height=2,
                         # 设置填充区距离、边框宽度和其样式（凹陷式）
                         padx=10, pady=15, borderwidth=10, relief='sunken')
        label.pack()

        # 主窗口流程
        choose_frame = tk.Frame(self)
        choose_frame.pack()

        tk.Button(choose_frame,
                  text="选择文件",
                  fg="black",
                  width=15, height=1,
                  command=self.choose_file) \
            .grid(row=0, column=0)

        tk.Button(choose_frame,
                  text="选择文件夹",
                  fg="black",
                  width=15, height=1,
                  command=self.choose_dir) \
            .grid(row=0, column=1)

        list_frame = tk.Frame(self)
        list_frame.pack()
        scroll_bar = tk.Scrollbar(list_frame)  # 垂直滚动条组件
        scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)  # 设置垂直滚动条显示的位置
        self.main_listbox = tk.Listbox(list_frame,
                                       width=60, height=20,
                                       yscrollcommand=scroll_bar.set)
        self.main_listbox.pack(side=tk.LEFT, fill=tk.BOTH)
        scroll_bar.config(command=self.main_listbox.yview)  # 设置Scrollbar组件的command选项为该组件的yview()方法

        # 退出按钮
        tk.Button(self,
                  text="退出",
                  fg="red",
                  width=15, height=1,
                  command=self.exit_app) \
            .pack(side=tk.BOTTOM)

    def add_file_list(self, file_list: list):
        success = 0
        index = len(self.file_list)

        # 序号前面补 0
        bit = len(str(index + len(file_list)))
        if bit < 4:
            bit = 4

        for file in file_list:
            index += 1
            logger.debug('添加文件：%s', file)
            self.main_listbox.insert('end', str(index).rjust(bit, '0') + ' - ' + file)  # 从最后一个位置开始加入值
            self.file_list.append(file)
            success += 1
        showinfo('添加结果', '添加成功{0}个'.format(success))

    def exit_app(self):
        if askokcancel('退出', '是否退出程序？'):
            self.master.destroy()
            quit()

    def choose_file(self):
        files_tuple = filedialog.askopenfilename(title='请选择docx文件', filetypes=[('Word', '.docx')],
                                                 defaultextension='.docx',
                                                 multiple=True)
        if files_tuple:
            file_list = []
            for file in files_tuple:
                if file_check(file):
                    self.file_list.append(file)
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
    app = Application(master=root, version=v)
    root.title('导出docx')
    root.iconbitmap('images/icon.ico')
    app.mainloop()


if __name__ == '__main__':
    run()

import os
import time
from random import Random
from tkinter.messagebox import askyesno

import log
from multiprocessing import Pool
from threading import Thread
from gui import ExportGUI



export_dialog = None

# 日志对象
logger = log.Logger().log_create()


class CopeException(BaseException):
    def __init__(self, index, file, e):
        self.index = index
        self.file = file
        self.e = e


def dispose_error_callback(e: BaseException):
    """
    异常回调
    :param e: 异常对象
    :return:  None
    """
    if not isinstance(e, CopeException):
        return
    if isinstance(export_dialog, Export):
        export_dialog.show_progress(e.index, '错误 -> {0} | {1}'.format(e.file, e.e))
        fail_file_list.append(e.file)


def dispose_callback(p: dict):
    state = p.get('state')
    file = p.get('file')
    child = p.get('child')
    index = p.get('index')

    msg = ''
    if file:
        if state:
            success_file_dict[file] = child
            msg = '成功 -> ' + file
        else:
            msg = '失败 -> ' + file

    if isinstance(export_dialog, Export):
        export_dialog.show_progress(index, msg)


def dispose(index: int, file: str, parameter: dict):
    try:
        print(index, '处理文件：', file, parameter)
        time.sleep(1 * Random().random())
        if index % 20 == 0:
            if askyesno('提示', '是否覆盖？'):
                print('覆盖')
            else:
                print('取消覆盖')
        return {
            'index': index,
            'state': 'success',
            'file': file,
            'child': ['1', '2']
        }
    except Exception as e:
        # 统一处理异常
        raise CopeException(index, file, e)


def start_task(file_list, parameter):
    p = Thread(target=run, args=(file_list, parameter))  # 实例化进程对象
    p.start()


def run(file_list, parameter):
    # 开启处理进程
    pool = Pool(5)  # 创建一个5个进程的进程池
    for index, file in enumerate(file_list):
        print('创建任务', index, file)
        pool.apply_async(func=dispose,
                         args=(index, file, parameter),
                         callback=dispose_callback,
                         error_callback=dispose_error_callback)
    if len(file_list):
        pool.close()
        pool.join()

    print('文件处理结束')
    if isinstance(export_dialog, Export):
        export_dialog.after_destroy()


class Export(ExportGUI):

    def __init__(self, master=None, parameter=None):
        super().__init__(master)
        self.master = master

        # {'文件名':['导出文件1','导出文件2',......]}
        self.success_file_dict = {}
        self.fail_file_list = []

        self.exit_time = 0  # 点两次退出
        self.pr_val = 0  # 导出进度

        self.result = '未完成'

        # **********************************************************************
        self.file_list = []
        self.export_type = ['文本', '表格', '图片', '附件', '合并', '信息']
        self.filename_format = '|自增编号||连接符||原文件名||后缀名|'
        self.save_way = 1
        self.save_dir = None
        self.delete_raw_file = False
        # **********************************************************************

        check = self.check_init(parameter)

        self.file_list_len = len(self.file_list)

        parameter = {
            '保存方式': self.save_way,
            '保存目录': self.save_dir,
            '文件名格式': self.filename_format,
            '导出类型': self.export_type,
            '导出后删除原文件': self.delete_raw_file
        }

        if check:
            start_task(file_list=self.file_list, parameter=parameter)
            self.result = self.success_file_dict
        else:
            self.after_destroy()

    def check_init(self, parameter):
        """
        检查参数和初始化参数
        :param parameter: 传入参数
        :return: 继续执行(True)或停止执行(False)
        """
        is_exit = False
        if isinstance(parameter, dict):
            obj = parameter.get('导出文件')
            if isinstance(obj, list) and obj:
                self.file_list = obj
            else:
                self.result = '没有文件需要处理'
                is_exit = True

            if parameter.get('导出后删除原文件'):
                self.delete_raw_file = True
            else:
                self.delete_raw_file = False

            obj = parameter.get('导出类型')
            if isinstance(obj, list) and obj:
                self.export_type = obj

            obj = parameter.get('文件名格式')
            if isinstance(obj, str) and str:
                self.filename_format = obj

            obj = parameter.get('保存方式')
            if isinstance(obj, int) and obj:
                self.save_way = obj

            if self.save_way == 3:
                obj = parameter.get('保存目录')
                if isinstance(obj, str) and os.path.isdir(obj):
                    self.save_dir = obj
                else:
                    self.result = '保存目录参数错误'
                    is_exit = True
        else:
            self.result = '参数不正确'
            is_exit = True
        return not is_exit

    def after_destroy(self):
        self.after(1000, self.close)

    def show_progress(self, current, message):
        if self.file_list_len != 0:
            current_percent = current / self.file_list_len * 100
        else:
            current_percent = 0
        if current_percent > self.pr_val:
            self.pr_val = current_percent
        else:
            current_percent = self.pr_val
        current_time = int(time.time())
        if current_time - self.exit_time <= 3:
            self.lb_tips_val.set('再点一次退出')
            self.lb_tips.configure(background="#f4984d")
            self.lb_tips.configure(foreground="#ffffff")
        else:
            self.lb_tips_val.set('正在处理：{0}%'.format(current_percent))
            self.lb_tips.configure(background="#f0f0f0")
            self.lb_tips.configure(foreground="#000000")
        self.entry_info_var.set(message)
        self.pb_main_val.set(current_percent)

    def close(self):
        pass

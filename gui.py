# ! /usr/bin/env python3
#  -*- coding: utf-8 -*-
#
# GUI module generated by PAGE version 7.6
#  in conjunction with Tcl version 8.6
#    Jun 05, 2023 01:52:12 PM CST  platform: Windows NT

import tkinter as tk
import tkinter.ttk as ttk


class ApplicationGUI:
    def __init__(self, top=None):
        width = 800
        height = 450
        screenwidth = top.winfo_screenwidth()
        screenheight = top.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        top.geometry(geometry)
        top.resizable(1, 1)
        top.configure(background="#f0f0f0")
        top.configure(highlightbackground="#f2f2f2")
        top.configure(highlightcolor="black")

        self.top = top
        self.combobox_save_path = tk.StringVar()
        self.combobox_export_cover = tk.StringVar()
        self.combobox_name1 = tk.StringVar()
        self.combobox_name2 = tk.StringVar()
        self.combobox_name3 = tk.StringVar()
        self.combobox_name4 = tk.StringVar()
        self.entry_save_position_val = tk.StringVar()
        self.entry_tips_val = tk.StringVar()
        self.che_text = tk.IntVar()
        self.che_table = tk.IntVar()
        self.che_image = tk.IntVar()
        self.che_attachment = tk.IntVar()
        self.che_combine = tk.IntVar()
        self.che_info = tk.IntVar()
        self.che_delete_raw = tk.IntVar()

        self.lf_file_list = tk.ttk.LabelFrame(self.top)
        self.lf_file_list.place(relx=0.013, rely=0.022, relheight=0.956, relwidth=0.5)
        self.lf_file_list.configure(text='''导出文件列表''')
        self.btn_import_files = tk.ttk.Button(self.lf_file_list)
        self.btn_import_files.place(relx=0.025, rely=0.046, height=28, width=60, bordermode='ignore')
        self.btn_import_files.configure(text='''导入文件''')
        self.btn_import_dir = tk.ttk.Button(self.lf_file_list)
        self.btn_import_dir.place(relx=0.195, rely=0.046, height=28, width=80, bordermode='ignore')
        self.btn_import_dir.configure(text='''导入文件夹''')
        self.btn_delete_list_items = tk.ttk.Button(self.lf_file_list)
        self.btn_delete_list_items.place(relx=0.775, rely=0.046, height=28, width=80, bordermode='ignore')
        self.btn_delete_list_items.configure(text='''删除选中项''')
        self.btn_delete_list_all = tk.ttk.Button(self.lf_file_list)
        self.btn_delete_list_all.place(relx=0.56, rely=0.046, height=28, width=80, bordermode='ignore')
        self.btn_delete_list_all.configure(text='''删除所有项''')
        self.frame_list = tk.Frame(self.lf_file_list)
        self.frame_list.place(relx=0.025, rely=0.118, relheight=0.858, relwidth=0.95, bordermode='ignore')
        self.frame_list.configure(relief='groove')
        self.frame_list.configure(borderwidth="2")
        self.frame_list.configure(relief="groove")
        self.frame_list.configure(background="#f0f0f0")
        self.frame_list.configure(highlightbackground="#f2f2f2")
        self.frame_list.configure(highlightcolor="black")
        self.lf_export_setting = tk.ttk.LabelFrame(self.top)
        self.lf_export_setting.place(relx=0.525, rely=0.022, relheight=0.956, relwidth=0.463)
        self.lf_export_setting.configure(labelanchor="ne")
        self.lf_export_setting.configure(text='''导出设置''')
        self.lb_save_position = tk.Label(self.lf_export_setting)
        self.lb_save_position.place(relx=0.022, rely=0.040, height=23, width=348, bordermode='ignore')
        self.lb_save_position.configure(activebackground="#f0f0f0")
        self.lb_save_position.configure(anchor='w')
        self.lb_save_position.configure(background="#f0f0f0")
        self.lb_save_position.configure(compound='left')
        self.lb_save_position.configure(disabledforeground="#a3a3a3")
        self.lb_save_position.configure(foreground="#000000")
        self.lb_save_position.configure(highlightbackground="#f2f2f2")
        self.lb_save_position.configure(highlightcolor="black")
        self.lb_save_position.configure(justify='left')
        self.lb_save_position.configure(text='''保存位置：''')

        self.lb_export = tk.Label(self.lf_export_setting)
        self.lb_export.place(relx=0.022, rely=0.515, height=23, width=348, bordermode='ignore')
        self.lb_export.configure(activebackground="#f0f0f0")
        self.lb_export.configure(anchor='w')
        self.lb_export.configure(background="#f0f0f0")
        self.lb_export.configure(compound='left')
        self.lb_export.configure(disabledforeground="#a3a3a3")
        self.lb_export.configure(foreground="#000000")
        self.lb_export.configure(highlightbackground="#f2f2f2")
        self.lb_export.configure(highlightcolor="black")
        self.lb_export.configure(justify='left')
        self.lb_export.configure(text='''导出策略：''')

        self.cb_save_position = ttk.Combobox(self.lf_export_setting)
        self.cb_save_position.place(relx=0.027, rely=0.094, height=27, relwidth=0.71, bordermode='ignore')
        self.cb_save_position.configure(textvariable=self.combobox_save_path)
        self.cb_save_position.configure(takefocus="")
        self.cb_save_position.configure(state="readonly")
        self.btn_choose_position = tk.ttk.Button(self.lf_export_setting)
        self.btn_choose_position.place(relx=0.755, rely=0.094, height=28, width=80, bordermode='ignore')
        self.btn_choose_position.configure(text='''选择文件夹''')
        self.entry_save_position = tk.ttk.Entry(self.lf_export_setting)
        self.entry_save_position.place(relx=0.027, rely=0.164, height=27, relwidth=0.941, bordermode='ignore')
        self.entry_save_position.configure(textvariable=self.entry_save_position_val)
        self.lb_save_name = tk.Label(self.lf_export_setting)
        self.lb_save_name.place(relx=0.022, rely=0.235, height=23, width=348, bordermode='ignore')
        self.lb_save_name.configure(activebackground="#f9f9f9")
        self.lb_save_name.configure(anchor='w')
        self.lb_save_name.configure(background="#f0f0f0")
        self.lb_save_name.configure(compound='left')
        self.lb_save_name.configure(disabledforeground="#a3a3a3")
        self.lb_save_name.configure(foreground="#000000")
        self.lb_save_name.configure(highlightbackground="#f2f2f2")
        self.lb_save_name.configure(highlightcolor="black")
        self.lb_save_name.configure(justify='left')
        self.lb_save_name.configure(text='''保存文件名设置：''')
        self.cb_name_part1 = ttk.Combobox(self.lf_export_setting)
        self.cb_name_part1.place(relx=0.027, rely=0.294, relwidth=0.218, bordermode='ignore')
        self.cb_name_part1.configure(textvariable=self.combobox_name1)
        self.cb_name_part1.configure(takefocus="")
        self.cb_name_part2 = ttk.Combobox(self.lf_export_setting)
        self.cb_name_part2.place(relx=0.269, rely=0.294, relwidth=0.218, bordermode='ignore')
        self.cb_name_part2.configure(textvariable=self.combobox_name2)
        self.cb_name_part2.configure(takefocus="")
        self.cb_name_part3 = ttk.Combobox(self.lf_export_setting)
        self.cb_name_part3.place(relx=0.511, rely=0.294, relwidth=0.218, bordermode='ignore')
        self.cb_name_part3.configure(textvariable=self.combobox_name3)
        self.cb_name_part3.configure(takefocus="")
        self.cb_name_part4 = ttk.Combobox(self.lf_export_setting)
        self.cb_name_part4.place(relx=0.753, rely=0.294, relwidth=0.218, bordermode='ignore')
        self.cb_name_part4.configure(textvariable=self.combobox_name4)
        self.cb_name_part4.configure(takefocus="")
        self.lb_export_type = tk.Label(self.lf_export_setting)
        self.lb_export_type.place(relx=0.027, rely=0.352, height=23, width=348, bordermode='ignore')
        self.lb_export_type.configure(activebackground="#f9f9f9")
        self.lb_export_type.configure(anchor='w')
        self.lb_export_type.configure(background="#f0f0f0")
        self.lb_export_type.configure(compound='left')
        self.lb_export_type.configure(disabledforeground="#a3a3a3")
        self.lb_export_type.configure(foreground="#000000")
        self.lb_export_type.configure(highlightbackground="#f2f2f2")
        self.lb_export_type.configure(highlightcolor="black")
        self.lb_export_type.configure(text='''导出类型设置：''')
        self.cb_export_type_text = tk.ttk.Checkbutton(self.lf_export_setting)
        self.cb_export_type_text.place(relx=0.027, rely=0.4, relheight=0.063, relwidth=0.134, bordermode='ignore')
        self.cb_export_type_text.configure(text='''文本''')
        self.cb_export_type_text.configure(variable=self.che_text)
        self.cb_export_type_table = tk.ttk.Checkbutton(self.lf_export_setting)
        self.cb_export_type_table.place(relx=0.161, rely=0.4, relheight=0.063, relwidth=0.215, bordermode='ignore')
        self.cb_export_type_table.configure(text='''表格文本''')
        self.cb_export_type_table.configure(variable=self.che_table)
        self.cb_export_type_image = tk.ttk.Checkbutton(self.lf_export_setting)
        self.cb_export_type_image.place(relx=0.027, rely=0.453, relheight=0.063, relwidth=0.142, bordermode='ignore')
        self.cb_export_type_image.configure(text='''图片''')
        self.cb_export_type_image.configure(variable=self.che_image)
        self.cb_export_type_attachment = tk.ttk.Checkbutton(self.lf_export_setting)
        self.cb_export_type_attachment.place(relx=0.161, rely=0.453, relheight=0.063, relwidth=0.142,
                                             bordermode='ignore')
        self.cb_export_type_attachment.configure(text='''附件''')
        self.cb_export_type_attachment.configure(variable=self.che_attachment)

        self.cb_export_type_info = tk.ttk.Checkbutton(self.lf_export_setting)
        self.cb_export_type_info.place(relx=0.296, rely=0.453, relheight=0.063, relwidth=0.277,
                                       bordermode='ignore')
        self.cb_export_type_info.configure(text='''文档信息''')
        self.cb_export_type_info.configure(variable=self.che_info)

        self.cb_export_type_combine_text_table = tk.ttk.Checkbutton(self.lf_export_setting)
        self.cb_export_type_combine_text_table.place(relx=0.58, rely=0.4, relheight=0.063, relwidth=0.277,
                                                     bordermode='ignore')
        self.cb_export_type_combine_text_table.configure(text='''文本表格合并''')
        self.cb_export_type_combine_text_table.configure(variable=self.che_combine)
        self.btn_export = tk.ttk.Button(self.lf_export_setting)
        self.btn_export.place(relx=0.806, rely=0.658, height=28, width=60, bordermode='ignore')
        self.btn_export.configure(text='''导出''')

        self.cb_export_cover = ttk.Combobox(self.lf_export_setting)
        self.cb_export_cover.place(relx=0.027, rely=0.57, height=27, relwidth=0.5, bordermode='ignore')
        self.cb_export_cover.configure(textvariable=self.combobox_export_cover)
        self.cb_export_cover.configure(takefocus="")
        self.cb_export_cover.configure(state="readonly")

        self.cb_delete_raw_file = tk.ttk.Checkbutton(self.lf_export_setting)
        self.cb_delete_raw_file.place(relx=0.58, rely=0.57, relheight=0.063, relwidth=0.4, bordermode='ignore')
        self.cb_delete_raw_file.configure(text='''导出成功后删除原文件''')
        self.cb_delete_raw_file.configure(variable=self.che_delete_raw)
        self.lb_brand = tk.Label(self.lf_export_setting, text='批量导出docx文本、图片和附件程序\n',
                                 font=('宋体', 14, 'bold italic'),
                                 bg='#7CCD7C',
                                 # 设置标签内容区大小
                                 width=34, height=2,
                                 # 设置填充区距离、边框宽度和其样式（凹陷式）
                                 padx=10, pady=15, borderwidth=10, relief='sunken')
        self.lb_brand.place(relx=0.027, rely=0.785, relwidth=0.9, height=82, width=17, bordermode='ignore')
        self.lb_tips = tk.Entry(self.lf_export_setting)
        self.lb_tips.configure(textvariable=self.entry_tips_val)
        self.lb_tips.configure(bg='#ffeeee')
        self.lb_tips.configure(selectbackground='#ff2121')
        self.lb_tips.configure(selectforeground='#000000')
        self.lb_tips.place(relx=0.027, rely=0.734, relwidth=0.9, height=20, width=17, bordermode='ignore')


class ExportGUI:
    def __init__(self, top=None):
        top.configure(background="#ededed")
        self.top = top
        self.lb_d_tips_var = tk.StringVar()
        self.pd_d_main_var = tk.IntVar()
        self.cb_d_all_var = tk.IntVar()
        self.btn_d_right_var = tk.StringVar()
        self.btn_d_left_var = tk.StringVar()
        self.btn_d_e_var = tk.StringVar()

        self.tf_d_title = ttk.Labelframe(self.top)
        self.tf_d_title.configure(relief='sunken')
        self.tf_d_title.configure(labelanchor="n")
        self.tf_d_title.configure(text='''文件处理''')
        self.tf_d_title.place(relx=0.013, rely=0.022, relheight=0.956, relwidth=0.975)
        self.tf_d_title.configure(relief="sunken")
        self.pb_d_main = ttk.Progressbar(self.tf_d_title)
        self.pb_d_main.place(relx=0.013, rely=0.047, relwidth=0.974, relheight=0.0, height=22, bordermode='ignore')
        self.pb_d_main.configure(length="759")
        self.pb_d_main.configure(variable=self.pd_d_main_var)
        self.pb_d_main.configure(value='20')
        self.lb_d_tips = ttk.Label(self.tf_d_title)
        self.lb_d_tips.place(relx=0.013, rely=0.105, height=21, width=760, bordermode='ignore')
        self.lb_d_tips.configure(background="#eaeaea")
        self.lb_d_tips.configure(foreground="#000000")
        self.lb_d_tips.configure(font="TkDefaultFont")
        self.lb_d_tips.configure(relief="flat")
        self.lb_d_tips.configure(anchor='w')
        self.lb_d_tips.configure(justify='center')
        self.lb_d_tips.configure(text='''处理进度：96.2144455%''')
        self.lb_d_tips.configure(textvariable=self.lb_d_tips_var)
        self.lb_d_tips_var.set('''处理进度：∞%''')
        self.lb_d_tips.configure(compound='left')
        self.frame_d_result_show = ttk.Frame(self.tf_d_title)
        self.frame_d_result_show.place(relx=0.013, rely=0.163, relheight=0.744, relwidth=0.974, bordermode='ignore')
        self.frame_d_result_show.configure(relief='groove')
        self.frame_d_result_show.configure(borderwidth="2")
        self.frame_d_result_show.configure(relief="groove")
        self.txt_d_result_show = tk.Text(self.frame_d_result_show)
        self.txt_d_result_show.place(relx=0.013, rely=0.034, relheight=0.931, relwidth=0.974)
        self.txt_d_result_show.configure(background="white")
        self.txt_d_result_show.configure(font="TkTextFont")
        self.txt_d_result_show.configure(foreground="black")
        self.txt_d_result_show.configure(highlightbackground="#d9d9d9")
        self.txt_d_result_show.configure(highlightcolor="black")
        self.txt_d_result_show.configure(insertbackground="black")
        self.txt_d_result_show.configure(selectbackground="#c4c4c4")
        self.txt_d_result_show.configure(selectforeground="black")
        self.txt_d_result_show.configure(wrap="word")

        self.txt_d_ask_show = tk.Text(self.frame_d_result_show)
        self.txt_d_ask_show.place(relx=0, rely=0, relheight=1, relwidth=1)
        self.txt_d_ask_show.configure(background="#fb7299")
        self.txt_d_ask_show.configure(font="-family {宋体} -size 23 -weight bold -underline 1")
        self.txt_d_ask_show.configure(foreground="#ffffff")
        self.txt_d_ask_show.configure(selectbackground="#c4c4c4")
        self.txt_d_ask_show.configure(selectforeground="black")
        self.txt_d_ask_show.configure(wrap="word")
        self.txt_d_ask_show.insert('end', '导出')

        self.btn_d_e = ttk.Button(self.tf_d_title)
        self.btn_d_e.place(relx=0.6, rely=0.925, height=27, width=87, bordermode='ignore')
        self.btn_d_e.configure(takefocus="")
        self.btn_d_e.configure(textvariable=self.btn_d_e_var)
        self.btn_d_e.configure(compound='left')
        self.btn_d_left = ttk.Button(self.tf_d_title)
        self.btn_d_left.place(relx=0.738, rely=0.925, height=27, width=87, bordermode='ignore')
        self.btn_d_left.configure(takefocus="")
        self.btn_d_left.configure(textvariable=self.btn_d_left_var)
        self.btn_d_left.configure(compound='left')
        self.btn_d_right = ttk.Button(self.tf_d_title)
        self.btn_d_right.place(relx=0.876, rely=0.925, height=27, width=87, bordermode='ignore')
        self.btn_d_right.configure(takefocus="")
        self.btn_d_right.configure(textvariable=self.btn_d_right_var)
        self.btn_d_right.configure(compound='left')
        self.cb_d_all = ttk.Checkbutton(self.tf_d_title)
        self.cb_d_all.place(relx=0.013, rely=0.93, relwidth=0.4, relheight=0.0, height=23, bordermode='ignore')
        self.cb_d_all.configure(variable=self.cb_d_all_var)
        self.cb_d_all.configure(takefocus="")
        self.cb_d_all.configure(text='''不再提示，一律执行此操作''')
        self.cb_d_all.configure(compound='left')

    def close(self):
        self.tf_d_title.destroy()

    def switch_func(self, btn_e=None, btn_l='暂停', btn_r='取消', cb=None, tip_pad=None):
        # 扩展按钮
        if btn_e:
            self.btn_d_e.place(relx=0.6, rely=0.925, height=27, width=87, bordermode='ignore')
            self.btn_d_e_var.set(str(btn_e))
        else:
            self.btn_d_e.place(relx=0, rely=0, height=0, width=0, bordermode='ignore')

        # 左边按钮
        if btn_l:
            self.btn_d_left.place(relx=0.738, rely=0.925, height=27, width=87, bordermode='ignore')
            self.btn_d_left_var.set(str(btn_l))
        else:
            self.btn_d_left.place(relx=0.0, rely=0.0, height=0, width=0, bordermode='ignore')

        # 右边按钮
        if btn_r:
            self.btn_d_right.place(relx=0.876, rely=0.925, height=27, width=87, bordermode='ignore')
            self.btn_d_right_var.set(str(btn_r))
        else:
            self.btn_d_right.place(relx=0.0, rely=0.0, height=0, width=0, bordermode='ignore')

        # 多选框
        if cb:
            self.cb_d_all.place(relx=0.013, rely=0.93, relwidth=0.4, relheight=0.0, height=23, bordermode='ignore')
            self.cb_d_all.configure(text=str(cb))
        else:
            self.cb_d_all.place(relx=0, rely=0, relwidth=0, relheight=0.0, height=0, bordermode='ignore')

        # 提示面板
        if tip_pad:
            self.txt_d_ask_show.place(relx=0, rely=0, relheight=1, relwidth=1)
            self.txt_d_ask_show.delete('1.0', 'end')
            self.txt_d_ask_show.insert('end', str(tip_pad))
        else:
            self.txt_d_ask_show.place(relx=0, rely=0, relheight=0, relwidth=0)

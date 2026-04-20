#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""桌牌生成器 - 图形界面版"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

# 尝试导入 tkinterdnd2，不可用时降级为按钮选择
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

from card_core import read_names_from_excel, generate_pdf, create_template_excel


class NameCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title('桌牌生成器')
        self.root.geometry('480x380')
        self.root.resizable(False, False)

        # 居中窗口
        self.root.update_idletasks()
        w, h = 480, 380
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f'{w}x{h}+{x}+{y}')

        self.excel_path = tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        # 标题
        tk.Label(self.root, text='桌牌生成器', font=('PingFang SC', 20, 'bold')).pack(pady=(20, 5))
        tk.Label(self.root, text='将Excel文件拖入下方区域，或点击选择文件', font=('PingFang SC', 10), fg='gray').pack()

        # 拖放区域
        drop_frame = tk.Frame(self.root, bd=2, relief='groove', bg='#f0f0f0', height=120)
        drop_frame.pack(padx=30, pady=15, fill='x')
        drop_frame.pack_propagate(False)

        drop_label = tk.Label(drop_frame, text='拖入Excel文件到这里\n或点击此处选择文件',
                              font=('PingFang SC', 12), fg='#888', bg='#f0f0f0', cursor='hand2')
        drop_label.pack(expand=True, fill='both')

        if HAS_DND:
            drop_frame.drop_target_register(DND_FILES)
            drop_frame.dnd_bind('<<Drop>>', self._on_drop)
        drop_label.bind('<Button-1>', lambda e: self._choose_file())

        # 显示选中文件
        self.file_label = tk.Label(self.root, textvariable=self.excel_path,
                                   font=('PingFang SC', 9), fg='#333', wraplength=420, anchor='w')
        self.file_label.pack(padx=35, fill='x')

        # 进度条
        self.progress = tk.Frame(self.root, height=6, bg='#ddd')
        self.progress.pack(padx=35, pady=(8, 0), fill='x')
        self.progress_bar = tk.Frame(self.progress, bg='#4a90d9', width=0)
        self.progress_bar.place(x=0, y=0, relheight=1, width=0)

        self.status_label = tk.Label(self.root, text='', font=('PingFang SC', 9), fg='gray')
        self.status_label.pack(pady=2)

        # 按钮行
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        self.gen_btn = tk.Button(btn_frame, text='生成桌牌PDF', font=('PingFang SC', 11),
                                 bg='#4a90d9', fg='white', width=14, cursor='hand2',
                                 command=self._generate)
        self.gen_btn.pack(side='left', padx=8)

        self.tpl_btn = tk.Button(btn_frame, text='下载填写模板', font=('PingFang SC', 11),
                                 bg='#5cb85c', fg='white', width=14, cursor='hand2',
                                 command=self._download_template)
        self.tpl_btn.pack(side='left', padx=8)

    def _on_drop(self, event):
        path = event.data.strip('{}')
        if not path.lower().endswith(('.xlsx', '.xls')):
            messagebox.showwarning('文件格式错误', '请拖入Excel文件（.xlsx 或 .xls）')
            return
        self.excel_path.set(path)
        self.status_label.config(text=f'已选择: {os.path.basename(path)}', fg='#333')

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title='选择Excel文件',
            filetypes=[('Excel文件', '*.xlsx *.xls'), ('所有文件', '*.*')]
        )
        if path:
            self.excel_path.set(path)
            self.status_label.config(text=f'已选择: {os.path.basename(path)}', fg='#333')

    def _download_template(self):
        save_path = filedialog.asksaveasfilename(
            title='保存模板文件',
            defaultextension='.xlsx',
            initialfile='桌牌姓名模板.xlsx',
            filetypes=[('Excel文件', '*.xlsx')]
        )
        if save_path:
            try:
                create_template_excel(save_path)
                messagebox.showinfo('成功', f'模板已保存到:\n{save_path}\n\n请在「姓名」列填写姓名后使用。')
            except Exception as e:
                messagebox.showerror('错误', f'保存模板失败:\n{e}')

    def _update_progress(self, current, total, message):
        self.root.after(0, self._set_progress, current, total, message)

    def _set_progress(self, current, total, message):
        frac = current / total if total > 0 else 0
        bar_width = int(self.progress.winfo_width() * frac)
        self.progress_bar.place(x=0, y=0, relheight=1, width=bar_width)
        self.status_label.config(text=message, fg='#4a90d9')

    def _generate(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning('未选择文件', '请先拖入或选择Excel文件')
            return
        if not os.path.exists(path):
            messagebox.showerror('文件不存在', f'文件未找到:\n{path}')
            return

        # 选择保存位置
        base_name = os.path.splitext(os.path.basename(path))[0]
        default_name = f'{base_name}_桌牌.pdf'
        save_path = filedialog.asksaveasfilename(
            title='保存桌牌PDF',
            defaultextension='.pdf',
            initialfile=default_name,
            filetypes=[('PDF文件', '*.pdf')]
        )
        if not save_path:
            return

        # 禁用按钮
        self.gen_btn.config(state='disabled')
        self.tpl_btn.config(state='disabled')

        def task():
            try:
                names = read_names_from_excel(path)
                if not names:
                    self.root.after(0, lambda: messagebox.showwarning('无数据', 'Excel文件中未找到有效姓名'))
                    return
                generate_pdf(names, save_path, progress_callback=self._update_progress)
                self.root.after(0, lambda: self._on_success(save_path, len(names)))
            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))
            finally:
                self.root.after(0, self._restore_buttons)

        threading.Thread(target=task, daemon=True).start()

    def _on_success(self, save_path, count):
        self._set_progress(1, 1, '生成完成！')
        messagebox.showinfo('成功',
                            f'已生成 {count} 个桌牌\nPDF已保存到:\n{save_path}\n\n'
                            '打印提示:\n1. 选择A4纸张\n2. 沿裁剪线裁剪\n3. 对折使用')

    def _on_error(self, msg):
        messagebox.showerror('生成失败', f'出错:\n{msg}')

    def _restore_buttons(self):
        self.gen_btn.config(state='normal')
        self.tpl_btn.config(state='normal')


def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = NameCardApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()

import os
import shutil
import threading
import tkinter as tk
from copy import deepcopy
from pathlib import Path
from typing import Optional

import chardet
import pylrc
import windnd
from win32com.shell import shell, shellcon


def transform_lrc(input: Path, output: Optional[Path] = None, ops: str = 'add', file_type: str = 'lrc',
                  deleted: bool = False) -> None:
    # show(f"--处理lrc:{input}")
    if 'original_lrc' in input.parts:
        return
    if deleted:
        mv_to_trush(input)
    else:
        if not (input.parent / 'original_lrc').exists():
            Path.mkdir(input.parent / 'original_lrc')
        if not (input.parent / 'original_lrc' / input.name).exists():
            shutil.copy(input, input.parent / 'original_lrc' / input.name)
    if output is None:
        output = input
    with open(input, 'rb') as f:
        result = chardet.detect(f.read())
    lrc_file = open(input,encoding=result['encoding'])
    lrc_string = ''.join(lrc_file.readlines())
    lrc_file.close()
    subs_output = pylrc.parse('')
    subs_input = pylrc.parse(lrc_string)
    first_line = 0

    if ops == 'add':
        for sub in subs_input:
            if first_line == 0:
                subs_output.append(sub)
                first_line = 1
                continue
            if sub.text.strip() != '':
                sub_insert = deepcopy(sub)
                sub_insert.shift(milliseconds=-1)
                sub_insert.text = ''
                subs_output.append(sub_insert)
                subs_output.append(sub)
    elif ops == 'delete':
        for sub in subs_input:
            if sub.text.strip() != '':
                subs_output.append(sub)
    if file_type == 'srt':
        lrc_string = subs_output.toSRT()
        output = output.with_suffix('.srt')
    elif file_type == 'lrc':
        lrc_string = subs_output.toLRC()

    lrc_file = open(output, 'w',encoding='utf-8')
    lrc_file.write(lrc_string)
    lrc_file.close()


def dragged_files(files):
    global filesname
    global flag
    filesname = files
    flag = True


def process():
    global flag
    global filesname
    while (True):
        if flag == True:
            for filename in filesname:
                try:
                    filename = Path(filename)
                    global ops_checked
                    ops = 'add' if ops_checked.get() else 'delete'
                    global type_checked
                    file_type = 'srt' if type_checked.get() else 'lrc'
                    if filename.is_file():
                        transform_lrc(filename, ops=ops, file_type=file_type)
                    elif filename.is_dir():
                        for file in Path(filename).rglob("*.lrc"):
                            transform_lrc(file, ops=ops, file_type=file_type)
                except Exception as e:
                    show(f"{e}")
        flag = False


def show(info):
    global info_text
    info_text.insert('end', info + "\n")
    info_text.see("end")
    info_text.update()


def mv_to_trush(filename):
    try:
        filename = filename.__str__()
        res = shell.SHFileOperation((0, shellcon.FO_DELETE, filename, None,
                                     shellcon.FOF_SILENT | shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION, None,
                                     None))  # 删除文件到回收站
        if not res[1]:
            os.system('del ' + filename)
    except:
        pass


def checkbox_register(text='NULL', value=1):
    checked = tk.IntVar(value=value)
    button = tk.Checkbutton(text=text, variable=checked)
    button.pack()
    return checked


def label_register(args, **kwargs):
    lable = tk.Label(args, **kwargs)
    lable.pack()


if __name__ == '__main__':
    global filesname
    global flag
    flag = False
    filesname = ''

    window = tk.Tk()
    window.title('更改lrc')
    window.geometry('400x400')
    window.update()

    label_register(window, text='将待处理的文件夹拖入窗口,可一次拖入多个文件夹.', font=('宋体', 12), wraplength=window.winfo_width())
    label_register(window, text='勾选”添加空行“会在歌词中间添加一个空行'
                                , font=('宋体', 12), fg='red',
                   wraplength=window.winfo_width())
    label_register(window, text='取消勾选”添加空行“会删除空行'
                   , font=('宋体', 12), fg='red',
                   wraplength=window.winfo_width())
    label_register(window, text='勾选”转换srt会生成srt文件“'
                   , font=('宋体', 12), fg='red',
                   wraplength=window.winfo_width())
    label_register(window, text='原始的lrc会保存在orignal_lrc中'
                   , font=('宋体', 12), fg='red',
                   wraplength=window.winfo_width())
    global ops_checked
    ops_checked = checkbox_register('添加空行')
    global type_checked
    type_checked = checkbox_register('转换为srt', value=0)

    global info_text
    info_text = tk.Text()
    scroll = tk.Scrollbar()
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    scroll.config(command=info_text.yview)
    info_text.config(yscrollcommand=scroll.set)
    info_text.pack()

    thread = threading.Thread(target=process, daemon=True)
    thread.start()
    windnd.hook_dropfiles(window, func=dragged_files, force_unicode='utf-8')
    tk.mainloop()

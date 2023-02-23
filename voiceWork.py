import Levenshtein
import chardet
import collections
import csv
import logging
import mimetypes
import os
import pylrc
import re
import shutil
import srt
import stat
import threading
import tkinter as tk
import traceback
import urllib.request
import windnd
from PIL import Image as img
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, wait
from copy import deepcopy
from functools import wraps
from mutagen.id3 import ID3FileType, APIC, TIT2, TPE1, TALB, TPE2, TRCK
from pathlib import Path
from pydub import AudioSegment
from subprocess import run
from time import sleep
from typing import Optional
from vtt_to_srt.vtt_to_srt import ConvertFile
from win32com.shell import shell, shellcon

from translate import translate


# 拖拽时执行的函数


def dragged_files(files):
    for filename in files:
        pool.submit(process, filename)

# 处理函数
def process(filename):
    try:
        show(f"开始处理{filename}")
        filename = Path(filename)
        filename = unzip(filename)
        filename = clear(filename)
        if work_mode.get()==1:
            show(f"--处理完成，文件位于{filename}")
            return

        extract_mp3_from_video(filename)

        RJ = RJ_No(filename)
        if not RJ:
            show(f"--处理完成，文件位于{filename}")
            return

        for filename, id in RJ.items():
            while filename.with_name(id).exists() and filename.with_name(id)!=filename:
                pass

            show(f"开始处理{id}")

            if filename.is_file():
                if filename.suffix in ['.lrc','.mp3','.srt','.mp4','.txt','.wav']:
                    if not filename.with_name(id).exists():
                        filename.with_name(id).mkdir()
                    filename.rename(filename.with_name(id) / filename.name)
                else:
                    show(f"--处理完成，文件位于{filename}")
                    return
            else:
                mv_dir(filename,filename.with_name(id))
            filename= filename.with_name(id)
            trans_wav_or_flac_to_mp3(filename)
            change_lrc(filename)
            filename = archieve(filename)
            if work_mode.get() == 0:
                tags = spider(filename, id)
                if not tags:
                    show(f"--处理完成，文件位于{filename}")
                    continue
                change_tags(filename, tags, filename / (id + '.jpg'))
                filename = change_name(filename, tags, id)
                find_no_mp3(filename)
                icon(id, filename)
            mv_lrc(filename)
            show(f"处理完成，文件位于{filename}")
    except Exception as e:
        show(f"{e}")
        logging.error(e)
        logging.error("\n" + traceback.format_exc())

def find_no_mp3(filename):
    suffix = ['.mp3', '.mp4', '.avi']
    flag = 0
    for file in filename.rglob('*'):
        for suf in suffix:
            if file.suffix == suf:
                flag = 1
        if flag == 1:
            break
    if flag == 0:
        warning(filename, '缺少音声')


def warning(filename, e):
    with open('warning.txt', 'a+',encoding='utf-8') as f:
        show(f'warning:{filename}'+e)
        writer=csv.writer(f)
        writer.writerow([filename,e])


def mv_lrc(filename):
    for file in filename.rglob('*.lrc'):
        break
    else:
        return

    pattern=re.compile('\d+')
    changed=[]
    # 恢复原本文件名
    for mp3_file in filename.rglob('*.mp3'):
        if 'TIT2' in ID3FileType(mp3_file):
            mp3_TIT2 = ID3FileType(mp3_file)['TIT2'].__str__()
            mp3_file.replace(mp3_file.with_name(mp3_TIT2).with_suffix(mp3_file.suffix))
    # 匹配字幕
    for mp3_file in filename.rglob('*'):
        if mp3_file.suffix not in ['.mp3','.flac','.wav'] or mp3_file in changed:
            continue
        # 得出匹配分数
        distance = collections.OrderedDict()
        for lrc_file in filename.rglob("*.lrc"):
            if 'original_lrc' in lrc_file.__str__():
                continue
            ptn_lrc = []
            for i in pattern.findall(lrc_file.stem):
                ptn_lrc.append(''.join([str(int(j)) for j in list(i)]))
            ptn_lrc = ' '.join(ptn_lrc)

            # mp3_TIT2 = ID3FileType(mp3_file)['TIT2'].__str__()
            ptn_mp3=[]
            for i in pattern.findall(mp3_file.name):
                ptn_mp3.append(''.join([str(int(j)) for j in list(i)]))
            ptn_mp3=' '.join(ptn_mp3)
            try:
                # distance[Levenshtein.jaro(str(int(pattern.findall(lrc_file.stem)[0])),
                #                           str(int(num)))] = mp3_file
                distance[Levenshtein.jaro(ptn_mp3,
                                          ptn_lrc)] = lrc_file
            except:
                distance[Levenshtein.jaro(lrc_file.stem, mp3_file.stem)] = lrc_file
        if len(distance)==0:
            show(f"--{filename.name}:{Path(mp3_file).name}没有对应的LRC")
            continue
        # 寻找最匹配的lrc
        lrc_file_list=[]
        for k,v in distance.items():
            if v.stem==mp3_file.stem:
                lrc_file_list.append(v)
                break
        else:
            score=sorted(list(distance.items()))[-1][0]
            for k, v in distance.items():
                if k == score:
                    lrc_file_list.append(v)
                    break
        # 移动
        for lrc_file in lrc_file_list:
            if not (mp3_file.parent / lrc_file.name).exists():
                if lrc_file.exists():
                    lrc_file.replace(mp3_file.parent / lrc_file.name)
            if not mp3_file.with_stem(lrc_file.stem).exists():
                if mp3_file.exists():
                    mp3_file.replace(mp3_file.with_stem(lrc_file.stem))
                    changed.append(mp3_file.with_stem(lrc_file.stem))

    clear_empty_dir(filename)


def transform_lrc(input: Path, output: Optional[Path] = None, ops: str = 'add', file_type: str = 'lrc',
                  deleted: bool = False) -> None:
    if 'original_lrc' in input.parts:
        return
    # show(f"--处理lrc:{input}")
    # vtt转srt,lrc
    try:
        if input.suffix == '.vtt':
            if not input.with_stem(input.name.split('.')[0]).with_suffix('.srt').exists():
                convert_file = ConvertFile(input.__str__(), "utf-8")
                convert_file.convert()
                input.with_suffix('.srt').replace(input.with_stem(input.name.split('.')[0]).with_suffix('.srt'))
            input=input.with_stem(input.name.split('.')[0]).with_suffix('.srt')


        if input.suffix == '.srt':
            if not input.with_suffix('.lrc').exists():
                encodings=get_encoding(input)
                vtt=open_file(encodings,input)
                vtt=srt.parse(vtt)
                subs_output = pylrc.parse('')
                for item in list(vtt):
                    sec = srt.timedelta.total_seconds(item.start)
                    time = item.start.__str__().split(':')[1:]
                    time[0] = str(int(sec // 60))
                    time[1] = time[1][0:5]
                    lyc = pylrc.classes.LyricLine('[' + ':'.join(time) + ']', item.content)
                    subs_output.append(lyc)
                if not input.with_suffix('.lrc').exists():
                    with open(input.with_suffix('.lrc'), 'w', encoding='utf-8') as lrc_file:
                        lrc_file.write(subs_output.toLRC())
            input = input.with_suffix('.lrc')


        if not (input.parent / 'original_lrc').exists():
            Path.mkdir(input.parent / 'original_lrc')
        if not (input.parent / 'original_lrc' / input.name).exists():
            shutil.copy(input, input.parent / 'original_lrc' / input.name)
        if deleted:
            mv_to_trush(input.parent / 'original_lrc')
        if output is None:
            output = input
        encodings = get_encoding(input)
        lrc_string = open_file(encodings, input)
        subs_output = pylrc.parse('')
        lrc = ''
        for line in lrc_string.splitlines():
            if len(line) >= 7 and line[6] == ':':
                line = list(line)
                line[6] = '.'
                line = ''.join(line)
            lrc += line
            lrc += '\n'
        lrc_string = lrc
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
        with open(output, 'w', encoding='utf-8') as lrc_file:
            lrc_file.write(lrc_string)
    except:
        pass

def open_file(encodings, input):
    with open(input, 'rb') as f:
        lrc_file = f.read()
    for encoding in encodings:
        try:
            lrc_string = lrc_file.decode(encoding)
        except:
            pass
        else:
            break
    else:
        raise Exception("未知编码")
    return lrc_string


def get_encoding(input):
    with open(input, 'rb') as f:
        result = chardet.detect(f.read())
    if result['encoding'] == None:
        result['encoding'] = 'utf-8'
    if 'SIG' in result['encoding']:
        with open(input, 'rb') as f:
            lrc_string = f.read()[3:]
        with open(input, 'wb') as f:
            f.write(lrc_string)
        with open(input, 'rb') as f:
            result = chardet.detect(f.read())
    encodings = [result['encoding'], 'gbk', 'utf-8']
    return encodings


def clear_empty_dir(filename):
    for dir in filename.iterdir():
        try:
            dir.rmdir()
        except:
            pass

def icon(id, filename):
    if not icon_checked.get():
        return
    # mv_dir(filename,filename.with_suffix('.tmp'))
    iconPath = get_icon(id, filename)
    change_icon(iconPath)
    show('--已设置文件图标。')
    # mv_dir(filename.with_suffix('.tmp'), filename)

def get_icon(id, newname):
    with img.open(newname / (id + '.jpg')) as image:
        iconPath = newname / (id + '.ico')
        x, y = image.size
        size = max(x, y)
        new_im = img.new('RGBA', (size, size), (255, 255, 255, 0))
        new_im.paste(image, ((size - x) // 2, (size - y) // 2))
        new_im.save(iconPath)
    return iconPath

def unzip(filename):
    try:
        passwd = filename.stem.split()
        passwd.extend(filename.name.split('.')[0].split())
        passwd.extend(read_config('config.txt','保存的密码:'))
        if filename.is_file():
            filename = file_unzip(filename, passwd)
        else:
            for file in filename.rglob('*'):
                file_unzip(file, passwd)

    except:
        show(f'{filename}解压失败')
    return filename

def RJ_No(filename):
    if not filename or not filename.exists():
        show('--空文件，结束。')
        return None

    RJ = {}
    id = re.search("RJ\d{8}", filename.name.__str__())
    if not id:
        id = re.search("RJ\d{6}", filename.name.__str__())
    if id:
        id = (filename.name.__str__())[id.regs[0][0]:id.regs[0][1]]
        show(f'--找到{id}')
        RJ[filename] = id
        return RJ
    else:
        for file in filename.rglob("*"):
            id = re.search("RJ\d{8}", file.name.__str__())
            if not id:
                id = re.search("RJ\d{6}", file.name.__str__())
            if file.exists() and id:
                id = (file.name.__str__())[id.regs[0][0]:id.regs[0][1]]
                show(f'--找到{id}')
                file.replace(filename.parent / file.name)
                RJ[filename.parent / file.name] = id
    if RJ:
        mv_to_trush(filename)
    else:
        show(f'--未找到RJ编号，结束，文件位于{filename}。')
    return RJ


def get_other_name(filename,id, RJ_path):
    raletive= '*'+id + '*'
    othername=[]
    RJ_path.append(filename.parent)
    for path in RJ_path:
        for file in path.glob(raletive):
            if (file/'desktop.ini').exists() and file!=filename:
                othername.append(file)
    return othername


def show(info):
    # if info[0]=='-':
    #     return
    try:
        print(info)
        info_text.insert('end', info + "\n")
        info_text.see("end")
        info_text.update()
    except:
        pass

def show_wait(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        show_text = threading.Thread(target=wait_text, args=())
        show_text.daemon = True
        show_text.start()
        result=f(*args, **kwargs)
        global wait_flag
        wait_flag = False
        show_text.join()
        return result
    return decorated

def wait_text():
    try:
        global wait_flag
        wait_flag = True
        sleep(1)
        info_text.insert('end', ' ')
        info_text.update()
        while wait_flag:
            for i in range(6):
                if not wait:
                    break
                info_text.insert('end', '*')
                info_text.update()
                sleep(1)
            index = info_text.index('end')
            index = str(int(index.split('.')[0]) - 1) + '.1'
            info_text.delete(index, 'end')
            info_text.update()
            sleep(1)
        index = info_text.index('end')
        index = str(int(index.split('.')[0]) - 1) + '.0'
        info_text.delete(index, 'end')
        info_text.insert('end', '\n')
        info_text.update()
    except:
        pass


def change_name(filename, tags, id):
    group, title, cv = tags
    # 查找翻译好的中文标题
    t=None
    for _ in filename.rglob('*.chinese_title'):
        t=_
    if t!=None:
        title=t.stem.__str__()
        title = title.split(']')
        title = title[-1]
    elif '汉化组' in filename.__str__():
        name_list=filename.name.__str__().split('-')
        name_list.sort(key=lambda x: len(x))
        title= name_list[-1]
        Path.touch((filename / title).with_suffix('.chinese_title'))
        title=title.split(']')
        title=title[-1]
    # 翻译标题
    try:
        if translate_checked.get():
            trans = translate(title)
            if 'from' in trans and trans['from'] != 'zh':
                title = trans['trans_result'][0]['dst']
            for file in filename.rglob("*"):
                if not file.is_file() or file.suffix not in ['.lrc'] or 'original_lrc' in file.parent.__str__():
                    continue
                trans = translate(file.stem)
                if 'from' in trans and trans['from'] != 'zh':
                    name = trans['trans_result'][0]['dst']
                    file.replace(file.with_stem(name))
    except:
        pass

    # 构建文件名
    newname=id
    lcr = 0
    audio_list=['.mp3','.mp4','.wav','flac']
    audio=0
    for _ in Path(filename).rglob('*.lrc'):
        lcr = 1
        break
    for _ in Path(filename).rglob('*'):
        if _.suffix in audio_list:
            audio=1
            break
    if lcr == 1:
        newname = newname + ' ' + '[汉化]'
    if audio == 0:
        newname = newname + ' ' + '[缺少音频]'
    if group_checked.get():
        newname = newname + ' ' + '[' + group + ']'
    if title_checked.get():
        newname = newname + ' ' + title
    if cv_checked.get() and cv != '':
        newname = newname + ' ' + r'(CV ' + ' '.join(cv.split(';')) + ')'
    newname = re.sub('[\\\/:\*\?"<>\|]', '', newname)
    mv_dir(filename,filename.with_name(newname))

    show(f"--已重命名文件夹为{newname}")
    return filename.with_name(newname)

def archieve(filename):
    archive = archive_checked.get()
    replace=copy_checked.get()
    if archive:
        # 合并同RJ
        id = filename.name
        lrc = 0
        for _ in Path(filename).rglob('*.lrc'):
            lrc = 1
            break
        original_RJ_path = read_config('config.txt','RJ归档文件夹:')
        chinese_RJ_path = read_config('config.txt', 'RJ汉化归档文件夹:')

        if not original_RJ_path or not Path(original_RJ_path).is_dir():
            original_RJ_path = Path(filename.parent)
        else:
            original_RJ_path = Path(original_RJ_path)
        if not chinese_RJ_path or not Path(chinese_RJ_path).is_dir():
            chinese_RJ_path = Path(filename.parent)
        else:
            chinese_RJ_path = Path(chinese_RJ_path)

        others_RJ_path = []
        RJ_path = [original_RJ_path, chinese_RJ_path] + others_RJ_path
        othername_list = get_other_name(filename, id, RJ_path)

        newname=None
        for file in othername_list:
            if file.parent==chinese_RJ_path:
                newname = file
                break
        if not newname and lrc==1:
            newname= chinese_RJ_path/id
        if not newname:
            for file in othername_list:
                if file.parent == original_RJ_path:
                    newname = file
                    break
        if not newname:
            newname= original_RJ_path/id

        for othername in othername_list:
            if othername and othername.exists():
                mv_dir(othername, newname, replace)
        # 移动文件
        mv_dir(filename, newname, replace)
    return newname


def mv_dir(filename, newname, replace=True):
    # 移动
    if filename.exists() and newname != filename:
        for file in filename.rglob("*"):
            if file.exists() and file.is_file():
                new = newname / file.relative_to(filename)
                if not new.parent.exists():
                    new.parent.mkdir(parents=True)
                try:
                    if not new.exists():
                        if replace:
                            shutil.move(file,new)
                        else:
                            shutil.copy(file,new)
                except:
                    pass
        if replace and filename.exists():
            try:
                shutil.rmtree(filename)
            except:
                try:
                    mv_to_trush(filename)
                except:
                    pass
    return newname


# 修改MP3信息
def change_tags(filename, tags, picPath):
    if not mp3_checked.get():
        return
    show("--开始设置mp3信息")
    tag_pool = ThreadPoolExecutor(max_workers=10)
    group, title, cv = tags
    with open(picPath, 'rb') as f:
        picData = f.read()
    all_task=[]
    for file in filename.rglob("*.mp3"):
        info = {'picData': picData, 'title': file.stem,
                'artist': cv, 'album': title, 'albumartist': group}
        all_task.append(tag_pool.submit(set_Info, file, info))
        pass

    wait(all_task)
    show("--已设置mp3信息")


# 修改MP3的ID3标签
def set_Info(path, info):
    # show(f"--处理{path.stem}")
    songFile = ID3FileType(path)
    try:
        songFile.add_tags()
    except:
        pass
    if 'APIC' not in songFile.keys() and 'APIC:' not in songFile.keys():
    # if True:
        songFile['APIC'] = APIC(  # 插入封面
            encoding=3,
            mime='image/jpeg',
            type=3,
            desc=u'Cover',
            data=info['picData']
        )
    if 'TIT2' not in songFile.keys():
        songFile['TIT2'] = TIT2(  # 插入歌名
            encoding=3,
            text=info['title']
        )
    if 'TPE1' not in songFile.keys():
        songFile['TPE1'] = TPE1(  # 插入声优
            encoding=3,
            text=info['artist']
        )
    if 'TALB' not in songFile.keys():
        songFile['TALB'] = TALB(  # 插入专辑名
            encoding=3,
            text=info['album']
        )
    if 'TPE2' not in songFile.keys():
        songFile['TPE2'] = TPE2(  # 插入社团名
            encoding=3,
            text=info['albumartist']
        )
    songFile['TRCK'] = TRCK(  # track设为空
        encoding=3,
        text=''
    )
    songFile.save()
    del songFile


# 爬取信息与图片
def spider(filename, id):
    show('--开始爬取信息')
    url = 'https://www.dlsite.com/maniax/work/=/product_id/' + id + '.html/?locale=zh_CN'

    # req = urllib.request.Request(url=url, method="POST")

    try:
        response = urllib.request.urlopen(url)
    except:
        try:
            #使用代理尝试
            proxy = {
                'socks5': '127.0.0.1:7890'
            }
            handler = urllib.request.ProxyHandler(proxies=proxy)
            # 获取opener对象
            opener = urllib.request.build_opener(handler)
            response = opener.open(url)
        except:
            show('--网络故障,未找到网页，结束。')
            return None
    bs = BeautifulSoup(response.read().decode('utf-8'), 'html.parser')

    name = bs.select('#work_name')[0].text.strip()
    name = re.sub(r"【.*?】", "", name).strip()

    cv = ''
    for i in bs.select('#work_outline>tr'):
        if i.th.text == '声優' or i.th.text == '声优':
            cv = i.td.text.replace('/', ' ')
            cv = ';'.join(cv.split())
    group = bs.select('.maker_name>a')[0].text.strip()
    # imgurl = r'https:' + bs.select('.active>img')[0]['src']
    imgurl = r'https:' + bs.select('.active>picture>img')[0]['srcset']
    assert filename.exists()
    urllib.request.urlretrieve(imgurl, Path(filename) / (id + ".jpg"))
    show("--爬取完成")
    return group, name, cv


# 清理空文件夹，并找到RJ号开头的文件夹
def clear(filename):
    if filename.is_file():
        return filename
    no_process=['original_lrc']
    flag = True
    file_count=0
    for file in filename.rglob("*"):
        if file.is_file():
            file_count+=1
    if file_count==0:
        mv_to_trush(filename)
        return filename

    while (flag):
        flag = False
        for root, dirs, files in os.walk(filename.__str__()):
            if Path(root).name in no_process:
                continue
            fullname=True
            if len(dirs) == 1 and len(files) == 0:
                if fullname:
                    file = Path(root + ' ' + dirs[0])
                else:
                    file= Path(root).parent/dirs[0]
                if root == filename.__str__():
                    filename = Path(file)
                Path.rename(Path(root) / dirs[0],file)
                shutil.rmtree(root)
                flag = True
                break
            if len(dirs) == 0 and len(files) == 1:
                if fullname:
                    file = Path(root + ' ' + files[0])
                else:
                    file = Path(root).parent / files[0]
                if root == filename.__str__():
                    filename = Path(file)
                Path.rename(Path(root) / files[0],file)
                shutil.rmtree(root)
                flag = True
                break
            if len(dirs) == 0 and len(files) == 0:
                shutil.rmtree(root)
                flag = True
    newname = []
    for item in filename.stem.split(' '):
        if item not in newname:
            newname.append(item)
    newname= filename.with_stem(' '.join(newname))
    if filename !=newname:
        filename.replace(newname)
    return newname

@ show_wait
def cmd_call(cmd):
    return run(cmd, shell=True).returncode

# 更改文件夹图标
def change_icon(icon):
    root = icon.parent
    os.chmod(root, stat.S_IREAD)
    iniline1 = "[.ShellClassInfo]"
    iniline2 = "IconResource=" + icon.name + ",0"
    iniline = iniline1 + "\n" + iniline2 + '\n[ViewState]\nMode=\nVid=\nFolderType=Music\n'
    cmd1 = icon.drive
    cmd2 = 'cd "{}"'.format(icon.parent)
    cmd3 = "attrib -h -s desktop.ini"
    cmd = cmd1 + " && " + cmd2 + " && " + cmd3
    cmd_call(cmd)
    with open(root / "desktop.ini", "w+", encoding='utf-8') as inifile:
        inifile.write(iniline)
    cmd1 = icon.drive
    cmd2 = 'cd "{}"'.format(icon.parent)
    cmd3 = "attrib +h +s desktop.ini"
    cmd = cmd1 + " && " + cmd2 + " && " + cmd3
    cmd_call(cmd)
    pass


# wav文件转换为mp3
def trans_wav_or_flac_to_mp3(filesname):
    if not wav_to_mp3_checked.get():
        return
    show('--开始转换mp3')
    mp3, wav, flac = 0, 0, 0
    for filename in Path(filesname).rglob('*'):
        if filename.suffix == '.mp3':
            mp3 = 1
        if filename.suffix == '.wav':
            wav = 1
        if filename.suffix == '.flac':
            flac = 1

    thread_list=[]
    if wav == 1:
        for filename in Path(filesname).rglob('*.wav'):
            if filename.suffix == '.wav':
                thread_list.append(trans_pool.submit(to_mp3,filename))
    if flac == 1:
        for filename in Path(filesname).rglob('*.flac'):
            if filename.suffix == '.flac':
                thread_list.append(trans_pool.submit(to_mp3,filename))
    for thread in thread_list:
        thread.result()

    if mp3 == 1:
        if wav == 1:
            for filename in Path(filesname).rglob('*.wav'):
                if filename.exists() and filename.is_file():
                    for file in Path(filesname).rglob('*.mp3'):
                        if file.stem == filename.stem:
                            mv_to_trush(filename)
                            break
        if flac == 1:
            for filename in Path(filesname).rglob('*.flac'):
                if filename.exists() and filename.is_file():
                    for file in Path(filesname).rglob('*.mp3'):
                        if file.stem == filename.stem:
                            mv_to_trush(filename)
                            break

    show('--mp3转换完成')

@show_wait
def to_mp3(filename):
    show(f'--转换{filename.stem}')
    song = AudioSegment.from_file(filename)
    song.export(filename.with_suffix('.mp3'), format="mp3")
    mv_to_trush(filename)



def extract_mp3_from_video(filesname):
    try:
        if not extract_checked.get():
            return
        video= ['.mp4', '.ts','.mkv','.webm']
        for filename in [filesname]+(Path(filesname).rglob('*')):
            if filename.is_file() and filename.suffix in video:
                to_mp3(filename)
    except:
        show('--提取MP3失败')
# 解压缩

def file_unzip(filename, passwd):
    notzip = ['.lrc', '.ass', '.ini', '.url', '.apk', '.heic','.chinese_title','.srt','.vtt','mp3','.mp4','.ts','.mkv','.webm','.wav','.flac','.jpg','.png','txt']
    maybezip = ['.rar', '.zip', '.7z','.exe']
    if filename.exists() and filename.is_file():
        # 获取解压目录
        unzip = read_config('config.txt','解压目录:')
        if not unzip or not Path(unzip).is_dir():
            unzip_dir=filename.parent
        else:
            unzip_dir=Path(unzip)
        # 尝试解压
        if ((len(filename.suffixes) >1 and (mimetypes.guess_type(filename) == (None, None) or filename.suffix == '.exe')) or (filename.suffix in maybezip or mimetypes.guess_type(filename) == (None, None))) and filename.suffix not in notzip and filename.name not in ['.DS_Store']:

            result = 2
            # 解压目标文件夹名
            output_file = filename.with_name(filename.name.split('.')[0]) if filename.suffix != '' else filename.with_suffix(
                ".voiceWorkTemp")
            if unzip_dir.__str__() in filename.__str__():
                output = output_file
            else:
                output = unzip_dir / output_file.name
            # 尝试解压
            for pd in passwd:
                if output.exists():
                    shutil.rmtree(output, ignore_errors=True)
                cmd = f'bz.exe x -o:"{output}" -aoa -y -p:"{pd}" "{filename}"'
                result = result and cmd_call(cmd)
                global wait_flag
                wait_flag = False
                if not result:
                    break
                else:
                    if output.exists():
                        shutil.rmtree(output,ignore_errors=True)
            # 解压成功则
            if not result:
                show(f'--{filename.name}解压至{output}')
                # 删除分卷压缩
                for file in filename.parent.glob("*"):
                    if len(filename.suffixes) > 1:
                        if file.is_file and file.name.split('.')[0:-2] == filename.name.split('.')[0:-2] and len(
                                file.suffixes) >= 2:
                            mv_to_trush(file)
                    else:
                        mv_to_trush(filename)
                        # 删除zip分卷压缩
                        if filename.suffix == '.zip':
                            for file in filename.parent.glob("*"):
                                if file.is_file and file.stem == filename.stem and file.suffix[0:2] == '.z':
                                    mv_to_trush(file)
                # 递归解压
                filename = output
                if filename.exists():
                    for file in filename.rglob('*'):
                        file_unzip(file, passwd)
        else:
            return filename
        if result:
            show(f'--{filename.name}解压失败')
        return filename


def read_config(config,item):
    try:
        with open(config, encoding='utf-8') as f:
            cfg = f.readlines()
            for i, line in enumerate(cfg):
                if line.strip() == item:
                    if item!='保存的密码:':
                        return cfg[i + 1].strip()
                    else:
                        return [i.strip() for i in cfg[i + 1:]]
    except:
        return None


def mv_to_trush(filename):
    try:
        filename = filename.__str__()
        res = shell.SHFileOperation((0, shellcon.FO_DELETE, filename, None,
                                     shellcon.FOF_SILENT | shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION, None,
                                     None))  # 删除文件到回收站
        if (not res[1]) and Path(filename).exists():
            cmd_call('del ' + filename)
    except:
        pass


def checkbox_register(text='NULL', value=1,command=None,group=None):
    checked = tk.IntVar(value=value)
    button = tk.Checkbutton(text=text, variable=checked, command=command)
    button.pack()
    if group!=None:
        group.append((button,checked))
    return checked


def change_lrc(filename):
    ops = 'add' if 1 else 'delete'
    file_type = 'srt' if 0 else 'lrc'
    if filename.is_file():
        transform_lrc(filename, ops=ops, file_type=file_type)
    elif filename.is_dir():
        for file in Path(filename).rglob("*"):
            if file.suffix in ['.lrc','.srt','.vtt']:
                transform_lrc(file, ops=ops, file_type=file_type)


def info_register():
    info_text = tk.Text()
    scroll = tk.Scrollbar()
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    scroll.config(command=info_text.yview)
    info_text.config(yscrollcommand=scroll.set)
    info_text.pack()
    return info_text


def spider_switch():
    spider = 1 if work_mode.get() == 0 else 0
    for button,value in spider_group:
        button.config(state = 'normal' if spider else 'disabled')
    others = 1 if work_mode.get() != 1 else 0
    for button, value in others_group:
        button.config(state='normal' if others else 'disabled')

window = tk.Tk()

if __name__ == '__main__':
    # 窗口信息
    window.title('音声文件夹整理')
    window.geometry('400x600')
    window.update()

    # 提示信息
    lable = tk.Label(window, text='将文件或文件夹拖入窗口,自动执行所选定操作.', font=('宋体', 12), wraplength=window.winfo_width())
    lable.pack()
    # 工作模式
    global work_mode
    work_mode = tk.IntVar(value=0)
    unzip_checked = tk.Radiobutton(text='递归解压', variable=work_mode, command=spider_switch, value=1)
    unzip_checked.pack()
    onlyRJ_checked = tk.Radiobutton(text='本地模式', variable=work_mode, command=spider_switch, value=2)
    onlyRJ_checked.pack()
    spider_checked = tk.Radiobutton(text='在线模式', variable=work_mode, command=spider_switch, value=0)
    spider_checked.pack()

    # 变量
    spider_group = []
    others_group = []

    global wav_to_mp3_checked
    wav_to_mp3_checked = checkbox_register('wav,flac转换为mp3', group=others_group)
    global ops_checked
    ops_checked = checkbox_register('lrc空行间隔', group=others_group)
    global type_checked
    type_checked = checkbox_register('lrc转换为srt', 0, group=others_group)
    global extract_checked
    extract_checked = checkbox_register('video提取MP3', 0, group=others_group)
    global archive_checked
    archive_checked = checkbox_register('归档到指定文件夹', group=others_group)
    global copy_checked
    copy_checked = checkbox_register('删除归档前的文件', 1, group=others_group)
    global group_checked
    group_checked = checkbox_register('文件夹名包含社团',group = spider_group)
    global title_checked
    title_checked = checkbox_register('文件夹名包含标题', group=spider_group)
    global cv_checked
    cv_checked = checkbox_register('文件夹名包含CV', group=spider_group)
    global icon_checked
    icon_checked = checkbox_register('改变文件夹图标', group=spider_group)
    global mp3_checked
    mp3_checked = checkbox_register('修改MP3信息', group=spider_group)
    global translate_checked
    translate_checked = checkbox_register('机翻标题和lrc', group=spider_group)
    global info_text
    info_text = info_register()
    # 主逻辑
    global pool
    pool = ThreadPoolExecutor(max_workers=10)
    global trans_pool
    trans_pool = ThreadPoolExecutor(max_workers=1)
    windnd.hook_dropfiles(window, func=dragged_files, force_unicode='utf-8')
    tk.mainloop()

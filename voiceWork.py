import collections
import mimetypes
import os
import re
import stat
import tkinter as tk
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import Levenshtein
import windnd
from PIL import Image as img
from bs4 import BeautifulSoup
from mutagen.id3 import ID3FileType, APIC, TIT2, TPE1, TALB, TPE2, TRCK
from pydub import AudioSegment
from win32com.shell import shell, shellcon

# 拖拽时执行的函数
from transform_lrc import transform_lrc
from translate import translate


def dragged_files(files):
    for filename in files:
        pool.submit(process, filename)

# 处理函数
def process(filename):
    try:
        show(f"开始处理{filename}")
        filename = Path(filename)
        # find_no_mp3(filename)
        filename = unzip(filename)
        trans_wav_or_flac_to_mp3(filename)
        filename = clear(filename)
        RJ = RJ_No(filename)
        if not RJ:
            return
        for filename, id in RJ.items():
            show(f"开始处理{id}")
            tags = spider(filename, id)
            if not tags:
                return
            change_tags(filename, tags, filename / (id + '.jpg'))
            mv_lrc(filename)
            change_lrc(filename)
            filename = change_name(filename, tags, id)
            icon(id, filename)
            show(f"--处理完成，文件位于{filename}")
    except Exception as e:
        show(f"{e}")

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
        with open('no_mp3.txt', 'a+') as f:
            show(f'{filename}缺少音声\n')
            f.write(f'{filename}\n')

def mv_lrc(filename):
    pattern=re.compile('\d+')
    for lrc_file in filename.rglob("*.lrc"):
        distance = collections.OrderedDict()
        for mp3_file in lrc_file.parent.glob('*.mp3'):
            try:
                distance[Levenshtein.jaro(str(int(pattern.findall(lrc_file.stem)[0])),
                                          str(int(pattern.findall(mp3_file.stem)[0])))] = mp3_file
            except:
                distance[Levenshtein.jaro(lrc_file.stem, mp3_file.stem)] = mp3_file
        if len(distance)==0:
            for mp3_file in filename.rglob("*.mp3"):
                try:
                    distance[Levenshtein.jaro(str(int(pattern.findall(lrc_file.stem)[0])), str(int(pattern.findall(mp3_file.stem)[0])))]=mp3_file
                except:
                    distance[Levenshtein.jaro(lrc_file.stem, mp3_file.stem)]=mp3_file
        if len(distance)==0:
            show(f"--{filename.name}:{lrc_file}没有对应的MP3")
            continue
        mp3_file_list=[]
        for k,v in distance.items():
            if v.stem==lrc_file.stem:
                mp3_file_list.append(v)
                break
        else:
            score=sorted(list(distance.items()))[-1][0]
            for k, v in distance.items():
                if k == score:
                    mp3_file_list.append(v)
                    break
        for mp3_file in mp3_file_list:
            if lrc_file.stem != mp3_file.stem:
                if not (mp3_file.parent / lrc_file.name).exists():
                    if lrc_file.exists():
                        lrc_file.replace(mp3_file.parent / lrc_file.name)
                if not mp3_file.with_name(lrc_file.name).with_suffix(mp3_file.suffix).exists():
                    if mp3_file.exists():
                        mp3_file.replace(mp3_file.with_name(lrc_file.name).with_suffix(mp3_file.suffix))
    clear_empty_dir(filename)


def clear_empty_dir(filename):
    for dir in filename.iterdir():
        try:
            dir.rmdir()
        except:
            pass


def icon(id, newname):
    if not icon_checked.get():
        return
    iconPath = get_icon(id, newname)
    change_icon(iconPath)
    show('--已设置文件图标。')


def get_icon(id, newname):
    image = img.open(newname / (id + '.jpg'))
    iconPath = newname / (id + '.ico')
    x, y = image.size
    size = max(x, y)
    new_im = img.new('RGBA', (size, size), (255, 255, 255, 0))
    new_im.paste(image, ((size - x) // 2, (size - y) // 2))
    new_im.save(iconPath)
    return iconPath


def unzip(filename):
    passwd = filename.stem.split()[-2:]
    if Path('passwd.txt').exists():
        with open('passwd.txt', encoding='utf-8') as f:
            passwd.extend([i.strip() for i in f.readlines()])
    # passwd=passwd[::-1]
    pre_clear(filename)
    if filename.is_file():
        filename = file_unzip(filename, passwd)
    elif filename.is_dir():
        for file in filename.rglob('*'):
            file_unzip(file, passwd)
    return filename


def RJ_No(filename):
    if not filename.exists():
        show('--空文件，结束。')
        return None
    RJ = {}
    id = re.search("RJ\d{6}", filename.__str__())
    if id:
        id = (filename.__str__())[id.regs[0][0]:id.regs[0][1]]
        show(f'--找到{id}')
        RJ[filename] = id
        return RJ
    else:
        for file in filename.rglob("*"):
            id = re.search("RJ\d{6}", file.__str__())
            if file.exists() and id:
                id = (file.__str__())[id.regs[0][0]:id.regs[0][1]]
                show(f'--找到{id}')
                file.replace(filename.parent / file.name)
                RJ[filename.parent / file.name] = id
    if RJ:
        mv_to_trush(filename)
    else:
        show(f'--未找到RJ编号，结束，文件位于{filename}。')
    return RJ


def get_other_name(newname):
    othername = ''
    chinese = 1
    for i, n in enumerate(newname.__str__().split(' ')):
        if i == 1:
            if n != '[汉化]':
                chinese = 0
                othername = othername + ' [汉化] ' + n
            else:
                pass
        else:
            if i > 0:
                othername += ' '
            othername += n
    if not othername.exists():
        raletive= '*'+newname.name.__str__().split(' ')[0] + '*'
        for file in newname.parent.glob(raletive):
            if (file/'desktop.ini').exists():
                othername=file
    return Path(othername), chinese


def show(info):
    try:
        info_text.insert('end', info + "\n")
        info_text.see("end")
        info_text.update()
    except:
        pass


def pre_clear(filename):
    for file in filename.rglob('*baiduyun*'):
        mv_to_trush(file)


def change_name(filename, tags, id):
    group, title, cv = tags
    t=None
    for _ in filename.rglob('*.chinese_title'):
        t=_
    if t!=None:
        title=t.stem.__str__()
        title = title.split(']')
        title = title[-1]
    elif '汉化组' in filename.__str__():
        list=filename.name.__str__().split('-')
        list.sort(key=lambda x: len(x))
        title=list[-1]
        Path.touch((filename / title).with_suffix('.chinese_title'))
        title=title.split(']')
        title=title[-1]
    if translate_checked.get():
        trans = translate(title)
        if trans['from'] != 'zh':
            title = trans['trans_result'][0]['dst']

    newname=id
    lcr = 0
    for _ in Path(filename).rglob('*.lrc'):
        lcr = 1
        break
    if lcr == 1:
        newname = newname + ' ' + '[汉化]'
    if group_checked.get():
        newname = newname + ' ' + '[' + group + ']'
    if title_checked.get():
        newname = newname + ' ' + title
    if cv_checked.get() and cv != '':
        newname = newname + ' ' + r'(CV ' + ' '.join(cv.split(';')) + ')'
    newname = re.sub('[\\\/:\*\?"<>\|]', '', newname)
    newname = filename.parent / newname
    othername, chinese = get_other_name(newname)
    if othername.exists():
        if chinese == 1:
            mv_dir(othername, newname)
        else:
            newname = othername
    mv_dir(filename, newname)
    show(f"--已重命名文件夹为{newname.name}")
    return newname


def mv_dir(filename, newname):
    if not filename.exists():
        return
    if newname != filename:
        for file in filename.rglob("*"):
            if file.is_file() and file.exists:
                new = newname / file.relative_to(filename)
                if not new.parent.exists():
                    new.parent.mkdir(parents=True)
                try:
                    new.parent.touch()
                    file.replace(new)
                except:
                    pass
        if filename.exists():
            mv_to_trush(filename)


# 修改MP3信息
def change_tags(filename, tags, picPath):
    if not mp3_checked.get():
        return
    show("--开始设置mp3信息")
    group, title, cv = tags
    with open(picPath, 'rb') as f:
        picData = f.read()
    for file in filename.rglob("*.mp3"):
        info = {'picData': picData, 'title': file.stem,
                'artist': cv, 'album': title, 'albumartist': group}
        set_Info(file, info)
    show("--已设置mp3信息")


# 修改MP3的ID3标签
def set_Info(path, info):
    show(f"--处理{path.stem}")
    songFile = ID3FileType(path)
    try:
        songFile.add_tags()
    except:
        pass
    if 'APIC' not in songFile.keys() and 'APIC:' not in songFile.keys():
        songFile['APIC'] = APIC(  # 插入封面
            encoding=3,
            mime='image/jpeg',
            type=3,
            desc=u'Cover',
            data=info['picData']
        )
    songFile['TIT2'] = TIT2(  # 插入歌名
        encoding=3,
        text=info['title']
    )
    songFile['TPE1'] = TPE1(  # 插入声优
        encoding=3,
        text=info['artist']
    )
    songFile['TALB'] = TALB(  # 插入专辑名
        encoding=3,
        text=info['album']
    )
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
    url = 'https://www.dlsite.com/maniax/work/=/product_id/' + id + '.html'
    headers = {
        'authority': 'www.dlsite.com',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-user': '?1',
        'sec-fetch-dest': 'document',
        'accept-language': 'zh-CN,zh;q=0.9',
        'cookie': 'dlsite_dozen=7; uniqid=dh5wpo6du5; _ga=GA1.2.1094789266.1575957606; _d3date=2019-12-10T06:00:05.977Z; _d3landing=1; _gaid=1094789266.1575957606; adultchecked=1; adr_id=Ub2DpFVcRxLOTwlgsJ1aGEchUBznAqLnj8zWGwt6giSQaWyc; _ebtd=1.k2fnwjp5b.1575971225; locale=zh-cn; Qs_lvt_328467=1577849294^%^2C1577892629^%^2C1577932057^%^2C1591933560; Qs_pv_328467=1052826062922660000^%^2C480504100285960260^%^2C4442059557915936000^%^2C3571252907537905700^%^2C993384634696700900; _ts_yjad=1598842520972; utm_c=blogparts; _inflow_params=^%^7B^%^22referrer_uri^%^22^%^3A^%^22level-plus.net^%^22^%^7D; _inflow_ad_params=^%^7B^%^22ad_name^%^22^%^3A^%^22referral^%^22^%^7D; _im_vid=01ES81S84JB3TTCY5CZRD5EKS5; _gcl_au=1.1.1542253388.1608169517; _gid=GA1.2.1086092439.1612663765; __DLsite_SID=f4ik5q573dn9trjf0rugu9d043; __juicer_sesid_9i3nsdfP_=5c159006-e3e3-4f60-b5fe-dc05c244f1c6; DL_PRODUCT_LOG=^%^2CRJ306930^%^2CRJ300000^%^2CRJ298978^%^2CRJ315336^%^2CRJ315852^%^2CRJ307073^%^2CRJ306798^%^2CRJ309328^%^2CRJ303189^%^2CRJ316357^%^2CRJ234791^%^2CRJ312136^%^2CRJ131395^%^2CRJ282673^%^2CRJ264706^%^2CRJ242260^%^2CRJ250966^%^2CRJ313604^%^2CRJ313754^%^2CRJ295229^%^2CRJ300532^%^2CRJ262976^%^2CRJ311359^%^2CRJ310955^%^2CRJ268194^%^2CRJ289705^%^2CRJ260052^%^2CRJ315474^%^2CRJ316119^%^2CRJ315405^%^2CRJ312692^%^2CRJ167776^%^2CRJ314102^%^2CRJ303183^%^2CRJ309544^%^2CRJ211905^%^2CRJ133234^%^2CRJ307037^%^2CRJ302768^%^2CRJ305343^%^2CRJ299936^%^2CRJ282627^%^2CRJ304923^%^2520^%^2CRJ272689^%^2CRJ303021^%^2CR305282^%^2CRJ297002^%^2CRJ307645^%^2CRJ291292^%^2CRJ295048; _inflow_dlsite_params=^%^7B^%^22dlsite_referrer_url^%^22^%^3A^%^22https^%^3A^%^2F^%^2Fwww.dlsite.com^%^2Fmaniax^%^2Fwork^%^2F^%^3D^%^2Fproduct_id^%^2FRJ306798.html^%^22^%^7D; _dctagfq=1356:1613404799.0.0^|1380:1613404799.0.0^|1404:1613404799.0.0^|1428:1613404799.0.0^|1529:1613404799.0.0; __juicer_session_referrer_9i3nsdfP_=5c159006-e3e3-4f60-b5fe-dc05c244f1c6___; _td=287255fd-bbc9-470a-b97d-8c0b1c6b9cd9; _gat=1',
    }
    req = urllib.request.Request(url=url, headers=headers, method="POST")
    try:
        response = urllib.request.urlopen(req)
    except:
        show('--网络故障,未找到网页，结束。')
        return None
    bs = BeautifulSoup(response.read().decode('utf-8'), 'html.parser')
    cv = ''
    name = bs.select('#work_name')[0].text.strip()
    name = re.sub(r"【.*?】", "", name)
    name=name.strip()

    for i in bs.select('#work_outline>tr'):
        if i.th.text == '声優' or i.th.text == '声优':
            cv = i.td.text.replace('/', ' ')
            cv = ';'.join(cv.split())
    group = bs.select('.maker_name>a')[0].text.strip()
    # imgurl = r'https:' + bs.select('.active>img')[0]['src']
    imgurl = r'https:' + bs.select('.active>picture>img')[0]['srcset']
    urllib.request.urlretrieve(imgurl, Path(filename) / (id + ".jpg"))
    show("--爬取完成")
    return group, name, cv


# 清理空文件夹，并找到RJ号开头的文件夹
def clear(filename):
    no_process=['original_lrc']
    flag = True
    while (flag):
        flag = False
        for root, dirs, files in os.walk(filename.__str__()):
            if Path(root).name in no_process:
                continue
            if len(dirs) == 1 and len(files) == 0:
                Path.rename(Path(root)/dirs[0], Path(root).parent / (dirs[0] + '.voiceWorkTemp'))
                mv_to_trush(root)
                if dirs[0] not in Path(root).stem.split('-'):
                    file = root + '-' + dirs[0]
                else:
                    file = root
                if root == filename.__str__():
                    filename = Path(file)
                Path.rename(Path(root).parent/(dirs[0] + '.voiceWorkTemp'),Path(file))
                flag = True
                break
            if len(dirs) == 0 and len(files) == 1:
                Path.rename(Path(root)/files[0],
                            Path(root).parent/(files[0] + '.voiceWorkTemp'))
                mv_to_trush(root)
                if files[0].split('.')[0] not in Path(root).stem.split('-'):
                    file = root + '-' + files[0]
                else:
                    file = root + Path(files[0]).suffix
                if root == filename.__str__():
                    filename = Path(file)
                Path.rename(Path(root).parent / (files[0] + '.voiceWorkTemp'),Path(file))
                flag = True
                break
            if len(dirs) == 0 and len(files) == 0:
                mv_to_trush(root)
                flag = True
    return filename


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
    os.system(cmd)
    with open(root / "desktop.ini", "w+", encoding='utf-8') as inifile:
        inifile.write(iniline)
        inifile.close()
    cmd1 = icon.drive
    cmd2 = 'cd "{}"'.format(icon.parent)
    cmd3 = "attrib +h +s desktop.ini"
    cmd = cmd1 + " && " + cmd2 + " && " + cmd3
    os.system(cmd)
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

    if wav == 1:
        for filename in Path(filesname).rglob('*.wav'):
            show(f'--转换{filename.stem}')
            song = AudioSegment.from_file(filename)
            song.export(filename.with_suffix('.mp3'), format="mp3")
            mv_to_trush(filename)
    if flac == 1:
        for filename in Path(filesname).rglob('*.flac'):
            show(f'--转换{filename.stem}')
            song = AudioSegment.from_file(filename)
            song.export(filename.with_suffix('.mp3'), format="mp3")
            mv_to_trush(filename)
    show('--mp3转换完成')


# 解压缩
def file_unzip(filename, passwd):
    notzip = ['.lrc', '.ass', '.ini', '.url', '.apk', '.heic','.chinese_title']
    maybezip = ['.rar', '.zip', '.7z']
    if filename.exists() and filename.is_file():
        if len(filename.suffixes) >= 2 and (mimetypes.guess_type(filename) == (
                None, None) or filename.suffix == '.exe') and filename.suffix not in notzip:
            result = 2
            for pd in passwd:
                output = filename.with_suffix("")
                cmd = f'bz.exe x -o:"{output}" -aoa -y -p:"{pd}" "{filename}"'
                result = result and os.system(cmd)
                if not result:
                    break
                else:
                    if output.exists():
                        mv_to_trush(output)
            if not result:
                for file in filename.parent.glob("*"):
                    if file.is_file and file.name.split('.')[0:-2] == filename.name.split('.')[0:-2] and len(
                            file.suffixes) >= 2:
                        mv_to_trush(file)
                filename = output
                for file in filename.rglob('*'):
                    file_unzip(file, passwd)

        elif filename.suffix in maybezip or (
                mimetypes.guess_type(filename) == (None, None) and filename.suffix not in notzip):
            result = 2
            for pd in passwd:
                output = filename.with_suffix('') if filename.suffix != '' else filename.with_suffix(".voiceWorkTemp")
                output.mkdir()
                cmd = f'bz.exe x -o:"{output}" -aoa -y -p:"{pd}" "{filename}"'
                result = result and os.system(cmd)
                if not result:
                    break
                else:
                    if output.exists():
                        mv_to_trush(output)
            if not result:
                mv_to_trush(filename)
                filename = filename.with_suffix('')
                output.replace(filename)
                for file in filename.rglob('*'):
                    file_unzip(file, passwd)
        else:
            return filename
        if result:
            show(f'--{filename.name}解压失败')
        return filename



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


def change_lrc(filename):
    ops = 'add' if 1 else 'delete'
    file_type = 'srt' if 0 else 'lrc'
    if filename.is_file():
        transform_lrc(filename, ops=ops, file_type=file_type)
    elif filename.is_dir():
        for file in Path(filename).rglob("*.lrc"):
            transform_lrc(file, ops=ops, file_type=file_type)




def info_register():
    info_text = tk.Text()
    scroll = tk.Scrollbar()
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    scroll.config(command=info_text.yview)
    info_text.config(yscrollcommand=scroll.set)
    info_text.pack()
    return info_text


window = tk.Tk()

if __name__ == '__main__':
    # 窗口信息
    window.title('音声文件夹整理')
    window.geometry('400x600')
    window.update()

    # 提示信息
    lable = tk.Label(window, text='将名字带有RJ号的文件夹拖入窗口,可一次拖入多个待处理的文件夹.', font=('宋体', 12), wraplength=window.winfo_width())
    lable.pack()
    lable = tk.Label(window, text='该操作会删除所有空文件夹,当A文件夹只包含一个子文件夹B时，会将B中的所有文件放到A内，并删除B。'
                                  '如果文件名没有RJ号，会找到并处理文件夹中第一个包含的RJ号的子文件夹，并删除其他所有文件，请谨慎使用。', font=('宋体', 12), fg='red',
                     wraplength=window.winfo_width())
    lable.pack()
    # 变量
    global wav_to_mp3_checked
    wav_to_mp3_checked = checkbox_register('wav转换为mp3')
    global group_checked
    group_checked = checkbox_register('文件夹名包含社团')
    global title_checked
    title_checked = checkbox_register('文件夹名包含标题')
    global cv_checked
    cv_checked = checkbox_register('文件夹名包含CV')
    global icon_checked
    icon_checked = checkbox_register('改变文件夹图标')
    global mp3_checked
    mp3_checked = checkbox_register('修改MP3信息')
    global translate_checked
    translate_checked = checkbox_register('翻译标题')
    global ops_checked
    ops_checked = checkbox_register('lrc空行间隔')
    global type_checked
    type_checked = checkbox_register('lrc转换为srt',0)
    global info_text
    info_text = info_register()
    # 主逻辑
    global pool
    pool = ThreadPoolExecutor(max_workers=5)
    windnd.hook_dropfiles(window, func=dragged_files, force_unicode='utf-8')
    tk.mainloop()

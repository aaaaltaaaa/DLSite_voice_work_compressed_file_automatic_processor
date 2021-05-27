import tkinter as tk
import os
import shutil
import windnd
from bs4 import BeautifulSoup
import urllib.request
import stat
import re
from PIL import Image as img
from mutagen.mp3 import EasyMP3 as MP3
from mutagen.id3 import APIC
from mutagen.wave import WAVE
from mutagen.id3 import ID3FileType, APIC, TIT2, TPE1, TALB, TPE2, TRCK
import threading
from pydub import AudioSegment
from pathlib import Path


# 拖拽时执行的函数
def draggedFiles(files):
    global filesname
    global flag
    filesname = files
    flag = True


# 处理函数
def process():
    global flag
    global filesname
    while (True):
        if flag == True:
            for filename in filesname:
                try:
                    global wavToMp3Checked
                    global groupChecked
                    global titleChecked
                    global cvChecked
                    global iconChecked
                    global mp3Checked
                    global infoText
                    filename=Path(filename)
                    if filename.is_file():
                        infoText.insert('end', "解压{}\n".format(filename))
                        infoText.see("end")
                        infoText.update()
                        passwd = filename.stem.split()[-1]
                        unzip(filename, passwd,'unziptemp')
                        filename = filename.parent / Path('unziptemp')
                    if wavToMp3Checked.get() == '1':
                        infoText.insert('end', "转换为MP3\n")
                        infoText.see("end")
                        infoText.update()
                        trans_wav_to_mp3(filename)

                    filename= filename.__str__()
                    filename = clear(filename)
                    id = re.search("RJ\d{6}", filename)
                    id = filename[id.regs[0][0]:id.regs[0][1]]
                    if not id:
                        infoText.insert('end', '未找到RJ编号\n')
                        infoText.see("end")
                        infoText.update()
                        continue
                    infoText.insert('end', '正在处理' + id + '\n')
                    infoText.see("end")
                    infoText.update()
                    info = spider(filename, id)
                    if not info:
                        continue
                    group, title, cv = info

                    newname = id
                    if groupChecked.get() == '1':
                        newname = newname + ' ' + '[' + group + ']'
                    if titleChecked.get() == '1':
                        newname = newname + ' ' + title
                    if cvChecked.get() == '1' and cv != '':
                        newname = newname + ' ' + r'(CV ' + cv + ')'
                    newname = re.sub('[\\\/:\*\?"<>\|]', '', newname)
                    newname = os.path.join(os.path.dirname(filename), newname)
                    shutil.move(filename, newname)
                    infoText.insert('end', "已重命名文件夹为" + os.path.basename(newname) + "\n")
                    infoText.see("end")
                    infoText.update()
                    if iconChecked.get() == '1':
                        image = img.open(os.path.join(newname, id + '.jpg'))
                        iconPath = os.path.join(newname, id + '.ico')
                        x, y = image.size
                        size = max(x, y)
                        new_im = img.new('RGBA', (size, size), (255, 255, 255, 0))
                        new_im.paste(image, ((size - x) // 2, (size - y) // 2))
                        new_im.save(iconPath)
                        changeIcon(iconPath)
                        infoText.insert('end', '已设置文件夹图标\n')
                        infoText.see("end")
                        infoText.update()
                    if mp3Checked.get() == "1":
                        changeTags(newname, group, title, cv, os.path.join(newname, id + '.jpg'))
                        infoText.insert('end', "已设置mp3信息\n")
                        infoText.see("end")
                        infoText.update()
                except:
                    infoText.insert('end', "处理失败，请重试\n")
                    infoText.see("end")
                    infoText.update()
        flag = False


# 修改MP3信息
def changeTags(filename, group, title, cv, picPath):
    with open(picPath, 'rb') as f:
        picData = f.read()
    for root, dirs, files in os.walk(filename):
        for file in files:
            if os.path.splitext(file)[1] == ".mp3":
                info = {'picData': picData, 'title': os.path.splitext(file)[0],
                        'artist': cv, 'album': title, 'albumartist': group}
                SetInfo(os.path.join(root, file), info)


# 修改MP3的ID3标签
def SetInfo(path, info):
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


# 爬取信息与图片
def spider(filename, id):
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
        infoText.insert('end', "网络故障,未找到相关信息\n")
        infoText.update()
        return None
    bs = BeautifulSoup(response.read().decode('utf-8'), 'html.parser')
    cv = ''
    name = bs.select('#work_name>a')[0].text
    name = re.sub(r"【.*?】", "", name)

    for i in bs.select('#work_outline>tr'):
        if i.th.text == '声優' or i.th.text == '声优':
            cv = i.td.text.replace('/', ' ')
            cv = ' '.join(cv.split())
    group = bs.select('.maker_name>a')[0].text
    imgurl = r'https:' + bs.select('.active>img')[0]['src']
    urllib.request.urlretrieve(imgurl, os.path.join(filename, id + '.jpg'))
    return group, name, cv


# 清理空文件夹，并找到RJ号开头的文件夹
def clear(filename):
    for root, dirs, files in os.walk(filename):
        id = re.search("RJ\d{6}", root)
        if id != None:
            basename = os.path.basename(root)
            os.rename(filename, os.path.join(os.path.dirname(filename), basename))
            try:
                shutil.rmtree(filename)
            except:
                pass
            filename = os.path.join(os.path.dirname(filename), basename)
            break

    flag = True
    while (flag):
        flag = False
        for root, dirs, files in os.walk(filename):
            if len(dirs) == 1 and len(files) == 0:
                os.rename(os.path.join(root, dirs[0]), os.path.join(os.path.dirname(root), 'voiceWorkTemp'))
                os.removedirs(root)
                os.rename(os.path.join(os.path.dirname(root), 'voiceWorkTemp'), root)
                flag = True
                break
            if len(dirs)==0 and len(files) ==0:
               shutil.rmtree(root)
               flag = True
    return filename


# 更改文件夹图标
def changeIcon(icon):
    root = os.path.dirname(icon)
    os.chmod(root, stat.S_IREAD)
    iniline1 = "[.ShellClassInfo]"
    iniline2 = "IconResource=" + os.path.basename(icon) + ",0"
    iniline = iniline1 + "\n" + iniline2 + '\n[ViewState]\nMode=\nVid=\nFolderType=Music\n'
    cmd1 = icon[0:2]
    cmd2 = "cd " + '\"' + os.path.dirname(icon) + '\"'
    cmd3 = "attrib -h -s " + 'desktop.ini'
    cmd = cmd1 + " && " + cmd2 + " && " + cmd3
    os.system(cmd)
    with open(root + "\\" + "desktop.ini", "w+", encoding='utf-8') as inifile:
        inifile.write(iniline)
        inifile.close()
    cmd1 = icon[0:2]
    cmd2 = "cd " + '\"' + os.path.dirname(icon) + '\"'
    cmd3 = "attrib +h +s " + 'desktop.ini'
    cmd = cmd1 + " && " + cmd2 + " && " + cmd3
    os.system(cmd)


# wav文件转换为mp3
def trans_wav_to_mp3(filesname):
    mp3, wav = 0,0
    for filename in Path(filesname).rglob('*'):
        if filename.suffix == '.mp3':
            mp3 = 1
        if filename.suffix == '.wav':
            wav = 1
    if mp3 == 0 and wav == 1:
        for filename in Path(filesname).rglob('*.wav'):
            song = AudioSegment.from_wav(filename)
            song.export(filename.with_suffix('.mp3'), format="mp3")
            os.remove(filename)
    elif mp3==1 and wav ==1:
        for filename in Path(filesname).rglob('*.wav'):
            if filename.exists():
                if filename.is_file():
                    for file in Path(filesname).rglob('*.mp3'):
                        if file.stem == filename.stem:
                            os.remove(filename)


# 解压缩
def unzip(filesname, passwd, unzipdir):
    if filesname.is_file():
        if filesname.suffix == '' or filesname.suffix == '.rar' or filesname.suffix == '.zip' or filesname.suffix == '.7z' or filesname.suffix == '.part1':
            unzipdir = filesname.parent / Path(unzipdir)
            cmd = 'Bandizip.exe x -aoa -target:auto -p:' + passwd + ' -o:' + '"{}"'.format(unzipdir) + ' ' + '"{}"'.format(filesname)
            os.system(cmd)
            if filesname.exists():
                if filesname.is_file():
                    os.remove(filesname)
                else:
                    shutil.rmtree(filesname)

    newname = unzipdir / filesname.stem
    for filename in newname.rglob('*'):
        unzip(filename,passwd, unzipdir)


if __name__ == '__main__':
    window = tk.Tk()
    window.title('音声文件夹整理')
    window.geometry('400x400')
    global wavToMp3Checked
    global groupChecked
    global titleChecked
    global cvChecked
    global iconChecked
    global mp3Checked
    global infoText

    wavToMp3Checked = tk.StringVar()
    groupChecked = tk.StringVar()
    titleChecked = tk.StringVar()
    cvChecked = tk.StringVar()
    iconChecked = tk.StringVar()
    mp3Checked = tk.StringVar()

    wavToMp3Checked.set(1)
    groupChecked.set(1)
    titleChecked.set(1)
    cvChecked.set(1)
    iconChecked.set(1)
    mp3Checked.set(1)

    window.update()
    lable1 = tk.Label(window, text='将名字带有RJ号的文件夹拖入窗口,可一次拖入多个待处理的文件夹.', font=('宋体', 12), wraplength=window.winfo_width())
    lable1.pack()
    lable1 = tk.Label(window, text='该操作会删除所有空文件夹,当A文件夹只包含一个子文件夹B时，会将B中的所有文件放到A内，并删除B。'
                                   '如果文件名没有RJ号，会找到并处理文件夹中第一个包含的RJ号的子文件夹，并删除其他所有文件，请谨慎使用。', font=('宋体', 12), fg='red',
                      wraplength=window.winfo_width())
    lable1.pack()
    wavToMp3Radio = tk.Checkbutton(text='wav转换为mp3', variable=wavToMp3Checked)
    wavToMp3Radio.pack()
    groupRadio = tk.Checkbutton(text='文件夹名包含社团', variable=groupChecked)
    groupRadio.pack()
    nameRadio = tk.Checkbutton(text='文件夹名包含标题', variable=titleChecked)
    nameRadio.pack()
    cvRadio = tk.Checkbutton(text='文件夹名包含CV', variable=cvChecked)
    cvRadio.pack()
    iconRadio = tk.Checkbutton(text='更改文件夹图标', variable=iconChecked)
    iconRadio.pack()
    mp3Radio = tk.Checkbutton(text='更改mp3信息', variable=mp3Checked)
    mp3Radio.pack()
    infoText = tk.Text()
    scroll = tk.Scrollbar()
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    scroll.config(command=infoText.yview)
    infoText.config(yscrollcommand=scroll.set)
    infoText.pack()

    global filesname
    global flag
    flag = False
    filesname = ''
    thread = threading.Thread(target=process, daemon=True)
    thread.start()
    windnd.hook_dropfiles(window, func=draggedFiles, force_unicode='utf-8')
    tk.mainloop()

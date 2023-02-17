import collections
import csv
import logging
import mimetypes
import os
import re
import shutil
import stat
import tkinter as tk
import traceback
import urllib.request
from concurrent.futures import ThreadPoolExecutor, wait
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
        filename = unzip(filename)
        trans_wav_or_flac_to_mp3(filename)
        extract_mp3_from_video(filename)
        filename = clear(filename)
        RJ = RJ_No(filename)
        if not RJ or filename.is_file():
            show(f"--处理完成，文件位于{filename}")
            return
        for filename, id in RJ.items():
            show(f"开始处理{id}")
            mv_dir(filename,filename.with_name(id))
            filename= filename.with_name(id)
            tags = spider(filename, id)
            if not tags:
                return
            change_tags(filename, tags, filename / (id + '.jpg'))
            filename = change_name(filename, tags, id)
            mv_lrc(filename)
            change_lrc(filename)
            find_no_mp3(filename)
            icon(id, filename)
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
        error(filename,'缺少音声')


def error(filename,e):
    with open('error.txt', 'a+',encoding='utf-8') as f:
        show(f'error:{filename}'+e)
        writer=csv.writer(f)
        writer.writerow([filename,e])


def mv_lrc(filename):
    pattern=re.compile('\d+')
    changed=[]
    for mp3_file in filename.rglob('*.mp3'):
        mp3_TIT2 = ID3FileType(mp3_file)['TIT2'].__str__()
        mp3_file.replace(mp3_file.with_name(mp3_TIT2).with_suffix(mp3_file.suffix))
    for lrc_file in filename.rglob("*.lrc"):
        if 'original_lrc' in lrc_file.__str__():
            continue
        distance = collections.OrderedDict()
        ptn_lrc = []
        for i in pattern.findall(lrc_file.stem):
            ptn_lrc.append(''.join([str(int(j)) for j in list(i)]))
        ptn_lrc = ' '.join(ptn_lrc)
        for mp3_file in lrc_file.parent.rglob('*.mp3'):
            if mp3_file not in changed:
                mp3_TIT2 = ID3FileType(mp3_file)['TIT2'].__str__()
                ptn_mp3=[]
                for i in pattern.findall(mp3_TIT2):
                    ptn_mp3.append(''.join([str(int(j)) for j in list(i)]))
                ptn_mp3=' '.join(ptn_mp3)
                try:
                    # distance[Levenshtein.jaro(str(int(pattern.findall(lrc_file.stem)[0])),
                    #                           str(int(num)))] = mp3_file
                    distance[Levenshtein.jaro(ptn_mp3,
                                              ptn_lrc)] = mp3_file
                except:
                    distance[Levenshtein.jaro(lrc_file.stem, mp3_file.stem)] = mp3_file
        if len(distance)==0:
            for mp3_file in filename.rglob("*.mp3"):
                if mp3_file not in changed:
                    ptn_mp3 = ' '.join(pattern.findall(mp3_file.stem))
                    try:
                        # distance[Levenshtein.jaro(str(int(pattern.findall(lrc_file.stem)[0])), str(int(pattern.findall(mp3_file.stem)[0])))]=mp3_file
                        distance[Levenshtein.jaro(ptn_mp3,
                                                  ptn_lrc)] = mp3_file
                    except:
                        distance[Levenshtein.jaro(lrc_file.stem, mp3_file.stem)]=mp3_file
        if len(distance)==0:
            # show(f"--{filename.name}:{Path(lrc_file).name}没有对应的MP3")
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
            if not (mp3_file.parent / lrc_file.name).exists():
                if lrc_file.exists():
                    lrc_file.replace(mp3_file.parent / lrc_file.name)
            if not mp3_file.with_name(lrc_file.name).with_suffix(mp3_file.suffix).exists():
                if mp3_file.exists():
                    newname= mp3_file.with_name(lrc_file.name).with_suffix(mp3_file.suffix)
                    mp3_file.replace(newname)
                    changed.append(newname)

    clear_empty_dir(filename)


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
    passwd = filename.stem.split()
    if Path('passwd.txt').exists():
        with open('passwd.txt', encoding='utf-8') as f:
            passwd.extend([i.strip() for i in f.readlines()])
    pre_clear(filename)
    if filename.is_file():
        filename = file_unzip(filename, passwd)
    elif filename.is_dir():
        for file in filename.rglob('*'):
            file_unzip(file, passwd)
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


def pre_clear(filename):
    for file in filename.rglob('*baiduyun*'):
        mv_to_trush(file)


def change_name(filename, tags, id):
    # 合并相同RJ
    original_RJ_path=Path(r'E:\音声\RJ')
    chinese_RJ_path=Path(r'E:\音声\RJ汉化')
    move_to_RJ= archive_checked.get()
    if move_to_RJ:
        RJ_path=[original_RJ_path, chinese_RJ_path ]
    else:
        RJ_path=[]
    othername_list = get_other_name(filename,id, RJ_path)
    for othername in othername_list:
        if othername and othername.exists():
            filename=mv_dir(othername, filename)

    group, title, cv = tags
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
    if translate_checked.get():
        trans = translate(title)
        if 'from' in trans and trans['from'] != 'zh':
            title = trans['trans_result'][0]['dst']

    newname=id
    lcr = 0
    mp3 = 0
    for _ in Path(filename).rglob('*.lrc'):
        lcr = 1
        break
    for _ in Path(filename).rglob('*.mp3'):
        mp3 = 1
        break
    if lcr == 1:
        newname = newname + ' ' + '[汉化]'
    if mp3 == 0:
        newname = newname + ' ' + '[无mp3]'
    if group_checked.get():
        newname = newname + ' ' + '[' + group + ']'
    if title_checked.get():
        newname = newname + ' ' + title
    if cv_checked.get() and cv != '':
        newname = newname + ' ' + r'(CV ' + ' '.join(cv.split(';')) + ')'
    newname = re.sub('[\\\/:\*\?"<>\|]', '', newname)
    if move_to_RJ:
        if lcr:
            newname = chinese_RJ_path / newname
        else:
            newname= original_RJ_path/ newname
    else:
        newname=filename.parent / newname
    mv_dir(filename, newname)
    show(f"--已重命名文件夹为{newname.name}")
    return newname


def mv_dir(filename, newname):
    if filename.drive!= newname.drive:
        if filename.drive == 'E:':
            name = filename
            filename = newname
            newname = name
        filename.replace(filename.with_suffix(".voiceWorkTemp"))
        filename = filename.with_suffix(".voiceWorkTemp")
        if (newname.parent / filename.name).exists():
            mv_to_trush(newname.parent / filename.name)
        shutil.move(filename,newname.parent/filename.name)
        if filename.exists():
            mv_to_trush(filename)
        filename= newname.parent / filename.name

    if not filename.exists():
        return
    if newname != filename:
        for file in filename.rglob("*"):
            if file.exists() and file.is_file():
                new = newname / file.relative_to(filename)
                if not new.parent.exists():
                    new.parent.mkdir(parents=True)
                try:
                    if not new.exists():
                        file.replace(new)
                except:
                    pass
        filename=clear(filename)
        if filename.exists():
            mv_to_trush(filename)
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
    if not spider_checked.get():
        return None,None,None
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
        'cookie': 'dlsite_dozen=7; uniqid=dh5wpo6du5; _ga=GA1.2.1094789266.1575957606; _d3date=2019-12-10T06:00:05.977Z; _d3landing=1; _gaid=1094789266.1575957606; adultchecked=1; adr_id=Ub2DpFVcRxLOTwlgsJ1aGEchUBznAqLnj8zWGwt6giSQaWyc; _ebtd=1.k2fnwjp5b.1575971225; locale=zh-cn; Qs_lvt_328467=1577849294^%^2C1577892629^%^2C1577932057^%^2C1591933560; Qs_pv_328467=1052826062922660000^%^2C480504100285960260^%^2C4442059557915936000^%^2C3571252907537905700^%^2C993384634696700900; _ts_yjad=1598842520972; utm_c=blogparts; _inflow_params=^%^7B^%^22referrer_uri^%^22^%^3A^%^22level-plus.net^%^22^%^7D; _inflow_ad_params=^%^7B^%^22ad_name^%^22^%^3A^%^22referral^%^22^%^7D; _im_vid=01ES81S84JB3TTCY5CZRD5EKS5; _gcl_au=1.1.1542253388.1608169517; _gid=GA1.2.1086092439.1612663765; __DLsite_SID=f4ik5q573dn9trjf0rugu9d043; __juicer_sesid_9i3nsdfP_=5c159006-e3e3-4f60-b5fe-dc05c244f1c6; DL_PRODUCT_LOG=^%^2CRJ306930^%^2CRJ300000^%^2CRJ298978^%^2CRJ315336^%^2CRJ315852^%^2CRJ307073^%^2CRJ306798^%^2CRJ309328^%^2CRJ303189^%^2CRJ316357^%^2CRJ234791^%^2CRJ312136^%^2CRJ131395^%^2CRJ282673^%^2CRJ264706^%^2CRJ242260^%^2CRJ250966^%^2CRJ313604^%^2CRJ313754^%^2CRJ295229^%^2CRJ300532^%^2CRJ262976^%^2CRJ311359^%^2CRJ310955^%^2CRJ268194^%^2CRJ289705^%^2CRJ260052^%^2CRJ315474^%^2CRJ316119^%^2CRJ315405^%^2CRJ312692^%^2CRJ167776^%^2CRJ314102^%^2CRJ303183^%^2CRJ309544^%^2CRJ211905^%^2CRJ133234^%^2CRJ307037^%^2CRJ302768^%^2CRJ305343^%^2CRJ299936^%^2CRJ282627^%^2CRJ304923^%^2520^%^2CRJ272689^%^2CRJ303021^%^2CR305282^%^2CRJ297002^%^2CRJ307645^%^2CRJ291292^%^2CRJ295048; _inflow_dlsite_params=^%^7B^%^22dlsite_referrer_url^%^22^%^3A^%^22https^%^3A^%^2F^%^2Fwww.dlsite.com^%^2Fmania x^%^2Fwork^%^2F^%^3D^%^2Fproduct_id^%^2FRJ306798.html^%^22^%^7D; _dctagfq=1356:1613404799.0.0^|1380:1613404799.0.0^|1404:1613404799.0.0^|1428:1613404799.0.0^|1529:1613404799.0.0; __juicer_session_referrer_9i3nsdfP_=5c159006-e3e3-4f60-b5fe-dc05c244f1c6___; _td=287255fd-bbc9-470a-b97d-8c0b1c6b9cd9; _gat=1',
    }
    headers = {
        "authority": "www.dlsite.com",
        "method": "GET",
        "path": f"/maniax/work/=/product_id/{id}.html",
        "scheme": "https",
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "zh-CN,zh;q=0.9,zh-TW;q=0.8",
        "cache-control": "max-age=0",
        "cookie": "uniqid=0.d8q0kidmkq6; dlsite_dozen=7; _ts_yjad=1637081116030; _gaid=1276157997.1637081115; adultchecked=1; __lt__cid=29701fed-f43e-42bf-a284-9ce3d4d378b1; wovn_selected_lang=zh-CHS; carted=1; locale=zh-cn; localesuggested=true; adr_id=Ub2DpFVcRxLOTwlgsJ1aGEchUBznAqLnj8zWGwt6giSQaWyc; _inflow_ad_params=%7B%22ad_name%22%3A%22organic%22%7D; WAPID=KgGRdRXUDlj2xA38ukpc7WWnjlgvqK8Joa4; wap_last_event=showWidgetPage; _im_vid=01GPXE3DKF8HPGZ5A8SYS27DXZ; __DLsite_SID=t9uqp84uq2j4o7a7p6o07j8tft; _gcl_au=1.1.213668751.1676022501; wovn_mtm_showed_langs=%5B%22zh-CHS%22%5D; _gid=GA1.2.1344715748.1676465956; _inflow_params=%7B%22referrer_uri%22%3A%22www.google.com%22%7D; _ga_YG879NVEC7=GS1.1.1676473130.2.0.1676473138.0.0.0; DL_PRODUCT_LOG=%2CRJ300435%2CRJ290120%2CRJ406462%2CRJ374401%2CRJ411265%2CRJ340514%2CRJ419386%2CRJ235149%2CRJ329961%2CRJ252619%2CRJ297628%2CRJ374986; _inflow_dlsite_params=%7B%22dlsite_referrer_url%22%3A%22https%3A%2F%2Fwww.dlsite.com%2Fmaniax%2Fwork%2F%3D%2Fproduct_id%2FRJ300435.html%22%7D; _ga_ZW5GTXK6EV=GS1.1.1676514918.79.1.1676514962.0.0.0; _ga=GA1.2.1276157997.1637081115; OptanonConsent=isGpcEnabled=0&datestamp=Thu+Feb+16+2023+10%3A36%3A07+GMT%2B0800+(%E4%B8%AD%E5%9B%BD%E6%A0%87%E5%87%86%E6%97%B6%E9%97%B4)&version=6.23.0&isIABGlobal=false&hosts=&landingPath=NotLandingPage&groups=C0004%3A1%2CC0003%3A1%2CC0002%3A1%2CC0001%3A1&AwaitingReconsent=false&geolocation=AU%3BNSW",
        "sec-fetch-dest": "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "none",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (iPad; CPU OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/87.0.4280.77 Mobile/15E148 Safari/604.1"
    }
    req = urllib.request.Request(url=url, method="POST",headers=headers)

    try:
        response = urllib.request.urlopen(url)
    except:
        try:
            #使用代理尝试
            proxy = {
                'https': '127.0.0.1:7890'
            }
            handler = urllib.request.ProxyHandler(proxies=proxy)
            # 获取opener对象
            opener = urllib.request.build_opener(handler)
            response = opener.open(url)
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
            if len(dirs) == 1 and len(files) == 0:
                Path.rename(Path(root)/dirs[0], Path(root).parent / (Path(root).name.__str__() +'-'+ dirs[0] + '.voiceWorkTemp'))
                mv_to_trush(root)
                if dirs[0] not in Path(root).stem.split('-'):
                    file = root + '-' + dirs[0]
                else:
                    file = root
                if root == filename.__str__():
                    filename = Path(file)
                Path.rename(Path(root).parent / (Path(root).name.__str__() + '-' + dirs[0] + '.voiceWorkTemp'),Path(file))
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


def extract_mp3_from_video(filesname):
    if not extract_checked.get():
        return
    video= ['.mp4', '.ts','.mkv','.webm']
    if filesname.is_file() and filesname.suffix in video:
        show(f'--开始提取{filesname}')
        audio = AudioSegment.from_file(filesname)
        audio.export(filesname.with_suffix('.mp3'), format="mp3")
        mv_to_trush(filesname)
    else:
        for filename in Path(filesname).rglob('*'):
            if filename.is_file() and filename.suffix in video:
                show(f'--开始提取{filename}')
                audio = AudioSegment.from_file(filename)
                audio.export(filename.with_suffix('.mp3'), format="mp3")
                mv_to_trush(filename)
# 解压缩
def file_unzip(filename, passwd):
    notzip = ['.lrc', '.ass', '.ini', '.url', '.apk', '.heic','.chinese_title','.srt']
    maybezip = ['.rar', '.zip', '.7z','.exe']
    if filename.exists() and filename.is_file():
        if len(filename.suffixes) >= 2 and (mimetypes.guess_type(filename) == (
                None, None) or filename.suffix == '.exe') and filename.suffix not in notzip:
            result = 2
            # 尝试解压
            for pd in passwd:
                output = filename.with_suffix("")
                if output.exists():
                    shutil.rmtree(output, ignore_errors=True)
                cmd = f'bz.exe x -o:"{output}" -aoa -y -p:"{pd}" "{filename}"'
                result = result and os.system(cmd)
                if not result:
                    break
                else:
                    if output.exists():
                        shutil.rmtree(output,ignore_errors=True)
            # 解压成功则
            if not result:
                for file in filename.parent.glob("*"):
                    if file.is_file and file.name.split('.')[0:-2] == filename.name.split('.')[0:-2] and len(
                            file.suffixes) >= 2:
                        mv_to_trush(file)
                filename = output
                if filename.exists():
                    for file in filename.rglob('*'):
                        file_unzip(file, passwd)

        elif filename.suffix in maybezip or (
                mimetypes.guess_type(filename) == (None, None) and filename.suffix not in notzip):
            result = 2
            for pd in passwd:
                output = filename.with_suffix('') if filename.suffix != '' else filename.with_suffix(".voiceWorkTemp")
                if output.exists():
                    shutil.rmtree(output, ignore_errors=True)
                output.mkdir()
                cmd = f'bz.exe x -o:"{output}" -aoa -y -p:"{pd}" "{filename}"'
                result = result and os.system(cmd)
                if not result:
                    break
                else:
                    if output.exists():
                        shutil.rmtree(output, ignore_errors=True)
            if not result:
                mv_to_trush(filename)
                filename = filename.with_suffix('')
                output.replace(filename)
                if filename.exists():
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
        if (not res[1]) and Path(filename).exists():
            os.system('del ' + filename)
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


def spider_switch():
    state = spider_checked.get()
    # if not state:
    for button,value in spider_group:
        value.set(state)
        button.config(state = 'normal' if state else 'disabled')


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
    global spider_checked
    global ops_checked
    ops_checked = checkbox_register('lrc空行间隔')
    global type_checked
    type_checked = checkbox_register('lrc转换为srt', 0)
    global extract_checked
    extract_checked = checkbox_register('mp4或ts提取MP3', 0)
    global archive_checked
    archive_checked = checkbox_register('归档')

    spider_group=[]
    spider_checked = checkbox_register('爬取信息',command=spider_switch)
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
    translate_checked = checkbox_register('翻译标题', group=spider_group)
    global info_text
    info_text = info_register()
    # 主逻辑
    global pool
    pool = ThreadPoolExecutor(max_workers=6)
    windnd.hook_dropfiles(window, func=dragged_files, force_unicode='utf-8')
    tk.mainloop()

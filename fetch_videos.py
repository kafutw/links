"""抓卡夫卡愛旅行頻道最新的長影片（不含 Shorts），寫進 videos.json 給連結頁用。
先讀 RSS，讀不到再改抓頻道影片頁；兩個都失敗就保留原本的 videos.json。"""
import json, re, sys, urllib.request
import xml.etree.ElementTree as ET

CHANNEL = 'UCqrIaayLSSFAxk2z9hRrJ5A'
COUNT = 2
UA = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36',
      'Accept-Language': 'zh-TW,zh;q=0.9'}

def get(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
        return r.read().decode('utf-8')

def from_rss():
    ns = {'a': 'http://www.w3.org/2005/Atom', 'yt': 'http://www.youtube.com/xml/schemas/2015'}
    root = ET.fromstring(get(f'https://www.youtube.com/feeds/videos.xml?channel_id={CHANNEL}'))
    out = []
    for e in root.findall('a:entry', ns):
        if '/shorts/' in e.find('a:link', ns).get('href', ''):
            continue
        out.append({'id': e.find('yt:videoId', ns).text, 'title': e.find('a:title', ns).text})
    return out[:COUNT]

def from_page():
    html = get(f'https://www.youtube.com/channel/{CHANNEL}/videos')
    ids = []
    for v in re.findall(r'"videoId":"([\w-]{11})"', html):
        if v not in ids:
            ids.append(v)
    out = []
    for v in ids[:COUNT]:
        d = json.loads(get(f'https://www.youtube.com/oembed?url=https://www.youtube.com/watch?v={v}&format=json'))
        out.append({'id': v, 'title': d['title']})
    return out

videos = []
for fn in (from_rss, from_page):
    try:
        videos = fn()
        if len(videos) == COUNT:
            print(f'{fn.__name__}: {[v["id"] for v in videos]}')
            break
    except Exception as e:
        print(f'{fn.__name__} 失敗：{e}', file=sys.stderr)
if len(videos) != COUNT:
    sys.exit('抓不到最新影片，保留原本的 videos.json')
with open('videos.json', 'w', encoding='utf-8') as f:
    json.dump(videos, f, ensure_ascii=False, indent=1)
    f.write('\n')

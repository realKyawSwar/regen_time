import pandas as pd
import re
import numpy as np
from operator import itemgetter
from itertools import groupby

ser_str = r"(k-s-s4s-\d{4})|(k-s-s413-\d{4})|(27p\d{5}[a-z])|(14z\d{5}[a-z]{1,2})|(v\d{6}-\d{2})|(\d{4}che\d{5})|(\d{4}-\d{5})|([1-2]\d{3})"


def remove_middle(templist1):
    listx = []
    serial = f'({ser_str})'
    new_re = f'new.*?{serial}'
    old_re = f'old.*?{serial}'
    for i in templist1:
        if re.search(new_re, i):
            serial_match = re.search(serial, i).group(0)
            listx.append(re.sub(new_re, 'new ' + serial_match, i))
        elif re.search(old_re, i):
            serial_match = re.search(serial, i).group(0)
            listx.append(re.sub(old_re, 'old ' + serial_match, i))
        else:
            listx.append(i)
    return listx


def coolfunc(ng):
    # ['new 14z05180bs', 'p11 pump old 14z15092c']
    serial = f'({ser_str})'
    new_re = f'new.*?{serial}'
    old_re = f'old.*?{serial}'
    chamber = r"(\bp\d{1,2}\b)|(\bc\d{1}\b)"
    listx = []
    for ele in ng:
        listy = []
        for i in ele:
            if re.search(old_re, i):
                a = re.search(old_re, i).group(0)
                listy.append(a)
            elif i.startswith('old') and re.search(serial, i):
                a = 'old ' + re.search(serial, i).group(0)
                listy.append(a)
        listx.append(listy)
    listj = []
    for ele in ng:
        listy = []
        for i in ele:
            if re.search(new_re, i):
                a = re.search(new_re, i).group(0)
                listy.append(a)
            elif i.startswith('new') and re.search(serial, i):
                a = 'new ' + re.search(serial, i).group(0)
                listy.append(a)
        listj.append(listy)
    listz = []
    for ele in ng:
        listy = []
        for i in ele:
            if i.startswith('load'):
                listy.append('load')
            elif i.startswith('unload'):
                listy.append('unload')
            elif re.search(chamber, i):
                a = re.search(chamber, i).group(0)
                listy.append(a)
        listz.append(listy)
    # print(listz)
    listw = [a+b+c for a, b, c in zip(filter(None, listz),
                                      filter(None, listx),
                                      filter(None, listj))]
    listv = [i for i in listw if len(i) == 3]
    return [list(set(remove_middle(i))) for i in listv]


def extract(words):
    """Remove punctuation from list of tokenized words"""
    new_words = []
    stringy = words.splitlines()
    list_ = ['pump', 'cylinder', 'tdu', 'robot', 'arm', 'compressor', 'motor']
    list1_ = ['driver', 'alarm', 'rmc motor', 'tdu motor', 'speed', 'ring',
              'mandrel', 'pub', 'pick', 'cable', 'solenoid', 'sensor', 'gauge',
              'remote', 'coupling', 'cathode', 'flexible', 'config', 'setting',
              'limit', 'app', 'damper', 'temp']
    for word in stringy:
        # remove puntuation and
        new_word = re.sub(r'[^\w\s\'\"]', ' ', word).strip().lower()
        # change
        new_word = re.sub(r'^chg', 'action', new_word)
        new_word = re.sub(r'(\bld\b)', 'load', new_word)
        new_word = re.sub(r'(\buld\b)', 'unload', new_word)
        # remove double space
        new_word = re.sub(re.compile(r' +'), ' ', new_word)
        if new_word.startswith("change") or new_word.startswith("replace"):
            if any(word in new_word for word in list_) and all(word not in new_word for word in list1_):
                # new_words.append(new_word)
                for i in ['change ', ' joshua', ' ngai', ' toh kc', ' lim th',
                          ' koo', ' vincent toh']:
                    new_word = new_word.replace(i, '')

                new_words.append(new_word)
    return new_words


def extract_serial(words):
    number_list = []
    stringy = words.splitlines()
    stringx = [re.sub(r'(:|=|<|>|/)', '', i).strip().lower() for i in stringy]
    stringz = [re.sub(r'(\s{3}|\s{2})', ' ', i) for i in stringx]
    stringj = [re.sub(r'(\bld\b)', 'load', i) for i in stringz]
    stringj = [re.sub(r'(\bcompressor\b)', 'crc', i) for i in stringj]
    raw = [i.replace('new sn',
                     'new').replace('old sn',
                                    'old').replace('uld',
                                                   'unload') for i in stringj]
    temp_list = []
    for i, new_word in enumerate(raw):
        research = re.search(r"((k-s-s4s-\d{4}$)|(k-s-s413-\d{4}$)|(27p\d{5}[a-z]$)|(14z\d{5}[a-z]{1,2}$)|(v\d{6}-\d{2}$)|(\d{4}che\d{5}$)|(\d{4}-\d{5}$)|([1-2]\d{3}$))", new_word)
        research1 = re.search(f"({ser_str})", new_word)
        if new_word.startswith(('old', 'new')) and research:
            # old.append(new_word)
            number_list.append(i)
        elif research or research1:
            temp_list.append(new_word)
    # [58, 59],[62, 63],[66, 67] output
    clean_lol = []
    matches = ['load', 'unload', 'vtc']
    for k, g in groupby(enumerate(number_list), key=lambda x: x[0] - x[1]):
        pair = list(map(itemgetter(1), g))
        add_min_item = min(pair) - 1
        pair.append(add_min_item)
        raw2 = [raw[i] for i in pair if raw[i]]
        if len(raw2) == 3 and re.search(r"(\bp\d{1,2}\b)|(\bc\d{1}\b)|(\bcrc\s\d{1})", raw2[2]):
            raw2[2] = re.search(r"(\bp\d{1,2}\b)|(\bc\d{1}\b)|(\bcrc\s\d{1})", raw2[2]).group(0)
        elif len(raw2) == 3 and any(x in raw2[2] for x in matches):
            raw2[2] = re.search(r"\bload\b|\bunload\b|\bvtc\b", raw2[2]).group(0)
        clean_lol.append(raw2)
    templist1 = []
    if temp_list:
        for i in temp_list:
            new = re.search(f"(\bnew\b.*?({ser_str}))", i)
            if new:
                templist1.append(new.group(0))
        for i in temp_list:
            old = re.search(f"(\bold\b.*?({ser_str}))", i)
            if old:
                templist1.append(old.group(0))
    if temp_list and templist1:
        chamber = re.search(r"(\bp\d{1,2}\b)|(\bc\d{1}\b)", temp_list[0])
        otherch = re.search(r"\bload\b|\bunload\b|\bvtc\b|\bhll\b|\bctc\b|crc\s\d{1}",
                            temp_list[0])
        if chamber:
            templist1.append(chamber.group(0))
        elif otherch:
            templist1.append(otherch.group(0))
        else:
            pass
    listx = []
    if templist1:
        listx = remove_middle(templist1)
    clean_lol2 = []
    for i in clean_lol:
        # print(i)
        clean_lol2.append(remove_middle(i))
    for data in clean_lol2:
        # print(data)
        i = sum('old' in s for s in data)
        j = sum('new' in s for s in data)
        if i > 1 or j > 1:
            data.pop(0)
            # print(data)
    # print(clean_lol2)
    if listx:
        clean_lol2.append(listx)
    return clean_lol2


def determine_ok(listy):
    return [i for i in listy if len(i) == 3]


def determine_ng(listy):
    return [i for i in listy if len(i) != 3]


def match(i):
    ch = ''
    chamber = re.search(r"(\bp\d{1,2}\b)|(\bc\d{1}\b)", i)
    otherch = re.search(r"\bload\b|\bunload\b|\bvtc\b|\bhll\b|\bctc\b|crc\s\d{1}", i)
    if chamber:
        ch = chamber.group(0)
    elif otherch:
        ch = otherch.group(0)
    # lol.append(ch)
    return ch


df = pd.read_excel('trouble_report.xlsx')
df['job'] = df.ACTION.apply(lambda x: extract(x))
df['job'] = df['job'].apply(lambda x: ','.join(map(str, x)))
df['job'].replace('', np.nan, inplace=True)
df.dropna(subset=['job'], inplace=True)
df['serial_no'] = df.ACTION.apply(lambda x: extract_serial(x))
df['serial_no'] = df['serial_no'].apply(lambda x: '' if not x else x)
df['good'] = df.serial_no.apply(lambda x: determine_ok(x))
# df['good'] = df.good.apply(lambda x: '' if not x else x)
df['bad'] = df.serial_no.apply(lambda x: determine_ng(x))
df['bad'] = df.bad.apply(lambda x: '' if not x else x)
df['another'] = df.bad.apply(lambda x: coolfunc(x))
df['bad'] = df.apply(lambda x: x['bad'] if not x['another'] else '', axis=1)
df['another'] = df.another.apply(lambda x: '' if not x else x)
df['good'] = df.apply(lambda x: x['good']+x['another'] if x['another'] else x['good'], axis=1)
df.reset_index(inplace=True)
df['lol'] = df.job.apply(lambda x: match(x))
df['lel'] = df.bad.apply(lambda x: ','.join(x[0]) if len(x) == 1 else [])
df['temp'] = df.apply(lambda x: x['lel']+','+x['lol'] if x['lol'] and x['lel'] else '', axis=1)
df['temp'] = df.temp.apply(lambda x: [x.split(",")])
df['temp'] = df.temp.apply(lambda x: coolfunc(x))
df['bad'] = df.apply(lambda x: '' if x['temp'] else x['bad'], axis=1)
df['bad'] = df.bad.apply(lambda x: np.nan if not x else x)
df['good'] = df.apply(lambda x: x['good']+x['temp'] if x['temp'] else x['good'], axis=1)
df['good'] = df.good.apply(lambda x: np.nan if not x else x)
# df = df[['job', 'bad', 'temp']]
df.rename(columns={'TROUBLE_DATE': 'date', 'LINE_ID': 'line'},
          inplace=True)
df['date'] = pd.to_datetime(df['date'])
df['date'] = df['date'].dt.strftime('%y-%m-%d')
df_bad = df[['line', 'date', 'bad']]
df_bad.dropna(subset=['bad'], inplace=True)
df_good = df[['line', 'date', 'good']]
df_good.dropna(subset=['good'], inplace=True)

df_bad.to_excel("final7.xlsx")
df_good.to_excel("fi.xlsx")
# print(df.serial_no.iloc[14])
# with pd.option_context('display.max_rows', None, 'display.max_columns', None):
#     print(df.temp)
# df[df.bad.map(len) == 1]
# df.loc[np.array(list(map(len, df.bad.values)))==1]

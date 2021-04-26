import pandas as pd
import re
import numpy as np
from operator import itemgetter
from itertools import groupby


def extract(words):
    """Remove punctuation from list of tokenized words"""
    new_words = []
    stringy = words.splitlines()
    list_ = ['pump', 'cylinder', 'tdu', 'robot', 'arm', 'compressor', 'motor']
    list1_ = ['driver', 'alarm', 'rmc motor', 'tdu motor', 'speed', 'ring',
              'mandrel', 'pub', 'pick', 'cable', 'solenoid']
    for word in stringy:
        # remove puntuation and
        new_word = re.sub(r'[^\w\s\'\"]', ' ', word).strip().lower()
        # change
        new_word = re.sub(r'^chg', 'action', new_word)
        # remove double space
        new_word = re.sub(re.compile(r' +'), ' ', new_word)
        if new_word.startswith("change") or new_word.startswith("replace"):
            if any(word in new_word for word in list_) and all(word not in new_word for word in list1_):
                # new_words.append(new_word)
                for i in ['change ', ' joshua', ' ngai', ' toh kc', ' lim th',
                          ' koo', ' vincent toh']:
                    new_word = new_word.replace(i, '')
                new_words.append(new_word.capitalize())
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
        research1 = re.search(r"((k-s-s4s-\d{4})|(k-s-s413-\d{4})|(27p\d{5}[a-z])|(14z\d{5}[a-z]{1,2})|(v\d{6}-\d{2})|(\d{4}che\d{5})|(\d{4}-\d{5})|([1-2]\d{3}))", new_word)
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
        raw2 = [raw[i] for i in pair]
        if len(raw2) == 3 and re.search(r"(\bp\d{1,2}\b)|(\bc\d{1}\b)", raw2[2]):
            raw2[2] = re.search(r"(\bp\d{1,2}\b)|(\bc\d{1}\b)", raw2[2]).group(0)
        elif len(raw2) == 3 and any(x in raw2[2] for x in matches):
            raw2[2] = re.search(r"\bload\b|\bunload\b|\bvtc\b", raw2[2]).group(0)
        clean_lol.append(raw2)
    templist1 = []
    if temp_list:
        for i in temp_list:
            new = re.search(r"(\bnew\b.*?((k-s-s4s-\d{4})|(k-s-s413-\d{4})|(27p\d{5}[a-z])|(14z\d{5}[a-z]{1,2})|(v\d{6}-\d{2})|(\d{4}che\d{5})|(\d{4}-\d{5})|([1-2]\d{3})))", i)
            if new:
                templist1.append(new.group(0))
            else:
                pass
    if temp_list:
        for i in temp_list:
            old = re.search(r"(\bold\b.*?((k-s-s4s-\d{4})|(k-s-s413-\d{4})|(27p\d{5}[a-z])|(14z\d{5}[a-z]{1,2})|(v\d{6}-\d{2})|(\d{4}che\d{5})|(\d{4}-\d{5})|([1-2]\d{3})))", i)
            if old:
                templist1.append(old.group(0))
            else:
                pass
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
    # print(temp_list)
    return clean_lol+templist1


df = pd.read_excel('trouble_report.xlsx')
df['job'] = df.ACTION.apply(lambda x: extract(x))
df['job'] = df['job'].apply(lambda x: ','.join(map(str, x)))
df['job'].replace('', np.nan, inplace=True)
df.dropna(subset=['job'], inplace=True)
df['serial_no'] = df.ACTION.apply(lambda x: extract_serial(x))
df['serial_no'] = df['serial_no'].apply(lambda x: '' if len(x)==0 else x)
df = df[['LINE_ID', 'TROUBLE_DESC', 'TROUBLE_DATE', 'job', 'serial_no']]
df.reset_index(inplace=True)
df.to_excel("final2.xlsx")
# print(df.serial_no.iloc[14])

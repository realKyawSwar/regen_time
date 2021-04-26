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
    # new = []
    # old = []
    number_list = []
    stringy = words.splitlines()
    stringx = [re.sub(r'(:|=|<|>|/)', '', i).strip().lower() for i in stringy]
    stringz = [re.sub(r'(\s{3}|\s{2})', ' ', i) for i in stringx]
    lol = [i.replace('new sn ', 'new ').replace('old sn ', 'old ') for i in stringz]
    for i, new_word in enumerate(lol):
        if new_word.startswith(('old', 'new')) and re.search(r"((k-s-s4s-\d{4}$)|(k-s-s413-\d{4}$)|(27p\d{5}[a-z]$)|(14z\d{5}[a-z]{1,2}$)|(v\d{6}-\d{2}$)|(\d{4}che\d{5}$)|(\d{4}-\d{5}$)|([1-2]\d{3}$))", new_word):
            # old.append(new_word)
            number_list.append(i)
        elif re.search(r"((k-s-s4s-\d{4}$)|(k-s-s413-\d{4}$)|(27p\d{5}[a-z]$)|(14z\d{5}[a-z]{1,2}$)|(v\d{6}-\d{2}$)|(\d{4}che\d{5}$)|(\d{4}-\d{5}$)|([1-2]\d{3}$))", new_word):
            number_list.append(i)
    # listy = [stringz[i-1] for i in number_list]
    # good_list = np.unique(old+listy)
    # [58, 59],[62, 63],[66, 67] output
    lolx = []
    for k, g in groupby(enumerate(number_list), key=lambda x: x[0] - x[1]):
        pair = list(map(itemgetter(1), g))
        add_min_item = min(pair) - 1
        pair.append(add_min_item)
        lolx.append([lol[i] for i in pair])
    return lolx


df = pd.read_excel('trouble_report.xlsx')
df['job'] = df.ACTION.apply(lambda x: extract(x))
df['job'] = df['job'].apply(lambda x: ','.join(map(str, x)))
df['job'].replace('', np.nan, inplace=True)
df.dropna(subset=['job'], inplace=True)
df['serial_no'] = df.ACTION.apply(lambda x: extract_serial(x))
df['serial_no'] = df['serial_no'].apply(lambda x: ','.join(map(str, x)))
df = df[['LINE_ID', 'TROUBLE_DESC', 'TROUBLE_DATE', 'job', 'serial_no']]
df.to_excel("final1.xlsx")
# print(df.loc[5])
# print(df.ACTION[5])
# listy = tuple(new_word.split(' '))
# elif new_word.startswith('new') and re.search(r"((k-s-s4s-\d{4}$)|(k-s-s413-\d{4}$)|(27p\d{5}[a-z]$)|(14z\d{5}[a-z]{1,2}$)|(v\d{6}-\d{2}$)|(\d{4}che\d{5}$)|(\d{4}-\d{5}$)|([1-2]\d{3}$))", new_word):
#     listx = tuple(new_word.split(' '))
#     new.append(listx)
# elif re.search(r"((k-s-s4s-\d{4}$)|(k-s-s413-\d{4}$)|(27p\d{5}[a-z]$)|(14z\d{5}[a-z]{1,2}$)|(v\d{6}-\d{2}$)|(\d{4}che\d{5}$)|(\d{4}-\d{5}$)|([1-2]\d{3}$))", new_word):
#     new_words.append(new_word)

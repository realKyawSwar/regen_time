import pandas as pd
import numpy as np
import re


df = pd.read_csv('lol.csv')

df = df.drop(columns=['MACHINE_ID',
                      'PROCESS_ID', 'PROCESS_DESC', 'EQUIP_ID', 'COMMENT_TIME',
                      'EQUIP_DESC', 'UPDATED_DATE', 'UPDATED_TIME',
                      'ENGINEER_COMMENT', 'COMMENT_BY', 'COMMENT_DATE',
                      'ATTEND_BY', 'UPDATED_BY', 'CAUSES', 'TROUBLE_ID'])
# df = df.drop(columns=['TROUBLE_STIME', 'TROUBLE_ETIME', 'MACHINE_ID',
#                       'PROCESS_ID', 'PROCESS_DESC', 'EQUIP_ID', 'COMMENT_TIME',
#                       'EQUIP_DESC', 'UPDATED_DATE', 'UPDATED_TIME',
#                       'ENGINEER_COMMENT', 'COMMENT_BY', 'COMMENT_DATE',
#                       'ATTEND_BY', 'UPDATED_BY', 'CAUSES', 'TROUBLE_ID'])
# header = df.columns.values.tolist()
# print(header)
# report = df.loc[df.index, 'ACTION'].iat[0]
# 'REGEN' OR TROUBLE_DESC= 'CONVERSION/PLAN DOWN'
df.TROUBLE_DATE = pd.to_datetime(df.TROUBLE_DATE, format="%Y-%m-%d")
# df = df[(df['TROUBLE_DESC'] == 'REGEN') | (df['TROUBLE_DESC'] == 'CONVERSION/PLAN DOWN')]
# df = df[(df.TROUBLE_DATE > '2020-01-01') & (df.TROUBLE_DATE <= '2020-10-29')]
# df.to_csv('ori.csv')
# print(df.head(10))
# report = df.loc[df.index, 'ACTION'].iat[10]
# # tokens = nltk.word_tokenize(report)
# clean_string = report.splitlines()


def extract(words):
    """Remove punctuation from list of tokenized words"""
    new_words = []
    stringy = words.splitlines()
    list_ = ['pump', 'cylinder', 'tdu', 'robot', 'arm', 'compressor', 'motor']
    list1_ = ['driver', 'alarm', 'rmc motor', 'tdu motor', 'speed', 'ring',
              'mandrel', 'pub', 'pick']
    for word in stringy:
        # remove puntuation and
        new_word = re.sub(r'[^\w\s\'\"]', ' ', word).strip().lower()
        # change
        new_word = re.sub(r'^chg', 'change', new_word)
        # remove double space
        new_word = re.sub(re.compile(r' +'), ' ', new_word)
        if new_word.startswith("change") or new_word.startswith("replace"):
            if any(word in new_word for word in list_) and all(word not in new_word for word in list1_):
                new_words.append(new_word)
            # elif re.search(r"(([\d]+[a-z]{3,})|([a-z]+[\d]{3,}))", new_word) or re.search(r"([\d]-{3,})", new_word):
            # elif re.search(r"(([\d]+[a-z]{3,})|([a-z]+[\d]{3,}))", new_word):
            #     new_words.append(new_word)
    return new_words


def extract_ser(words):
    """Remove punctuation from list of tokenized words"""
    new_words = []
    stringy = words.splitlines()
    for word in range(1, len(stringy)):
        # remove puntuation and
        new_word = re.sub(r'[^\w\s-]', ' ', word).strip().lower()
        # change
        new_word = re.sub(r'^chg', 'change', new_word)
        # remove double space
        new_word = re.sub(re.compile(r' +'), ' ', new_word)
        # 100che100 , 123dn54517, 27p05071d, 14z10061c
        if re.search(r"([\d]+[a-z]+[\d]+[a-z]{1,})|([\d]+[a-z]+[\d]{3,})",
                     new_word):
            new_words.append(new_word)
        elif re.search(r"([a-z]+[\d]{1,}-[\d]{2,})|([a-z]-[a-z]-[a-z][\d]{1,}-[\d]{3,})", new_word):
            # k-s-s413-0010, v063498-12
            new_words.append(new_word)
        elif re.search(r"([a-z]{1,}-[a-z][\d]-[\d]{3,})|([a-z][\d]{5,})|([\d]{3,}-[\d]{3,})|([\d]{5})", new_word):
            new_words.append(new_word)
    return new_words


pd.set_option('display.max_columns', None)
df['CHANGE'] = df.ACTION.apply(lambda x: extract(x))
df = df.drop(df[df.CHANGE.map(len) < 1].index)
df = df.drop(['DURATION', 'TROUBLE_DESC'], axis=1)
df = df[['LINE_ID', 'TROUBLE_DATE', 'TROUBLE_STIME',
         'TROUBLE_ETIME', 'ACTION', 'CHANGE']]
df.ACTION.apply(lambda x: x.splitlines())
df = df.assign(ACTION=df['ACTION'].str.split('\n')).explode('ACTION')
df = df.replace('\r', '', regex=True)
df = df[df['ACTION'].str.strip().astype(bool)]
multi = df.set_index(['TROUBLE_DATE', 'LINE_ID', 'TROUBLE_STIME',
                      'TROUBLE_ETIME']).sort_index()
# multi.to_csv('combined.csv')
multi.to_excel("combined.xlsx", sheet_name='Regen Report')
# print(multi.tail(10))
print('done')

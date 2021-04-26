import re
import pandas as pd
import numpy as np


def extract_serial(words):
    new_words = []
    stringy = words.splitlines()
    for word in stringy:
        new_word = re.sub(r'(:|=|<|>)', '', word).strip().lower()
        new_word = re.sub(r'(\s{3}|\s{2})', ' ', new_word)
        if new_word.startswith(('old', 'new')) and re.search(r"((k-s-s4s-\d{4}$)|(k-s-s413-\d{4}$)|(27p\d{5}[a-z]$)|(14z\d{5}[a-z]{1,2}$)|(v\d{6}-\d{2}$)|(\d{4}che\d{5}$)|(\d{4}-\d{5}$)|([1-2]\d{3}$))", new_word):
            new_words.append(new_word)
    return new_words


df = pd.read_excel('trouble_report.xlsx')
df['Activities'] = df.ACTION.apply(lambda x: extract_serial(x))
df['Activities'] = df['Activities'].apply(lambda x: ','.join(map(str, x)))
df = df[['LINE_ID', 'TROUBLE_DESC', 'TROUBLE_DATE', 'Activities']]
df['Activities'].replace('', np.nan, inplace=True)
df.dropna(subset=['Activities'], inplace=True)
df.to_excel('after_serial1.xlsx')
# print(lol[5])


# new = r"((K-S-S4S-\d{4}$)|(K-S-S413-\d{4}$)|(27P\d{5}[A-Z]$)|(14Z\d{5}[A-Z]{1,2}$)|(V\d{6}-\d{2}$)|(\d{4}CHE\d{5}$)|(\d{4}-\d{5}$)|([1-2]\d{3}$))"
# print(new.lower())

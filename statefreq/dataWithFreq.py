from sklearn.feature_extraction.text import CountVectorizer
from nltk.corpus import stopwords
import pandas as pd
import numpy as np

df = pd.read_excel("Vaxart Survey 2021.xlsx", None)

def freq(n_range, text):
    sw = stopwords.words('english') + ['really']

    vector = CountVectorizer(stop_words=sw, ngram_range=(2,2))
    ngrams = vector.fit_transform(text)
    count_values = ngrams.toarray().sum(axis=0)
    vocab = vector.vocabulary_
    df_ngram = pd.DataFrame(sorted([(count_values[i],k) for k,i in vocab.items()], reverse=True)).rename(columns={0: 'frequency', 1:'PHRASES'})
    return df_ngram

key = list(df)
data = {}
for i in range(len(df)):
    index = key[i]
    if df[index]["Unnamed: 2"].dtype == 'O':
        data[index] = df[index]["Unnamed: 2"].dropna()
        data[index] = [j for j in data[index] if not isinstance(j, int)]
        
key = list(data.keys())
for i in range(len(data)):
    curr = key[i]
    text = data[curr]
    text = ' '.join(text)
    text = text.split('.')
    print(curr, ": ")
    print(freq([2,3], text))


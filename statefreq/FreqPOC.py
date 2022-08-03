from sklearn.feature_extraction.text import CountVectorizer
from nltk.corpus import stopwords
import pandas as pd

def freq(n_range, text):
    sw = stopwords.words('english')

    vector = CountVectorizer(stop_words=sw, ngram_range=n_range)
    ngrams = vector.fit_transform(text)
    count_values = ngrams.toarray().sum(axis=0)
    vocab = vector.vocabulary_
    df_ngram = pd.DataFrame(sorted([(count_values[i],k) for k,i in vocab.items()], reverse=True)).rename(columns={0: 'frequency', 1:'bigram/trigram'})
    print(df_ngram)
    return df_ngram
    
text = ['UCSD TCG', 'TCG is cool', 'TCG is fun', 'TCG is cool', 'UCSD TCG', 'UCSD TCG', 'Dani']
lengths = [2,3]

freq(lengths, text)


def freq_more_words(n_range, text):
    ranges = [[n_range[i-1], n_range[i]] for i in range(1, len(n_range), 2)]
    sw = stopwords.words('english')
    for r in ranges:
        vector = CountVectorizer(stop_words=sw, ngram_range=r)
        ngrams = vector.fit_transform(text)
        count_values = ngrams.toarray().sum(axis=0)
        vocab = vector.vocabulary_
        df_ngram = pd.DataFrame(sorted([(count_values[i],k) for k,i in vocab.items()], reverse=True)).rename(columns={0: 'frequency', 1:'bigram/trigram'})
        print(df_ngram)
    return df_ngram
    
text = ['UCSD TCG', 'TCG is cool', 'TCG is fun', 'TCG is cool', 'UCSD TCG', 'UCSD TCG', 'Dani']
lengths = [2,3,4, 5]
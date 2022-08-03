from sklearn.feature_extraction.text import CountVectorizer
from nltk.corpus import stopwords

def freq(n, text):
    sw = stopwords.words
    c_vec = CountVectorizer(stop_words=sw,)
    grams = c_vec.fit_transform(text)
    vocab = c_vec.vocabulary_
    print(vocab)
    return vocab
    
text = ['UCSD TCG', 'TCG is cool', 'TCG is fun', 'TCG is cool', 'UCSD TCG', 'UCSD TCG', 'Dani']

freq(2, text)
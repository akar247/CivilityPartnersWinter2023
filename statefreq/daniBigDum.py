import numpy as np
import re
import nltk
from sklearn.datasets import load_files
nltk.download('stopwords')
import pickle
from nltk.corpus import stopwords

df = pd.read_excel("Vaxart Survey 2021.xlsx", None)
train = pd.read_excel('TrainingData.xlsx')

key = list(df)
data = {}
for i in range(len(df)):
    index = key[i]
    if df[index]["Unnamed: 2"].dtype == 'O':
        data[index] = df[index]["Unnamed: 2"].dropna()
        data[index] = [j for j in data[index] if not isinstance(j, int)]


import numpy as np
import pandas as pd
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import re

df25=pd.DataFrame(lst25, columns=['response'])
df25.iloc[0]

def sentiment_score(response): #score from 1 to 5
    tokens = tokenizer.encode(response, return_tensors='pt')
    result = model(tokens)
    return int(torch.argmax(result.logits))+1
    
sentiment_score(df25['response'].iloc[1])
df25['sentiment'] = df25['response'].apply(lambda x: sentiment_score(x[:512]))

df25 = df25.sort_values('sentiment', ascending=False)

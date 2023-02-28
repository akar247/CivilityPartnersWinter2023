import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification, pipeline
import pandas as pd

tokenizer = AutoTokenizer.from_pretrained("DhruvK0/test_trainer")

print("Downloaded tokenizer")

model = AutoModelForSequenceClassification.from_pretrained("DhruvK0/test_trainer")

print("Downloaded model")

classifier = pipeline("sentiment-analysis", model=model, tokenizer=tokenizer)

print("Created pipeline for sentiment analysis")

input_file = input('Enter the file name: ')

print("Reading input file")

df1 = pd.read_excel(input_file)

print("Input file read")

def give_sentiment(text):
    return classifier(text)[0]['label']


print("Applying sentiment analysis")

df1['Output'] = df1['Input'].apply(give_sentiment)

print("Applied sentiment analysis")

with pd.ExcelWriter('output2.xlsx') as writer:  
    df1.to_excel(writer, sheet_name='Sheet1')

print("Wrote to output2.xlsx")

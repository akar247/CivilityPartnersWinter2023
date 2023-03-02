import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification, pipeline
import pandas as pd

tokenizer = AutoTokenizer.from_pretrained("DhruvK0/test_trainer")

print("Downloaded tokenizer")

model = AutoModelForSequenceClassification.from_pretrained("DhruvK0/test_trainer")

print("Downloaded model")

classifier = pipeline("sentiment-analysis", model=model, tokenizer=tokenizer)

print("Created pipeline for sentiment analysis")


df1 = pd.read_excel('demo.xlsx')

print("Input file read")

def give_sentiment(text):
    return classifier(text)[0]['label']

print("Applying sentiment analysis")

df1['Output'] = df1['Input'].apply(give_sentiment)

print("Applied sentiment analysis")

with pd.ExcelWriter('demo_output.xlsx') as writer:  
    df1.to_excel(writer, sheet_name='Sheet1')

print("Wrote to demo_output.xlsx")

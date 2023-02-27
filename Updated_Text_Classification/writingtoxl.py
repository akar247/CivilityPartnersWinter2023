from transformers import AutoTokenizer, AutoModelForSequenceClassification, pipeline
import pandas as pd

tokenizer = AutoTokenizer.from_pretrained("DhruvK0/test_trainer")

model = AutoModelForSequenceClassification.from_pretrained("DhruvK0/test_trainer")

classifier = pipeline("sentiment-analysis", model = model, tokenizer = tokenizer)

input_file = input('Enter the file name: ')
df1 = pd.read_excel(input_file)

def give_sentiment(text):
    return classifier(text)

df1['Output'] = df1['Input'].apply(give_sentiment)
with pd.ExcelWriter('output.xlsx') as writer:  
    df1.to_excel(writer, sheet_name='Sheet1')
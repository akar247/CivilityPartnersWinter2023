
#from transformers import AutoTokenizer, AutoModelForSequenceClassification
import pandas as pd
import numpy as np

# tokenizer = AutoTokenizer.from_pretrained("DhruvK0/test_trainer")

# model = AutoModelForSequenceClassification.from_pretrained("DhruvK0/test_trainer")

#script to take in file input
#converts file to dataframe
#import trained model from huggingface
#for each row add a new column with output for text in that row
# return new dataframe as a csv file

input_data_file = input("Enter the file name: ")

#convert file to dataframe
df = pd.read_csv(input_data)
df.columns = ['text']
print(df)
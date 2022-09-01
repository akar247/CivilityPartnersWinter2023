from tkinter import *
from tkinter import filedialog
from tkinter import font
from tkinter.ttk import Progressbar
from tkinter.ttk import Scrollbar
import os
import csv
import ntpath
import pandas as pd

import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords 
from nltk.stem.snowball import SnowballStemmer

#Constant vars
stemmer = SnowballStemmer(language = "english")
stopwords = set(stopwords.words('english')) 
xfile = 'Vaxart Survey 2021.xlsx'

#Universal Vars
#list of classes they want
classes = []
#list of lists of keys words for each class. Outside list should have same size as classes list
kws = []
#Storeing all responses
data = {}

#DEFAULT TESTING
job = ['job', 'task', 'responsibility', 'promotion', 'career', 'pay', 'salary', 'compensation', 'possibilities', 'paid', 'flexibility', 'efficiency', 'advancement', 'growth']
engagement = ['challenging', 'engagement', 'environment', "integrity", "meaningful", "timeline", "motivated", "goal", "focus"]
communication = ['communication', 'internal', 'organize', 'organization', 'goal', 'vision', 'communicate', 'collaborate', 'team', 'performance', 'manager', 'teach', 'ask', 'feedback', 'supervisor', 'transparent' ]
inclusion = ['inclusive', 'diversity', 'enforce', 'harass', 'comfort', 'equality', 'behavior', 'hire', 'retain', 'talent', 'recruitment']
relationship = ['management', 'relation', 'relationship', 'leadership', 'transparency', 'cohesiveness', 'accessible', 'straining', 'collaborative', 'acceptance', 'supportive', 'respectful', 'appreciative', 'friendly', 'approachable', 'honest', 'relaxed', 'casual', 'toxic', 'dynamic', 'culture', 'moral', 'trust'] 

classes = ['Job', 'engagement', 'communication', 'Inclusion', 'Relationship']
kws = [job, engagement, communication, inclusion, relationship]

def grab_data():
    global xfile, data
    
    df = pd.read_excel(xfile, None)
    
    key = list(df)
    for i in range(len(df)):
        index = key[i]
        if df[index]["Unnamed: 2"].dtype == 'O':
            data[index] = df[index]["Unnamed: 2"].dropna()[1:]
            data[index] = [j for j in data[index] if not isinstance(j, int)]
    
    return None

def classify():
    global classes, kws, stemmer, stopwords, data
    
    classified = {'Miscellaneous': []}
    
    with open('Info.txt', 'r') as f:
        
        for i in range(len(kws)):
            kws[i] = [stemmer.stem(word) for word in kws[i]]
        

    for lst in range(len(kws)):
        for i in range(len(kws[lst])):
            #Apply Stemmer Filter on all single word keywords
            #Cannot stem phrases
            if ' ' not in kws[lst][i].strip():
                kws[lst][i] = stemmer.stem(kws[lst][i])
        #create empty dictionary of classes key values
        classified[classes[lst]] = []
    
    for question in data:
        #foratting response
        responses = data[question]
        
        for response in responses:
            response_tok = word_tokenize(response)
            clean_text = [stemmer.stem(word) for word in response_tok if word.isalpha() and word not in stopwords]
            
            sims = [0]*len(kws)
            
            for i in range(len(kws)):
                sims[i] = len(set(kws[i]) & set(clean_text))
            
            if max(sims) == 0:
                classified['Miscellaneous'].append(response)
            else:
                temp_cat = classes[sims.index(max(sims))]
                classified[temp_cat].append(response)
        
    with open('practice.txt', 'w') as f:
        for key in classified:
            f.write(key.upper()+'\n')
            print(key)
            for val in classified[key]:
                f.write('\t\t'+val+'\n')
            f.write('\n\n')
    return classified
            
            
grab_data()
classify()
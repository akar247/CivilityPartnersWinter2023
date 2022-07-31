#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:


df = pd.read_excel(r"C:\Users\kelly\Downloads\Vaxart Survey 2021.xlsx", None)


# In[3]:


key = list(df)
data = {}
for i in range(len(df)):
    index = key[i]
    if df[index]["Unnamed: 2"].dtype == 'O':
        data[index] = df[index]["Unnamed: 2"].dropna()
        data[index] = [j for j in data[index] if not isinstance(j, int)]


# In[6]:


data


# In[ ]:







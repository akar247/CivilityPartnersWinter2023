import pandas as pd
import numpy as np


file = 'Vaxart Survey 2021.xlsx'


def read_excel_questions(fp):
    sheets = pd.ExcelFile(fp).sheet_names
    q_data = [(sheet, pd.read_excel(fp, sheet_name=sheet, header=2)) for sheet in sheets]
    return q_data

read_excel_questions(file)
q_data = read_excel_questions(file)
# print(q_data[:5], '1')

def mc_data(dfs):
    col_check = ['Answer Choices', 'Responses', 'Unnamed: 2']
    tholds = {'Negative': [], 'Neutral': [], 'Positive' : []}
    for i in range(len(dfs)):
        sheet = dfs[i][0]
        df = dfs[i][1]
        if list(df.columns) == col_check:
            neg = df['Responses'].iloc[:2].sum()
            pos = df['Responses'].iloc[2:4].sum()

            if neg >= .25:
                tholds['Negative'].append((sheet, neg))
            elif pos >= .90:
                tholds['Positive'].append((sheet, pos))
            else:
                tholds['Neutral'].append(sheet)
    
    return tholds

mc_tholds = mc_data(q_data)
# print(mc_tholds)

def write_txt(dict):
    with open('thresholds.txt', 'w') as f:
        f.write('Negatives:\n')
        for qt in dict['Negative']:
            f.write('\t' + str(qt) + '\n')
        f.write('\n\n')
        f.write('Positives:\n')
        for qt in dict['Positive']:
            f.write('\t' + str(qt) + '\n')

write_txt(mc_tholds)




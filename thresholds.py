import pandas as pd
import numpy as np


file = 'Vaxart Survey 2021.xlsx'
statements = {}


def read_excel_questions(fp):
    sheets = pd.ExcelFile(fp).sheet_names
    q_data = []
    
    for sheet in sheets:
        q_data.append((sheet, pd.read_excel(fp, sheet_name=sheet, header=2)))
        statements[sheet] = pd.read_excel(fp, sheet_name=sheet).iloc[0,0]
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
                tholds['Negative'].append((sheet, statements[sheet], round(neg, 2)))
            elif pos >= .90:
                tholds['Positive'].append((sheet, statements[sheet], round(pos, 2)))
            else:
                tholds['Neutral'].append(sheet)
    
    return tholds

mc_tholds = mc_data(q_data)
# print(mc_tholds)

def write_txt(dict):
    with open('thresholds.txt', 'w') as f:
        f.write('Negatives:\n')
        for qt in dict['Negative']:
            f.write('\t' + str(qt[0]) + '\n')
            f.write('\t\t' + str(qt[1]) + " ({:.0%}) ".format(qt[2]) + '\n\n')
        f.write('\n\n')
        f.write('Positives:\n')
        for qt in dict['Positive']:
            f.write('\t' + str(qt[0]) + '\n')
            f.write('\t\t' + str(qt[1]) + " ({:.0%}) ".format(qt[2]) + '\n\n')

write_txt(mc_tholds)




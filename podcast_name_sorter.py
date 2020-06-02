from pandas import DataFrame
from collections import defaultdict
import re
import numpy as np
import pandas as pd
import collections
import csv
#------------ Input and output files ----------------
filename = r'popular_podcasts.xlsx'
saved_name = 'most_popular.csv'
def findMostPopular(input_name,output_name):
    df = pd.read_excel(input_name)
    number_of_columns = len(df.columns)
    column_names = []
    for list_n in df:
        column_names.append(list_n)
#-------------- Separates columns into different lists and puts them together into one -----------------
    lists_of_names = []
    for n in range(number_of_columns):
        column = list(df[column_names[n]].str.lower().str.replace('(“|”|\xa0|[^\x00-\x7F]+)','', regex=True).dropna())
        lists_of_names.append(column)
#------------- Find Most Popular -------------------
    most_popular = {}
    for list_n in lists_of_names:
        for item in list_n:
            if item not in most_popular:
                most_popular.update({item:1})
            elif item in most_popular:
                most_popular[item] +=1
#------------ Create CSV file
    with open(output_name, 'w') as f:
        for key in most_popular.keys():
            f.write("%s, %s\n" % (key, most_popular[key]))
findMostPopular(filename,saved_name)
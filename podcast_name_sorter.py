from pandas import DataFrame
from collections import defaultdict
import re
import numpy as np
import pandas as pd
import collections
import csv

filename = r'popular_podcasts.xlsx'
df = pd.read_excel(filename)
#-------------- Gets the number of columns ----------------
number_of_columns = len(df.columns)

#-------------- Gets the name of each columnm --------------
column_names = []
for list_n in df:
    column_names.append(list_n)

#-------------- Separates columns into different lists and puts them together into one -----------------

lists_of_podcasts = []
for n in range(number_of_columns):
    column = list(df[column_names[n]].str.lower().str.replace('(“|”|\xa0|[^\x00-\x7F]+)','', regex=True).dropna())
    lists_of_podcasts.append(column)

#------------- Find Most Popular -------------------
most_popular = {}
for list_n in lists_of_podcasts:
    for item in list_n:
        if item not in most_popular:
            most_popular.update({item:1})
        elif item in most_popular:
            most_popular[item] +=1
#most_popular = {k: str(v).encode("utf-8") for k,v in most_popular.items()}
sorted_dict = collections.OrderedDict(most_popular)

print (most_popular)

with open('most_popular.csv', 'w') as f:
    for key in most_popular.keys():
        f.write("%s, %s\n" % (key, most_popular[key]))
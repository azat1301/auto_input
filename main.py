import csv
import re
import pandas as pd

#with open(r'C:\Users\Azat\Downloads\csv1.csv', 'r', encoding='utf-8') as f:
#    f_contents = f.read()

#pattern = re.compile(r'(\bJH|\bUM|\BJH)\d\d\d\d\d\d\d\d\d?')

f_contents = pd.read_csv(r'C:\Users\Azat\Downloads\csv1.csv')

df = pd.DataFrame(f_contents)

#matches = pattern.finditer(f_contents)
#for match in matches:
#     print(match)
print(df.loc[index])
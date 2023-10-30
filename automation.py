import pandas as pd
import duckdb
import string
from time import sleep
from tqdm.auto import tqdm
import re

data = (pd.read_csv('merged_gw.csv')
        .drop(columns = "kickoff_time"))

indicators = pd.read_json('queries.json')

with pd.ExcelWriter("FPL.xlsx", engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        
        for index, row in tqdm(indicators.iterrows(), total=indicators.shape[0]):

                starting_column = ''.join(re.findall(r'[A-Za-z]', row["range"].split(':')[0]))
                
                starting_row = int(''.join(re.findall(r'\d', row["range"].split(':')[0]))) - 1

                num = 0
                for c in starting_column:
                        if c in string.ascii_letters:
                                num = (num * 26 + 1 + ord(c) - ord('A'))
                num = num - 1
                result = (duckdb.sql(row["query"])
                          .to_df())
                result.to_excel(writer, index=False,
                                startcol=num,
                                startrow=starting_row,
                                header=False)
                title = row["indicator_number"]

                sleep(2)

                print(f'Calculated indicator {title}'.format(title = title))
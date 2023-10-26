import pandas as pd
import duckdb
import string

data = (pd.read_csv('merged_gw.csv')
        .drop(columns = "kickoff_time"))

indicators = pd.read_json('queries.json')

with pd.ExcelWriter("FPL.xlsx", engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        
        for index, row in indicators.iterrows():
                num = 0
                for c in row["startcol"]:
                        if c in string.ascii_letters:
                            num = (num * 26 + 1 + ord(c) - ord('A'))
                num = num - 1
                
                result = (duckdb.sql(row["query"])
                          .to_df())
                
                result.to_excel(writer, index=False,
                                startcol=num,
                                startrow=row["startrow"]-1,
                                header=False)
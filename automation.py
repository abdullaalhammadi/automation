import pandas as pd
import duckdb
import string
from time import sleep
from tqdm import tqdm

data = (pd.read_csv('merged_gw.csv')
        .drop(columns = "kickoff_time"))

indicators = pd.read_json('queries.json')

with pd.ExcelWriter("FPL.xlsx", engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:

        iterator = tqdm(indicators.iterrows(), total=indicators.shape[0])

        for index, row in iterator:
                num = 0
                for c in row["startcol"]:
                        if c in string.ascii_letters:
                                num = (num * 26 + 1 + ord(c) - ord('A'))
  
                result = (duckdb.sql(row["query"])
                          .to_df())
                result.to_excel(writer, index=False,
                                startcol=num-1,
                                startrow=row["startrow"]-1,
                                header=False)
                title = row["indicator_number"]

                sleep(0.5)

                iterator.set_description(f"Indicator {title} calculated successfully"
                                         .format(title = row["indicator_number"]))
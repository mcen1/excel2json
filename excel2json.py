#!/bin/env python3
import pandas as pd
import json
from pprint import pprint
myfile="thing.xlsx"
 
# config
pd.set_option('display.max_rows', 999999)
pd.set_option('display.max_columns', 999999)
 
 
if __name__ == "__main__":
  xl = pd.ExcelFile(myfile)
  mydata={}
  ignoresheets=[‘pivot‘]
  for sheet in xl.sheet_names:
    if sheet in ignoresheets:
      continue
    if sheet not in mydata:
      mydata[sheet]={"data":[]}
    myexcel=pd.read_excel(open(myfile, 'rb'),sheet_name=sheet)
    df = myexcel.where(pd.notnull(myexcel), None)
    headers=df.columns.values.tolist()
    for row in df.itertuples():
      towrite={}
      for idx,item in enumerate(row):
        if idx==0:
          # skip first item
          continue
        if idx-1<len(headers):
          towrite[headers[idx-1]]=str(item)
      mydata[sheet]["data"].append(towrite)
  jsondict=json.loads(json.dumps(mydata))
  print(json.dumps(mydata))

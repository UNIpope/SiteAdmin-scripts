import pandas as pd
import numpy as np
import openpyxl as xl
from openpyxl.styles import PatternFill
import sys, re, os, collections

from pprint import pprint

names = []
frames = []
path = r"C:\Users\jduggan\Documents\A hole year ivana\\"
for i in os.listdir(path):
    #generate headder
    dirf = path + i

    with open(dirf, "r") as f:
        lines = [line for line in f][1:4]


    line = lines[0].split(",")
    month = line[2]
    print(month)

    days = lines[1].split(",")
    nums = lines[2].split(",")
    print(len(days),len(nums))

    head = ["Name"] 
    for i in  range(3, len(days)-1):
        head.append(days[i] + ":" + nums[i] + ":" + month)

    # read in rotations.csv + cleanup
    df = pd.read_csv(dirf, encoding="ISO-8859-1", skiprows=[1,2,3], skipfooter=9, engine='python')
    df = df.fillna("IN")
    df = df.apply(lambda x: x.astype(str).str.lower().str.strip())

    # drop sapno and role
    df = df.drop(df.columns[0], axis=1)
    df = df.drop(df.columns[0], axis=1)

    # drop total days
    df = df.drop(df.columns[-1], axis=1)
    #df = df.drop(df.columns[-1], axis=1)


    # assign headder to df
    df.columns = head

    # fix spelling mistake
    df['Name'] = df['Name'].apply(lambda x: "alice wells" if x == "aliice wells" else x)

    #extent list to lists
    names.extend(df['Name'].tolist())

    frames.append(df)

    print(df.columns)


# merge month dfs into year df
dfr = frames[0]
for i in frames[1:]:
    dfr = pd.merge(dfr, i , on=["Name"], how="outer" )

dfr.to_csv(path + r"\merge.csv",index=False)


#list to dict count to a-z sort
occurrences = dict(collections.Counter(names))
occurrences = list(collections.OrderedDict(sorted(occurrences.items())))
print(len(occurrences))



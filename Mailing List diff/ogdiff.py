import pandas as pd
import numpy as np

# clean up text
def strip(text):
    return text.strip().lower()

# import csvs
outdf = pd.read_csv("siteadmin\outlookdiff\outlook_DCBS_buss_unit.csv", sep='\t', names=["name", "email"], converters = {"name":strip, "email":strip})
excdf = pd.read_csv("siteadmin\outlookdiff\excel_DCBS_buss_unit.csv", sep='\t', names=["name", "email"], converters = {"name":strip, "email":strip})

# sort csvs
outdf = outdf.sort_values('name')
excdf = excdf.sort_values('name')

print(outdf.head)
print(excdf.head)

# outer merge dfs and get rows only in left df
merged = excdf.merge(outdf, indicator=True, how='outer', on="email")
oL = merged[merged['_merge'] == 'left_only']
oR = merged[merged['_merge'] == 'right_only']

# output and count rows
print(oL)
print(oL.shape[0] )

print(oR)
print(oR.shape[0] )
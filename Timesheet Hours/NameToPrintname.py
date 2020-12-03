import pandas as pd
import numpy as np
import openpyxl as xl
from openpyxl.styles import PatternFill
import sys, re

# read in names
with open(r"timesheet\fill\names.txt") as fname:
    content = fname.readlines()
    print(content)

for i in content:
    name = i.strip().split(" ")
    name = ", ".join(name[::-1]).strip()
    print(name)
    print()
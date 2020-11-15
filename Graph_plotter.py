import openpyxl
from pathlib import Path
import matplotlib.pyplot as plt

Charpy_results = [[],[]] #[[temp], [joule]]

# Setting the path to the xlsx file:
xlsx_file =  'Charpytest_resultater.xlsx'

book = openpyxl.load_workbook(xlsx_file)

sheet = book.active

cells = sheet['A2': 'B26']

for c1, c2 in cells:
    #print("{0:8} {1:8}".format(c1.value, c2.value))
    Charpy_results[0].append(c1.value)
    Charpy_results[1].append(c2.value)


plt.plot(Charpy_results[0], Charpy_results[1])
plt.xlabel("Temperatur")
plt.ylabel("Joule")
plt.show()

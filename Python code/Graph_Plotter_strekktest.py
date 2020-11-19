import openpyxl
import matplotlib.pyplot as plt


class Strekktest:
    def __init__(self, start, stopp, name_excel_file, sheet_name):
        self.start = start
        self.stopp = stopp
        self.name_excel_file = name_excel_file
        self.sheet_name = sheet_name

    def get_data(self):
        xlsx_file =  self.name_excel_file

        book = openpyxl.load_workbook(xlsx_file)
        sheet = book[self.sheet_name]

        cells = sheet[self.start:self.stopp]
        data_results = [[] for i in range(len(cells[0]))]
        for i in cells:
            for j in range(len(i)):
                data_results[j].append(i[j].value)

        return data_results

    def plot_data(self, x, y, x_label, y_label, show_grid = True, show_graph = False):
        #tar listen fra get_data(), og plotter grafen med punktene gitt av posisjonene
        # (get_data()[x], get_data()[y])
        list_data = self.get_data()
        plt.plot(list_data[x], list_data[y])
        plt.xlabel(x_label)
        plt.ylabel(y_label)

        if show_grid == True:
            plt.grid()
        if show_graph == True:
            plt.show()






a = Strekktest('A7', 'F576', 'strekking nov 2020_1.xlsx', '1b')
a.plot_data(0,1,'x-akse','y-akse')
plt.show()
#print(spec_1[0])
#print(a.get_data()[0])
#print(spec_1[0][0].value)
"""
#[[Load(kN)], [Stress(MPa)],[Displacement (mm)], [Strain (mm/mm)], [Strain (%)] ]
Strekktest_spec1 = [[],[],[],[],[]]
Strekktest_spec15 = [[],[],[],[],[]]

# Setting the path to the xlsx file:
xlsx_file =  'strekking nov 2020_1.xlsx'

book = openpyxl.load_workbook(xlsx_file)

sheet_1a = book.get_sheet_by_name("1b")

cells = sheet_1a['A7': 'F1238']



for c1, c2 in cells:
    #print("{0:8} {1:8}".format(c1.value, c2.value))
    Charpy_results[0].append(c1.value)
    Charpy_results[1].append(c2.value)

cells_varmebehandlet = sheet['D2': "E12"]
for c1, c2 in cells_varmebehandlet:
    print("{0:8} {1:8}".format(c1.value, c2.value))
    Charpy_results_varmebehandlet[0].append(c1.value)
    Charpy_results_varmebehandlet[1].append(c2.value)


plt.plot(Charpy_results[0], Charpy_results[1],"bo", \
label = 'Ikke varmebehandlet')
plt.plot(Charpy_results_varmebehandlet[0], Charpy_results_varmebehandlet[1], \
'ro', label = 'varmebehandlet')
plt.xlabel("Temperatur")
plt.ylabel("Joule")
plt.grid()
plt.legend()
plt.show()
"""

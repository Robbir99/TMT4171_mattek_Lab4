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

    def plot_data(self, x, y, x_label, y_label, label_name, show_grid = True, show_graph = False):
        #tar listen fra get_data(), og plotter grafen med punktene gitt av posisjonene
        # (get_data()[x], get_data()[y])
        list_data = self.get_data()
        plt.plot(list_data[x], list_data[y], label = label_name)
        plt.legend()
        plt.xlabel(x_label)
        plt.ylabel(y_label)

        if show_grid == True:
            plt.grid()
        if show_graph == True:
            plt.show()



specimen_nr1 = Strekktest('A7', 'F576', 'strekking nov 2020_1.xlsx', '1b')
specimen_nr4 = Strekktest('A585', 'F1274', 'strekking nov 2020_1.xlsx', '1b')
specimen_nr5 = Strekktest('A1284', 'F1772', 'strekking nov 2020_1.xlsx', '1b')

specimen_nr1.plot_data(4,2, 'Tøyning [mm]','Spenning [MPa]', 'specimen_nr1')
specimen_nr4.plot_data(4,2,'Tøyning [mm]','Spenning [MPa]', 'specimen_nr4')
specimen_nr5.plot_data(4,2,'Tøyning [mm]','Spenning [MPa]', 'specimen_nr5')
plt.show()

a = specimen_nr1.get_data()
print(a[4][10])
print(a[4][30])
print(a[2][10])
print(a[2][30])
dy = (a[2][30] - a[2][10])*10
dx = (a[4][30] - a[4][10])*10
print(dy/dx)

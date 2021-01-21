from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook

column_cell_1 = ""
column_cell_2 = ""
column_stock_1 = ""
column_stock_2 = ""
rand_fisier_1 = ""
rand_fisier_2 = ""
path1 = ""
path2 = ""


def get_file_path1():
    # Open and return file path
    global path1
    file_path1 = filedialog.askopenfilename(title="Alege un fișier excel", filetypes=[("Excel files", ".xlsx .xls")])
    path1 = file_path1


def get_file_path2():
    # Open and return file path
    global path2
    file_path2 = filedialog.askopenfilename(title="Alege un fișier excel", filetypes=[("Excel files", ".xlsx .xls")])
    path2 = file_path2


def get_columns():
    global column_cell_1, column_cell_2, column_stock_1, column_stock_2, rand_fisier_1, rand_fisier_2
    column_cell_1 = column_1.get()
    column_cell_2 = column_2.get()
    column_stock_1 = e_column_1.get()
    column_stock_2 = e_column_2.get()
    rand_fisier_1 = e_rand_1.get()
    rand_fisier_2 = e_rand_2.get()


window = Tk()
window.title('VLookup by Raul Glodean')
window.geometry("400x400")

Label(window, text="Număr coloană unică din fișierul 1:\t").grid(row=2)
Label(window, text="Număr coloană unică din fișierul 2:\t").grid(row=3)
Label(window, text="Număr coloană de căutat din fișierul 1:\t").grid(row=4)
Label(window, text="Număr coloană de căutat din fișierul 2:\t").grid(row=5)
Label(window, text="Primul rând din fișierul 1:\t").grid(row=6)
Label(window, text="Primul rând din fișierul 2:\t").grid(row=7)

column_1 = Entry(window)
column_1.grid(row=2, column=1)
column_2 = Entry(window)
column_2.grid(row=3, column=1)
e_column_1 = Entry(window)
e_column_1.grid(row=4, column=1)
e_column_2 = Entry(window)
e_column_2.grid(row=5, column=1)
e_rand_1 = Entry(window)
e_rand_1.grid(row=6, column=1)
e_rand_2 = Entry(window)
e_rand_2.grid(row=7, column=1)

button_get_columns = Button(window, text="TRIMITE", command=get_columns)
button_get_columns.grid(row=9, column=1)

button_get_path1 = Button(window, text="Alege fișier 1", command=get_file_path1)
button_get_path1.grid(row=0)
button_get_path2 = Button(window, text="Alege fișier 2", command=get_file_path2)
button_get_path2.grid(row=1)

window.mainloop()

column_cell_1 = int(column_cell_1)
column_cell_2 = int(column_cell_2)
column_stock_1 = int(column_stock_1)
column_stock_2 = int(column_stock_2)
rand_fisier_1 = int(rand_fisier_1)
rand_fisier_2 = int(rand_fisier_2)

wb_obj1 = load_workbook(path1, data_only=True)
sheet_obj1 = wb_obj1.active
wb_obj2 = load_workbook(path2, data_only=True)
sheet_obj2 = wb_obj2.active

for i in range(rand_fisier_1, sheet_obj1.max_row+1):
    cell_obj1 = sheet_obj1.cell(row=i, column=column_cell_1)
    if cell_obj1.value is None or cell_obj1.value == 0:
        continue
    if sheet_obj1.row_dimensions[i].hidden:
        continue
    for ii in range(rand_fisier_2, sheet_obj2.max_row+1):
        cell_obj2 = sheet_obj2.cell(row=ii, column=column_cell_2)
        if cell_obj1.value == cell_obj2.value:
            cell_obj_final = sheet_obj2.cell(row=ii, column=column_stock_2)
            cell_obj1_de_luat = sheet_obj1.cell(row=i, column=column_stock_1)
            cell_obj_final.value = cell_obj1_de_luat.value


wb_obj2.save(path2)

window_success = Tk()
window_success.title('VLookup by Raul Glodean')
window_success.geometry("400x400")

Label(window_success, text="PROCES FINALIZAT CU SUCCES!").grid(row=2)

window_success.mainloop()

import csv
import sqlite3
import tkinter
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import xlrd as xlrd
from pathlib import Path


def create_table(dbname):
    con = sqlite3.connect(dbname)
    cur = con.cursor()
    cur.execute('pragma encoding=UTF16')
    query = "CREATE TABLE IF NOT EXISTS table1(A TEXT,B TEXT,C TEXT,D TEXT,E TEXT,F TEXT, G TEXT,H TEXT,I TEXT,J TEXT)"
    cur.execute(query)
    return [con, cur]   # return the created connection


def data_entry(A, B, C, D, E, F, G, H, I, J, con, cur):

    query = 'INSERT INTO table1 values("'+A+'","'+B+'","'+C+'","'+D+'","'+E+'","'+F+'","'+G+'","'+H+'","'+I+'","'+J+'")'

    cur.execute(query)
    con.commit()



def load_file(file_path):
    wb = xlrd.open_workbook(file_path)
    sch = wb.encoding
    sheet = wb.sheet_by_index(0)
    return sheet


def verify_data(sheet, output_file, con, cur):
    rows = sheet.nrows
    columns = sheet.ncols
    successful = 0
    failed = 0

    with open(output_file, 'w',  newline='', encoding="utf-8")as myFile:
        writer = csv.writer(myFile)
        for i in range(1, rows):
            is_empty_cell = 0
            for j in range(0, 10):
                cell = sheet.cell_value(i, j)
                if cell == '':
                    is_empty_cell = 1

            A = str(sheet.cell_value(i, 0))
            B = str(sheet.cell_value(i, 1))
            C = str(sheet.cell_value(i, 2))
            D = str(sheet.cell_value(i, 3))
            E = str(sheet.cell_value(i, 4))
            F = str(sheet.cell_value(i, 5))
            G = str(sheet.cell_value(i, 6))
            H = str(sheet.cell_value(i, 7))
            if H == '1':
                H = 'TRUE'
            else:
                H = 'FALSE'

            I = str(sheet.cell_value(i, 8))
            if I == '1':
                I = 'TRUE'
            else:
                I = 'FALSE'
            J = str(sheet.cell_value(i, 9))
            if is_empty_cell == 0:
                data_entry(str(A), str(B), str(C), str(D), str(E), str(F), str(G), str(H), str(I), str(J), con, cur)
                successful += 1
            else:
                writer.writerow([A,B,C,D,E,F,G,H,I,J])
                failed += 1

    myFile.close()
    return [rows, successful, failed]


root = tkinter.Tk()
root.withdraw()
messagebox.showinfo("File Browser", "Select excel file. ")
file_path = askopenfilename()
input_filename = Path(file_path).name
input_filename = input_filename.replace(".xlsx","")

output_file_name = input_filename+"-bad.csv"
db_name = input_filename+".db"
sheet = load_file(file_path)

con, cur = create_table(db_name)
result = verify_data(sheet, output_file_name, con, cur)

cur.close()
con.close()


log_file_name = input_filename+".log"
f = open(log_file_name, "w+")
f.write('----------Result Summary----------'+"\n")
f.write('Number of record list : '+str(result[0]-1)+"\n")
f.write('Number of successful records : '+str(result[1])+"\n")
f.write('Number of failed records : '+str(result[2])+"\n")
f.close()







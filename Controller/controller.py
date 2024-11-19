import xlwings as xw
import numpy as np
from Model.flowrate import caudal
import os

def create_excel_file(filepath):
    if not os.path.exists(filepath):
        wb = xw.Book()
        wb.sheets[0].name = "Summary"
        wb.save(filepath)
        wb.close()

def main():

    filepath = r"Examen.xlsm"
    create_excel_file(filepath)

    Qmax = 250
    Pr = 2400
    Pwf = np.array([2400, 2000, 1500, 1000, 500, 0])

    Qo = caudal(Qmax, Pwf, Pr)

    wb = xw.Book(filepath)
    sheet = wb.sheets["Summary"]


    sheet["A1"].value = "Pwf (psi)"
    sheet["B1"].value = "Qo (bpd)"
    sheet["C1"].value = "Pr (psi)"
    sheet["D1"].value = "Qomax (bdp)"
    sheet["A2"].options(transpose=True).value = Pwf
    sheet["B2"].options(transpose=True).value = Qo
    sheet["C2"].options(transpose=True).value = Pr
    sheet["D2"].options(transpose=True).value = Qmax

    wb.save()
    wb.close()

if __name__ == "__main__":
    main()


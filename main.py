import xlrd
from xlrd.biffh import XL_CELL_EMPTY, XL_CELL_TEXT
from pathlib import Path
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import re


class Patient:
    def __init__(self, data):
        self.Patient = data[0].value
        self.Platform = data[1].value
        self.Cytogenetic_aberration = data[2].value
        random_not = data[3].value
        numbers = re.findall("\([0-9\?_]*\)", random_not)
        self.chrom = list()
        for n in numbers:
            n = n.replace("(", "").replace(")", "")
            n = n.split("_")
            if n[0] == "?":
                self.chrom.append(int(n[1]))
                self.chrom.append(int(n[1]))
            elif n[1] == "?":
                self.chrom.append(int(n[0]))
                self.chrom.append(int(n[0]))
            else:
                p = (int(n[0]) + int(n[1])) // 2
                self.chrom.append(p)
                self.chrom.append(p)
        self.mb = self.chrom[2] - self.chrom[1]
        self.length = data[4].value
        print(self.mb, self.length)
        self.OMIM_genes = data[5].value
        self.OMIM_disease_related_genes = data[6].value
        self.Origin = data[7].value


if __name__ == '__main__':
    patients = list()
    data_path = Path().cwd() / "data" / "testdata.xlsx"
    workbook = xlrd.open_workbook(str(data_path))
    geno = workbook.sheets()[0]
    rows = [row for row in geno.get_rows()]
    columns = [c.value for c in rows[0]]
    x = []
    for row in rows[1:]:
        if row[0].ctype is XL_CELL_TEXT and row[1].ctype is XL_CELL_EMPTY:
            patients.append("")
            x.append([])
        elif row[0].ctype is XL_CELL_TEXT and row[1].ctype is XL_CELL_TEXT:
            p = Patient(row)
            patients.append(p.Patient)
            x.append(p.chrom)

    x.reverse()
    patients.reverse()
    df = pd.DataFrame(x, index=patients)

    print(df)
    df.T.boxplot(vert=False)
    plt.subplots_adjust(left=0.25)
    plt.gca().xaxis.set_major_formatter(plt.FormatStrFormatter('%d'))
    plt.show()

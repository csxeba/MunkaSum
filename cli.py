import os

from MunkaSum.src.logic import *


def tk_get_path():
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename

    tk = Tk()
    tk.withdraw()
    inpath = askopenfilename()
    tk.destroy()

    return inpath


OUTPATH = "E:/tmp/output.xlsx"
path = tk_get_path()
matrix, header = parse_xl(path)

dump_to_xl(matrix, header, outpath=OUTPATH)
print("Dumped output file to", OUTPATH)
os.startfile(OUTPATH)

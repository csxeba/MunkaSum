from src.logic import *


def tk_get_path():
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename

    tk = Tk()
    tk.withdraw()
    inpath = askopenfilename()
    tk.destroy()

    return inpath


OUTFLTYPE = "xlsx"
path = tk_get_path()
matrix, header = parse_xl(path)

if OUTFLTYPE == "csv":
    dump_to_csv(matrix, header, outroot="E:/tmp/")
else:
    dump_to_xl(matrix, header, outpath="E:/tmp/output.xlsx")

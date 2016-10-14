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


path = tk_get_path()
outpath = "\\".join(path.split("\\")[:-1]) + "\\"
outpath += datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

matrix, header = parse_xl(path)

dump_to_xl(matrix, header, outpath=outpath)
print("Dumped output file to", outpath)
os.startfile(outpath)

import os
from tkinter import *
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showwarning, showinfo

from src.logic import *


class App(Tk):
    def __init__(self):
        Tk.__init__(self)

        self.geometry("400x300")
        self.title("Adatösszesítő alkalmazás")
        self.matrix = None
        self.header = None
        self.state = "input"
        self.chosenpath = StringVar(value="Nincs bemeneti fájl kiválasztva!")

        topbw = 5

        header = Frame(self, bd=topbw, relief=RAISED)
        Label(header, text="Érkezési/távozási adat összesítő",
              font=("Times New Roman", 16)
              ).pack(fill=BOTH)
        header.pack(fill=BOTH)

        midfix = Frame(self, bd=topbw, relief=RAISED)
        Label(midfix, textvariable=self.chosenpath, bg="white").pack(fill=BOTH)
        midfix.pack(fill=BOTH)

        footer = Frame(self, bd=topbw, relief=RAISED)
        self.browse_button = Button(footer, text="Tallózás", command=self.getpath)
        self.browse_button.pack(fill=BOTH)
        footer.pack(fill=BOTH)
        self.process_button = Button(footer, text="Feldolgozás", state=DISABLED,
                                     width=16, command=self.process)
        self.process_button.pack(fill=BOTH)

    def getpath(self):
        if self.state == "input":
            chosen = askopenfilename()
        else:
            chosen = asksaveasfilename()
            if chosen[-5:] != ".xlsx":
                chosen += ".xlsx"
        if chosen[-5:].lower() != ".xlsx":
            showwarning("Hiba!", "A választott fájl nem használható! (MS Excel XLSX fájlt kérek!)")
            self.process_button.configure(state=DISABLED)
            return
        self.chosenpath.set(chosen)
        self.process_button.configure(state=ACTIVE)

    def process(self):
        if self.chosenpath.get()[:5].lower() == "nincs":
            showwarning("Hiba!", "Nincs bemeneti fájl kiválasztva!")
            return
        try:
            self.matrix, self.header = parse_xl(self.chosenpath.get())
        except RuntimeError:
            showwarning("Hiba!", "Érvénytelen fájl!")
            self.chosenpath.set("Nincs bemeneti bemeneti kiválasztva!")
            self.process_button.configure(state=DISABLED)
        else:
            showinfo("Siker!", "Sikeres adatfeldolgozás! Válassz kimeneti mappát és fájlnevet!")
            self.chosenpath.set("Nincs kimeneti fájl kiválasztva!")
            self.process_button.configure(state=DISABLED, text="Mentés másként...", command=self.saveas)
            self.browse_button.configure(command=self.getpath)
            self.state = "output"

    def saveas(self):
        if self.chosenpath.get()[:5].lower() == "nincs":
            showwarning("Hiba!", "Nincs kimeneti fájl kiválasztva!")
            return
        if self.matrix is None or self.header is None:
            showwarning("Hiba!", "Nincs feldolgozott adat a memóriában!")
        dump_to_xl(self.matrix, self.header, self.chosenpath.get())
        showinfo("Siker!", "Sikeres mentés! Mentett fájl megnyitása...")
        os.startfile(self.chosenpath.get())
        self.browse_button.configure(state=DISABLED)
        self.process_button.configure(text="Bezárás", command=self.destroy)

if __name__ == '__main__':
    root = App()
    root.mainloop()

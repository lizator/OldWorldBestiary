import tkinter as tk
from tkinter import simpledialog

class MyDialog(simpledialog.askstring):
    def body(self, master):
        self.geometry("400x600")
        tk.Label(master, text="Enter your search string text:").grid(row=0)

        self.e1 = tk.Entry(master)
        self.e1.grid(row=0, column=1)
        return self.e1 # initial focus

    def apply(self):
        first = self.e1.get()
        self.result = first


root = tk.Tk()
root.withdraw()
test = MyDialog(root, "testing")
print (test.result)

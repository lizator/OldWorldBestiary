from tkinter import *
from tkinter import simpledialog

prompts = ["name", "age", "height", "wheight"]

root = Tk()

for p in prompts:
    answer = simpledialog.askstring(p, root)
    print(answer)

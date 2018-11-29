import tkinter as tk

def create_window():
    window = tk.Toplevel(root)
    window.minsize(400,120)
    window.title = "hello"

root = tk.Tk()
b = tk.Button(root, text="Create new window", command=create_window)
b.pack()

root.minsize(400,120)
root.mainloop()

frame = Frame(root, height=50, width=200, color="c3c3c3")
Frame.pack_propagate(0) # don't shrink
Frame.pack()

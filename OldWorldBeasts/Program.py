from openpyxl import Workbook
from openpyxl import load_workbook
from tkinter import *
from tkinter import filedialog
from tkinter import simpledialog
from easygui import enterbox
from tkinter import Toplevel
import re

class AutocompleteEntry(Entry):
    def __init__(self, lista, *args, **kwargs):

        Entry.__init__(self, *args, **kwargs)
        self.lista = entryList
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()

        self.var.trace('w', self.changed)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)

        self.lb_up = False

    def changed(self, name, index, mode):

        if self.var.get() == '':
            self.lb.destroy()
            self.lb_up = False
        else:
            words = self.comparison()
            if words:
                if not self.lb_up:
                    self.lb = Listbox()
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Right>", self.selection)
                    self.lb.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height())
                    self.lb_up = True

                self.lb.delete(0, END)
                for w in words:
                    self.lb.insert(END,w)
            else:
                if self.lb_up:
                    self.lb.destroy()
                    self.lb_up = False

    def selection(self, event):

        if self.lb_up:
            self.var.set(self.lb.get(ACTIVE))
            self.lb.destroy()
            self.lb_up = False
            self.icursor(END)

    def up(self, event):

        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != '0':
                self.lb.selection_clear(first=index)
                index = str(int(index)-1)
                self.lb.selection_set(first=index)
                self.lb.activate(index)

    def down(self, event):

        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != END:
                self.lb.selection_clear(first=index)
                index = str(int(index)+1)
                self.lb.selection_set(first=index)
                self.lb.activate(index)

    def comparison(self):
        pattern = re.compile('.*' + self.var.get() + '.*')
        return [w for w in self.lista if re.match(pattern, w)]

prompts = ["Navn (det der søges efter)", "Sidenummer (i WFRP Old World Beastiary)", "Weapon Skill (WS)", "Balistic Skill (BS)", "Strength (S)", "Toughness (T)",
"Agility (Ag)", "Intelligence (In)", "Will Power (WP)", "Fellowship (Fel)", "Attacks (A)", "Wounds (W)", "Movement (M)", "Magic (Mag)",
"Skills (separeret med et komma)", "Talents (separeret med et komma)", "Armour (None, Light, Medium, Heavy)", "Weapons (separeret med et komma)",
"Trappings (separeret med et komma)", "Special Rules (separeret med et komma)"]
entryCount = 0
nameData = ""
directory = "/"
entryDic = {}

def updateDirectory(filenameList):
    directory = ""
    filenameList[len(filenameList) - 1] = None
    for x in filenameList:
        if str(x) != "None":
            directory += str(x) + '/'

def selectData(): # FIXME: check for same name = err
    global entryCount; global nameData; global directory; global entryDic
    directory = "C:/Users/Frede/Desktop/OldWorldBeasts"
    entryCount = 0
    #Choosing Opsum file
    root.filename =  filedialog.askopenfilename(initialdir = directory, title = "Select XL Database", filetypes = (("Excel files","*.xlsx"),("All files","*.*")))
    wbData = load_workbook(root.filename)
    nameDataList = root.filename.split("/")
    nameData = nameDataList[len(nameDataList) - 1]
    updateDirectory(nameDataList)
    wsData = wbData.active

    #updating lable
    textData = "Dokument valgt: " + nameData
    vData.set(textData)

    #updating entryCount
    x = 1
    while wsData["A" + str(x)].value != None:
        entryCount += 1
        entryDic[wsData["A" + str(x)].value] = entryCount
        x += 1
    textEntries = "antal entries: " + str(entryCount)
    vEntries.set(textEntries)

def newEntry():
    global prompts; global entryCount; global directory; global nameData; global entryDic
    wbData = load_workbook(root.filename)
    wsData = wbData.active
    entryCount += 1
    promptCount = 0
    start = "A" + str(entryCount)
    slut = "T" + str(entryCount)
    for row in wsData[start:slut]:
        for cell in row:
            answer = enterbox(prompts[promptCount])
            print(answer)
            cell.value = answer.title()
            promptCount += 1
    entryDic[wsData[start].value] = entryCount
    wbData.save(root.filename)
    textEntries = "antal entries: " + str(entryCount)
    vEntries.set(textEntries)

def insertStatFrame(frame, row, column, color):
    name = Frame(frame, height=28, width=28)
    name.pack_propagate(0)
    name.grid(row=row, column=column, padx=(1, 1), pady=(1, 1))
    name.configure(bg=color)
    return name

def openByName():
    global entryDic
    #lineNR = int(entryDic[name])
    window = Toplevel(root)
    #window.title(name)
    window.minsize(450,170)
    window.maxsize(450,170)
    #window.configure(bg="red")
    nameFrame = Frame(window, height=46, width=210) #30x30 per box
    nameFrame.pack_propagate(0) # don't shrink
    nameFrame.pack(side=TOP)
    nameFrame.configure(bg="#040030")

    leftFrame = Frame(window, height=170-46, width=210) #30x30 per box
    leftFrame.pack_propagate(0) # don't shrink
    leftFrame.pack(side=BOTTOM)
    leftFrame.configure(bg="#1C1C1C")

    wsTop = insertStatFrame(leftFrame, 0, 0, "#6E6E6E")
    bsTop = insertStatFrame(leftFrame, 0, 1, "#6E6E6E")
    sTop = insertStatFrame(leftFrame, 0, 2, "#6E6E6E")
    tTop = insertStatFrame(leftFrame, 0, 3, "#6E6E6E")
    agTop = insertStatFrame(leftFrame, 0, 4, "#6E6E6E")
    intTop = insertStatFrame(leftFrame, 0, 5, "#6E6E6E")
    wpTop = insertStatFrame(leftFrame, 0, 6, "#6E6E6E")
    wsBot = insertStatFrame(leftFrame, 1, 0, "#A4A4A4")
    bsBot = insertStatFrame(leftFrame, 1, 1, "#A4A4A4")
    sBot = insertStatFrame(leftFrame, 1, 2, "#A4A4A4")
    tBot = insertStatFrame(leftFrame, 1, 3, "#A4A4A4")
    agBot = insertStatFrame(leftFrame, 1, 4, "#A4A4A4")
    intBot = insertStatFrame(leftFrame, 1, 5, "#A4A4A4")
    wpBot = insertStatFrame(leftFrame, 1, 6, "#A4A4A4")

    mellemStatFrame0 = Frame(leftFrame, height=4, width=30)
    mellemStatFrame1 = Frame(leftFrame, height=4, width=30)
    mellemStatFrame2 = Frame(leftFrame, height=4, width=30)
    mellemStatFrame3 = Frame(leftFrame, height=4, width=30)
    mellemStatFrame4 = Frame(leftFrame, height=4, width=30)
    mellemStatFrame5 = Frame(leftFrame, height=4, width=30)
    mellemStatFrame6 = Frame(leftFrame, height=4, width=30)
    mellemStatFrame0.grid(row=2, column=0);
    mellemStatFrame1.grid(row=2, column=1);
    mellemStatFrame2.grid(row=2, column=2);
    mellemStatFrame3.grid(row=2, column=3);
    mellemStatFrame4.grid(row=2, column=4);
    mellemStatFrame5.grid(row=2, column=5);
    mellemStatFrame6.grid(row=2, column=6);
    mellemStatFrame0.configure(bg="#383a39")
    mellemStatFrame1.configure(bg="#383a39")
    mellemStatFrame2.configure(bg="#383a39")
    mellemStatFrame3.configure(bg="#383a39")
    mellemStatFrame4.configure(bg="#383a39")
    mellemStatFrame5.configure(bg="#383a39")
    mellemStatFrame6.configure(bg="#383a39")

    felTop = insertStatFrame(leftFrame, 3, 0, "#6E6E6E")
    aTop = insertStatFrame(leftFrame, 3, 1, "#6E6E6E")
    wTop = insertStatFrame(leftFrame, 3, 2, "#6E6E6E")
    sbTop = insertStatFrame(leftFrame, 3, 3, "#6E6E6E")
    tbTop = insertStatFrame(leftFrame, 3, 4, "#6E6E6E")
    mTop = insertStatFrame(leftFrame, 3, 5, "#6E6E6E")
    magTop = insertStatFrame(leftFrame, 3, 6, "#6E6E6E")
    felBot = insertStatFrame(leftFrame, 4, 0, "#A4A4A4")
    aBot = insertStatFrame(leftFrame, 4, 1, "#A4A4A4")
    wBot = insertStatFrame(leftFrame, 4, 2, "#A4A4A4")
    sbBot = insertStatFrame(leftFrame, 4, 3, "#A4A4A4")
    tbBot = insertStatFrame(leftFrame, 4, 4, "#A4A4A4")
    mBot = insertStatFrame(leftFrame, 4, 5, "#A4A4A4")
    magBot = insertStatFrame(leftFrame, 4, 6, "#A4A4A4")

    insertStatFrame(bsTop, leftFrame, 0, 1)

    rightFrame = Frame(window, height=170, width=240)
    rightFrame.pack(side=RIGHT)
    rightFrame.pack_propagate(0) # don't shrink
    rightFrame.configure(bg="#383a39")

#UI

root = Tk()
root.title("New Worlds Beastiary")

def insertMellemFrame(height, place):
    mellemFrame = Frame(place, height=height, width=400)
    mellemFrame.pack_propagate(0) # don't shrink
    mellemFrame.pack()

insertMellemFrame(30, root)

dataFrame = Frame(root, height=30, width=400)
dataFrame.pack_propagate(0) # don't shrink
dataFrame.pack()

vData = StringVar()
textData = "Dokument valgt: " + nameData
vData.set(textData)
datal = Label(dataFrame, textvariable=vData)

datal.pack(side=TOP)

insertMellemFrame(15, root)

entryFrame = Frame(root, height=80, width=400)
entryFrame.pack_propagate(0) # don't shrink
entryFrame.pack()

vEntries = StringVar()
textEntries = "antal entries: " + str(entryCount)
vEntries.set(textEntries)
entryl1 = Label(entryFrame, text="For future reference n stuff :P 2")
entryl2 = Label(entryFrame, textvariable=vEntries)
entryb = Button(entryFrame, text="New Entry", command=newEntry, height=1, width=20)

entryl1.pack(side=TOP)
entryb.pack(side=LEFT)
entryl2.pack(side=RIGHT)

insertMellemFrame(30, root)

autoFrame = Frame(root, height=225, width=400)
autoFrame.pack_propagate(0) # don't shrink
autoFrame.pack()




insertMellemFrame(30, root) # FIXME:

testb = Button(root, text="test", command=openByName, height=1, width=20)
testb.pack()

root.minsize(440,360)

selectData()

root.mainloop()

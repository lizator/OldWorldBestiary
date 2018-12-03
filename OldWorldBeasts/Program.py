from openpyxl import Workbook
from openpyxl import load_workbook
from tkinter import *
from tkinter import filedialog
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
entryDic = {}

def selectData():
    global entryCount; global nameData; global entryDic
    entryCount = 0
    #Choosing Opsum file
    try:
        load_workbook("data.xlsx")
    except:
        wbData = Workbook()
        wbData.save("data.xlsx")
    wsData = wbData.active

    #updating entryCount
    x = 1
    while wsData["A" + str(x)].value != None:
        entryCount += 1
        entryDic[wsData["A" + str(x)].value] = entryCount
        x += 1
    textEntries = "antal entries: " + str(entryCount)
    vEntries.set(textEntries)

def newEntry(): # FIXME: check for same name = err
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

def statLabel(frame, txt, bg):
    name = Label(frame, text=txt)
    name.place(x=14, y=14, anchor="center")
    name.configure(bg=bg)
    return name

def statText(frame, txt):
    name = Text(frame)
    name.tag_configure("center", justify='center')
    name.insert("1.0", txt)
    name.tag_add("center", "1.0", "end")
    name.pack()
    name.configure(font=("Helvetica", 9), bg="#A4A4A4")
    return name

def openByName():
    global entryDic
    #lineNR = int(entryDic[name])
    window = Toplevel(root)
    #window.title(name)
    xs = 450; ys = 200
    window.minsize(xs,ys)
    window.maxsize(xs,ys)

    leftFrame = Frame(window, height=ys, width=210)
    leftFrame.pack_propagate(0) # don't shrink
    leftFrame.pack(side=LEFT)

    nameFrame = Frame(leftFrame, height=ys-124, width=210)
    nameFrame.pack_propagate(0) # don't shrink
    nameFrame.pack(side=TOP)
    nameFrame.configure(bg="#040030")

    nameT = Text(nameFrame, height=1, width=22)
    nameT.pack(pady=(9, 0))
    nameT.pack_propagate(0)
    nameT.insert(END, "123456789012345678901234567890")
    nameT.configure(font=("Helvetica", 13))

    #initiative + special
    botNameFrame = Frame(nameFrame, height=(ys-124)/2, width=210)
    botNameFrame.pack_propagate(0) # don't shrink
    botNameFrame.pack(side=BOTTOM, padx=(4,0))
    botNameFrame.configure(bg="#040030")

    iniT = Text(botNameFrame, height=1, width=10)
    iniT.pack(padx=(2,1), side=LEFT)
    iniT.pack_propagate(0)
    iniT.insert(END, "Initiative: xD")
    iniT.configure(font=("Helvetica", 13))

    specialb = Button(botNameFrame, text="Show Special", height=1, width=13)
    specialb.pack(pady=1, padx=(1, 3), side=RIGHT)

    #Stats
    statFrame = Frame(leftFrame, height=124, width=210) #30x30 per box + 4 pixel in between
    statFrame.pack_propagate(0) # don't shrink
    statFrame.pack(side=BOTTOM)
    statFrame.configure(bg="#1C1C1C")

    wsTop = insertStatFrame(statFrame, 0, 0, "#6E6E6E")
    wsTopL = statLabel(wsTop, "WS", "#6E6E6E")
    bsTop = insertStatFrame(statFrame, 0, 1, "#6E6E6E")
    bsTopL = statLabel(bsTop, "BS", "#6E6E6E")
    sTop = insertStatFrame(statFrame, 0, 2, "#6E6E6E")
    sTopL = statLabel(sTop, "S", "#6E6E6E")
    tTop = insertStatFrame(statFrame, 0, 3, "#6E6E6E")
    tTopL = statLabel(tTop, "T", "#6E6E6E")
    agTop = insertStatFrame(statFrame, 0, 4, "#6E6E6E")
    agTopL = statLabel(agTop, "Ag", "#6E6E6E")
    intTop = insertStatFrame(statFrame, 0, 5, "#6E6E6E")
    intTopL = statLabel(intTop, "Int", "#6E6E6E")
    wpTop = insertStatFrame(statFrame, 0, 6, "#6E6E6E")
    wpTopL = statLabel(wpTop, "WP", "#6E6E6E")

    wsBot = insertStatFrame(statFrame, 1, 0, "#A4A4A4")
    wsBotL = statText(wsBot, "WS")
    bsBot = insertStatFrame(statFrame, 1, 1, "#A4A4A4")
    bsBotL = statText(bsBot, "WS")
    sBot = insertStatFrame(statFrame, 1, 2, "#A4A4A4")
    sBotL = statText(sBot, "WS")
    tBot = insertStatFrame(statFrame, 1, 3, "#A4A4A4")
    tBotL = statText(tBot, "WS")
    agBot = insertStatFrame(statFrame, 1, 4, "#A4A4A4")
    agBotL = statText(agBot, "WS")
    intBot = insertStatFrame(statFrame, 1, 5, "#A4A4A4")
    intBotL = statText(intBot, "WS")
    wpBot = insertStatFrame(statFrame, 1, 6, "#A4A4A4")
    wpBotL = statText(wpBot, "WS")

    mellemStatFrame0 = Frame(statFrame, height=4, width=30)
    mellemStatFrame1 = Frame(statFrame, height=4, width=30)
    mellemStatFrame2 = Frame(statFrame, height=4, width=30)
    mellemStatFrame3 = Frame(statFrame, height=4, width=30)
    mellemStatFrame4 = Frame(statFrame, height=4, width=30)
    mellemStatFrame5 = Frame(statFrame, height=4, width=30)
    mellemStatFrame6 = Frame(statFrame, height=4, width=30)
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

    felTop = insertStatFrame(statFrame, 3, 0, "#6E6E6E")
    felTopL = statLabel(felTop, "Fel", "#6E6E6E")
    aTop = insertStatFrame(statFrame, 3, 1, "#6E6E6E")
    aTopL = statLabel(aTop, "A", "#6E6E6E")
    wTop = insertStatFrame(statFrame, 3, 2, "#6E6E6E")
    wTopL = statLabel(wTop, "W", "#6E6E6E")
    sbTop = insertStatFrame(statFrame, 3, 3, "#6E6E6E")
    sbTopL = statLabel(sbTop, "SB", "#6E6E6E")
    tbTop = insertStatFrame(statFrame, 3, 4, "#6E6E6E")
    tbTopL = statLabel(tbTop, "TB", "#6E6E6E")
    mTop = insertStatFrame(statFrame, 3, 5, "#6E6E6E")
    mTopL = statLabel(mTop, "M", "#6E6E6E")
    magTop = insertStatFrame(statFrame, 3, 6, "#6E6E6E")
    magTopL = statLabel(magTop, "Mag", "#6E6E6E")

    felBot = insertStatFrame(statFrame, 4, 0, "#A4A4A4")
    felBotL = statText(felBot, "WS")
    aBot = insertStatFrame(statFrame, 4, 1, "#A4A4A4")
    aBotL = statText(aBot, "WS")
    wBot = insertStatFrame(statFrame, 4, 2, "#A4A4A4")
    wBotL = statText(wBot, "WS")
    sbBot = insertStatFrame(statFrame, 4, 3, "#A4A4A4")
    sbBotL = statText(sbBot, "WS")
    tbBot = insertStatFrame(statFrame, 4, 4, "#A4A4A4")
    tbBotL = statText(tbBot, "WS")
    mBot = insertStatFrame(statFrame, 4, 5, "#A4A4A4")
    mBotL = statText(mBot, "WS")
    magBot = insertStatFrame(statFrame, 4, 6, "#A4A4A4")
    magBotL = statText(magBot, "WS")

    #right side
    rightFrame = Frame(window, height=ys, width=240)
    rightFrame.pack(side=RIGHT)
    rightFrame.pack_propagate(0) # don't shrink
    rightFrame.configure(bg="#383a39")

    skillT = Text(rightFrame, height=6, width=58)
    skillT.pack(pady=(1, 0))
    skillT.pack_propagate(0)
    skillT.insert(END, "this is a test\nthis is test xD\n12345678901234567890123456789012345678901234567890123456789012345678901234567890")
    skillT.configure(font=("Helvetica", 6))

    talentT = Text(rightFrame, height=6, width=58)
    talentT.pack(pady=(1, 0))
    talentT.pack_propagate(0)
    talentT.insert(END, "this is a test\nthis is test xD\n12345678901234567890123456789012345678901234567890123456789012345678901234567890")
    talentT.configure(font=("Helvetica", 6))

    trappingsT = Text(rightFrame, height=3, width=58)
    trappingsT.pack(pady=(1, 0))
    trappingsT.pack_propagate(0)
    trappingsT.insert(END, "this is a test\nthis is test xD\n12345678901234567890123456789012345678901234567890123456789012345678901234567890")
    trappingsT.configure(font=("Helvetica", 6))

    extraFrame = Frame(rightFrame, height=32, width=240)
    extraFrame.pack(pady=1, padx=1, side=BOTTOM)
    extraFrame.pack_propagate(0)
    extraFrame.configure(bg="#383a39")

    weaponT = Text(extraFrame, height=2, width=35)
    weaponT.pack(pady=(1, 0), padx=(1, 0), side=LEFT)
    weaponT.pack_propagate(0)
    weaponT.insert(END, "this is a test\nthis is test xD\n12345678901234567890123456789012345678901234567890123456789012345678901234567890")
    weaponT.configure(font=("Helvetica", 7))

    saveb = Button(extraFrame, text="Save", height=1, width=6)
    saveb.pack(pady=1, padx=2, side=RIGHT)

#UI

root = Tk()
root.title("New Worlds Beastiary")
root.minsize(400,360)
root.maxsize(400,360)

def insertMellemFrame(height, place):
    mellemFrame = Frame(place, height=height, width=400)
    mellemFrame.pack_propagate(0) # don't shrink
    mellemFrame.pack()
    mellemFrame.configure(bg="blue")

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
autoFrame.configure(bg="red")

insertMellemFrame(30, root) # FIXME:

testb = Button(root, text="test", command=openByName, height=1, width=20)
testb.pack()

selectData()

root.mainloop()

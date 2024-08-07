import tkinter
from tkinter import *

from tkinter.ttk import *
from tkinter import messagebox

from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename

from openpyxl import load_workbook

import os

from time import time, sleep
from datetime import datetime, timedelta

import subprocess as sp

class Application(Frame):

    def __init__(self, master):

        self.master = master
        self.main_container = Frame(self.master)

        # Define the source and target folder variables

        self.origin = os.getcwd()
        self.source = ""
        self.target = ""
        self.initFolders = IntVar()
        self.ftype = IntVar()
        self.pointer = 0
        self.identifier = StringVar()
        self.sheet = StringVar()
        self.sheetId = StringVar()
        self.allMoves = []
        self.winner = IntVar()
        self.advantage = IntVar()
        self.showFlag = IntVar()
        self.sheetsList = ['No Selection']
        self.sheetIdsList = ['No Selection']

        # Create main frame
        self.main_container.grid(column=0, row=0, sticky=(N,S,E,W))

        # Set Label styles
        Style().configure("M.TLabel", font="Courier 20 bold", height="20", foreground="blue", anchor="center")
        Style().configure("B.TLabel", font="Verdana 8", background="white", width="40")
        Style().configure("G.TLabel", font="Verdana 8")
        Style().configure("L.TLabel", font="Courier 20 bold", width="8")
        Style().configure("MS.TLabel", font="Verdana 10" )
        Style().configure("S.TLabel", font="Verdana 8" )
        Style().configure("G.TLabel", font="Verdana 8")

        # Set button styles
        Style().configure("B.TButton", font="Verdana 8", relief="ridge")

        # Set check button styles
        Style().configure("B.TCheckbutton", font="Verdana 8")
        Style().configure("B.TRadiobutton", font="Verdana 8")
        Style().configure("O.TLabelframe.Label", font="Verdana 8", foreground="black")

        # Create widgets
        self.sep_a = Separator(self.main_container, orient=HORIZONTAL)
        self.sep_b = Separator(self.main_container, orient=HORIZONTAL)
        self.sep_c = Separator(self.main_container, orient=HORIZONTAL)
        self.sep_d = Separator(self.main_container, orient=HORIZONTAL)
        self.sep_e = Separator(self.main_container, orient=HORIZONTAL)
        self.sep_f = Separator(self.main_container, orient=HORIZONTAL)
        self.mainLabel = Label(self.main_container, text="MOVE LISTER", style="M.TLabel" )
        self.subLabelA = Label(self.main_container, text="Displays chess moves listed in Excel spreadsheet in a more readable ", style="S.TLabel" )
        self.subLabelB = Label(self.main_container, text="format for easier following instead of straining your eyes to read. ", style="S.TLabel" )
        self.subLabelC = Label(self.main_container, text="The files may be actual games or just openings. ", style="S.TLabel" )

        self.sourceFolder = LabelFrame(self.main_container, text=' FILE OPTIONS ', style="O.TLabelframe")
        self.selectSource = Button(self.sourceFolder, text="NAME", style="B.TButton", command=self.setSource)
        self.sourceLabel = Label(self.sourceFolder, text="None", style="B.TLabel" )
        self.sheetLabel = Label(self.sourceFolder, text="   SHEETS  ", style="G.TLabel" )
        self.selectSheet = OptionMenu(self.sourceFolder, self.sheet, *self.sheetsList, command=self.getSheetIdsList)
        self.idLabel = Label(self.sourceFolder, text="   IDENTIFIERS  ", style="G.TLabel" )
        self.id = Entry(self.sourceFolder, textvariable=self.identifier, width="5")
        # self.idList = OptionMenu(self.sourceFolder, self.sheetId, *self.sheetIdsList)
        self.idList = Listbox(self.sourceFolder, selectmode='single', width=40, height=5)
        self.iscroller = Scrollbar(self.sourceFolder, orient=VERTICAL, command=self.idList.yview)
        self.idList.config(font=("Courier New", 8), yscrollcommand=self.iscroller.set)

        self.sep_s = Separator(self.sourceFolder, orient=HORIZONTAL)
        self.sep_t = Separator(self.sourceFolder, orient=HORIZONTAL)
        self.sep_u = Separator(self.sourceFolder, orient=HORIZONTAL)

        # self.whiteFrame = LabelFrame(self.main_container, text=' WHITE ', style="O.TLabelframe")
        # self.whiteMove = Label(self.whiteFrame, text=" ", style="L.TLabel" )
        # self.blackFrame = LabelFrame(self.main_container, text=' BLACK ', style="O.TLabelframe")
        # self.blackMove = Label(self.blackFrame, text=" ", style="L.TLabel" )
        # self.showAll = Button(self.main_container, text="SHOW ALL", style="B.TButton", command=self.displayMovesPanel)

        # self.next = Button(self.main_container, text="NEXT", style="B.TButton", command=self.getNextMove)
        # self.prev = Button(self.main_container, text="PREV", style="B.TButton", command=self.getPrevMove)
        # self.restart = Button(self.main_container, text="RESTART", style="B.TButton", command=self.restartMoves)

        self.typeOptions = LabelFrame(self.sourceFolder, text=' FILE TYPE ', style="O.TLabelframe")
        self.gameType = Radiobutton(self.typeOptions, text="Complete Games ", style="B.TRadiobutton", variable=self.ftype, value=0)
        self.openType = Radiobutton(self.typeOptions, text="Opening moves ", style="B.TRadiobutton", variable=self.ftype, value=1)

        self.start = Button(self.main_container, text="PLAY", style="B.TButton", command=self.startGame)
        self.reset = Button(self.main_container, text="RESET", style="B.TButton", command=self.resetProcess)
        self.exit = Button(self.main_container, text="EXIT", style="B.TButton", command=root.destroy)

        self.progress_bar = Progressbar(self.main_container, orient="horizontal", mode="indeterminate", maximum=50)

        # Position widgets
        self.mainLabel.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.sep_a.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')

        self.subLabelA.grid(row=2, column=0, columnspan=3, padx=5, pady=0, sticky='NSEW')
        self.subLabelB.grid(row=3, column=0, columnspan=3, padx=5, pady=0, sticky='NSEW')
        self.subLabelC.grid(row=4, column=0, columnspan=3, padx=5, pady=0, sticky='NSEW')

        self.sep_b.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')

        self.selectSource.grid(row=0, column=0, padx=5, pady=5, sticky='NSEW')
        self.sourceLabel.grid(row=0, column=1, padx=5, pady=5, sticky='NSEW')

        self.sep_s.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.gameType.grid(row=0, column=0, padx=(20,5), pady=(5,10), sticky='NSEW')
        self.openType.grid(row=0, column=1, padx=(80,5), pady=(5,10), sticky='NSEW')
        self.typeOptions.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.sep_t.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.sheetLabel.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        self.selectSheet.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.sep_u.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.idLabel.grid(row=6, column=0, padx=5, pady=(5,10), sticky='NSEW')
        # self.idList.grid(row=6, column=1, padx=5, pady=(5,10), sticky='NSEW')
        self.idList.grid(row=6, column=1, columnspan=1, padx=5, pady=5, sticky='W')
        self.iscroller.grid(row=6, column=2, columnspan=1, padx=5, pady=5, sticky='W')

        # self.id.grid(row=6, column=1, padx=5, pady=(5,10), sticky='NSEW')
        
        self.sourceFolder.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')        

        self.sep_c.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.start.grid(row=9, column=0, columnspan=1, padx=5, pady=0, sticky='NSEW')
        self.reset.grid(row=9, column=1, columnspan=1, padx=5, pady=0, sticky='NSEW')
        self.exit.grid(row=9, column=2, columnspan=1, padx=5, pady=0, sticky='NSEW')
        
        self.sep_d.grid(row=10, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        # self.whiteMove.grid(row=0, column=0, padx=5, pady=5, sticky="NSEW")
        # self.whiteFrame.grid(row=11, column=0, columnspan=1, rowspan=1, padx=5, pady=5, sticky="NSEW")
        
        # self.blackMove.grid(row=0, column=0, padx=5, pady=5, sticky="NSEW")
        # self.blackFrame.grid(row=11, column=1, columnspan=1, rowspan=1, padx=5, pady=5, sticky="NSEW")
        # self.showAll.grid(row=11, column=2, columnspan=1, rowspan=1, padx=5, pady=5, sticky="NSEW")

        # self.prev.grid(row=12, column=0, columnspan=1, padx=5, pady=0, sticky='NSEW')
        # self.next.grid(row=12, column=1, columnspan=1, padx=5, pady=0, sticky='NSEW')
        # self.restart.grid(row=12, column=2, columnspan=1, padx=5, pady=0, sticky='NSEW')

        # self.sep_e.grid(row=13, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        
        
        # self.sep_e.grid(row=15, column=0, columnspan=3, padx=5, pady=5, sticky='NSEW')
        
        self.progress_bar.grid(row=16, column=0, columnspan=3, padx=5, pady=0, sticky='NSEW')

        self.processControl(1)

    def setSource(self):

        pathname = askopenfilename()

        try:
            if self.source.endswith(".xlsx") or self.source.endswith(".xls"):
                messagebox.showerror("Invalid file selected", "Invalid file type was selected. Please select again.")
            else:
                self.sourceLabel["text"] = pathname
                self.source = pathname
                self.getSheetList()
        except:
            pass

    def getSheetList(self):

        wb = load_workbook(self.source)

        self.sheetsList = []

        for ws in wb.worksheets:
            self.sheetsList.append(ws)

        menu = self.selectSheet['menu']
        menu.delete(0, 'end')

        for sh in self.sheetsList:
            menu.add_command(label=sh.title, command=lambda value=sh.title: self.getSheetIdsList(value))


    def getSheetIdsList(self, sheet):

        self.sheet.set(sheet)
        if self.ftype.get() == 0:
            self.getGameIds(sheet)
        else:
            self.getOpeningIds(sheet)

        self.idList.delete(0, END)
        for id in self.sheetIdsList:
            self.idList.insert(END, id)

        # menu = self.idList['menu']
        # menu.delete(0, 'end')

        # for id in self.sheetIdsList:
        #     menu.add_command(label=id, command=lambda value=id: self.sheetId.set(value))


    def getOpeningIds(self, sheet):

        wb = load_workbook(self.source)
        ws = wb[sheet]

        cola = ['A', 'D', 'G', 'J', 'M', 'P', 'S', 'V']

        self.sheetIdsList = []

        count = 1
        eol = False
        
        while True:
            
            for col in cola:
                
                cell = col + str(count)
                
                if ws[cell].value == 'END':
                    eol = True
                    break
                    
                if ws[cell].value:
                    try:
                        id = ws[cell].value + ' - ' + ws[cell].comment.text.rstrip()
                    except:
                        id = ws[cell].value 

                    self.sheetIdsList.append(id)
                
            if eol:
                break
                
            count += 21

    def getGameIds(self, sheet):
        
        wb = load_workbook(self.source)
        
        ws = wb[sheet]

        cola = ['A', 'D', 'G', 'J', 'M', 'P', 'S', 'V']

        self.sheetIdsList = []

        count = 1
        
        for col in cola:
            
            cell = col + str(count)

            if ws[cell].value:
                try:
                    id = ws[cell].value + ' - ' + ws[cell].comment.text.rstrip()
                except:
                    id = ws[cell].value 
                    
                self.sheetIdsList.append(id)



    def startGame(self):

        if self.sheet.get():
            pass
        else: 
            messagebox.showerror("Select sheet", "Please select sheet from workbook.")
            return 

        # self.idList.curselection()
        # if self.idList.get():
        #     print(self.idList.get())

        if self.idList.curselection():
            self.sheetId.set(self.idList.get(self.idList.curselection()[0]))
        else: 
            messagebox.showerror("Select sheet ID", "Please select sheet ID from worksheet.")
            return 

        self.loadAllMoves()
        self.displayMovesPanel()
        self.processControl(1)

    def loadAllMoves(self):

        if self.ftype.get() == 0:
            return self.loadGameMoves()
        else:
            return self.loadOpenMoves()
        
    def loadGameMoves(self):

        wb = load_workbook(self.source)
        ws = wb[self.sheet.get()]

        cola = ['A', 'D', 'G', 'J', 'M', 'P', 'S', 'V']
        colb = ['B', 'E', 'H', 'K', 'N', 'Q', 'T', 'W']
        
        found_tag = False 

        for ca, cb in zip(cola, colb):
            
            cell = ca + '1'
            
            if ws[cell].value == self.sheetId.get():    
                found_tag = True 

                self.allMoves = []
                count = 2
                
                while True:
                    
                    cell = ca + str(count)
                    
                    if ws[cell].value:
                        self.allMoves.append(ws[cell].value) 
                    else:
                        break

                    cell = cb + str(count)
                    
                    if ws[cell].value:
                        self.allMoves.append(ws[cell].value)
                    else:
                        break
                    
                    count += 1

                self.pointer = 0

                break 

        return found_tag 

    def loadOpenMoves(self):

        print(self.sheetId.get())
        wb = load_workbook(self.source)
        ws = wb[self.sheet.get()]

        cola = ['A', 'D', 'G', 'J', 'M', 'P', 'S', 'V']
        colb = ['B', 'E', 'H', 'K', 'N', 'Q', 'T', 'W']
        
        found_tag = False 

        count = 1
        
        while True:

            for ca, cb in zip(cola, colb):
                    
                cell = ca + str(count)

                if ws[cell].value == self.sheetId.get()[:6]:    

                    found_tag = True 

                    self.allMoves = []
                    line_count = count + 1
                    
                    while True:
                        
                        cell = ca + str(line_count)
                        
                        if ws[cell].value:
                            self.allMoves.append(ws[cell].value)
                            bgColor = ws[cell].fill.fgColor.index
                            self.checkAdvantage(bgColor)
 
                        else:
                            break

                        cell = cb + str(line_count)
                        
                        if ws[cell].value:
                            self.allMoves.append(ws[cell].value)
                            bgColor = ws[cell].fill.fgColor.index
                            self.checkAdvantage(bgColor)

                        else:
                            break
                        
                        line_count += 1 

                    self.pointer = 0

                    break

                elif ws[cell].value == "END":
                    break

            if found_tag:
                break 

            count += 1

        return found_tag 
    
    def postFirstMove(self):

        move = self.allMoves[self.pointer]
        self.whiteMove["text"] = move
        self.blackMove["text"] = ""

    def getNextMove(self):

        if self.pointer + 1== len(self.allMoves):
            messagebox.showinfo("Last moves", "Last moves already displayed.")
            return

        self.pointer += 1

        move = self.allMoves[self.pointer]

        if self.pointer % 2 == 1:
            self.blackMove["text"] = move
        else:
            self.whiteMove["text"] = move
            self.blackMove["text"] = ""

    def getPrevMove(self):

        if self.pointer == 0:
            messagebox.showinfo("First moves", "First moves already displayed.")
            return

        self.pointer -= 1

        move = self.allMoves[self.pointer]

        if self.pointer % 2 == 1:
            self.blackMove["text"] = move 

            white = self.allMoves[self.pointer - 1]
            self.whiteMove["text"] = white
        else:
            self.whiteMove["text"] = move
            self.blackMove["text"] = ""


    def get_tag(self, tag):
        
        wb2 = load_workbook(self.source)
        
        try:
            ws_index = wb2['Index']
            
            count = 2
            
            while True:
                
                cell = 'A' + str(count)
                
                if ws_index[cell].value == tag:
                    opening = ws_index['B' + str(count)].value
                    white = ws_index['D' + str(count)].value
                    black = ws_index['E' + str(count)].value

                    return opening, white, black
                    
                count += 1          
        except:
            return self.sheet.get(), "Player 1", "Player 2"
        

    def displayMovesPanel(self):

        self.popMoves = Toplevel(self.main_container)
        self.popMoves.title("Moves")

        self.pop_a = Separator(self.popMoves, orient=HORIZONTAL)
        self.pop_b = Separator(self.popMoves, orient=HORIZONTAL)
        self.pop_c = Separator(self.popMoves, orient=HORIZONTAL)
        self.pop_d = Separator(self.popMoves, orient=HORIZONTAL)
        self.pop_e = Separator(self.popMoves, orient=HORIZONTAL)

        self.whiteFrame = LabelFrame(self.popMoves, text=' WHITE ', style="O.TLabelframe")
        self.whiteMove = Label(self.whiteFrame, text=" ", style="L.TLabel" )
        self.blackFrame = LabelFrame(self.popMoves, text=' BLACK ', style="O.TLabelframe")
        self.blackMove = Label(self.blackFrame, text=" ", style="L.TLabel" )
        self.showAll = Button(self.popMoves, text="SHOW ALL", style="B.TButton", command=self.showAllMoves)

        self.next = Button(self.popMoves, text="NEXT", style="B.TButton", command=self.getNextMove)
        self.prev = Button(self.popMoves, text="PREV", style="B.TButton", command=self.getPrevMove)
        self.restart = Button(self.popMoves, text="RESTART", style="B.TButton", command=self.restartMoves)

        self.popOpening = Label(self.popMoves, text=" ", style="S.TLabel" )
        self.popPlayers = Label(self.popMoves, text=" ", style="S.TLabel" )

        self.moveListFrame = LabelFrame(self.popMoves, text=' MOVES LIST ', style="O.TLabelframe")
        self.moveList = Listbox(self.moveListFrame, width=38, height=8)
        self.scroller = Scrollbar(self.moveListFrame, orient=VERTICAL, command=self.moveList.yview)
        self.moveList.config(font=("Courier New", 10), yscrollcommand=self.scroller.set)
        
        self.popOpening.configure(font=("Courier New", 10))
        self.popPlayers.configure(font=("Courier New", 10))

        self.close = Button(self.popMoves, text="CLOSE", style="B.TButton", command=self.hidePopup)

        self.popOpening.grid(row=2, column=0, columnspan=2, padx=5, pady=1, sticky="NSEW")
        self.popPlayers.grid(row=1, column=0, columnspan=3, padx=5, pady=1, sticky="NSEW")

        self.pop_a.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="NSEW")

        self.whiteMove.grid(row=0, column=0, padx=5, pady=1, sticky="NSEW")
        self.whiteFrame.grid(row=5, column=0, columnspan=1, padx=5, pady=1, sticky="NSEW")
        self.blackMove.grid(row=0, column=0, padx=5, pady=1, sticky="NSEW")
        self.blackFrame.grid(row=5, column=1, columnspan=1, padx=5, pady=1, sticky="NSEW")
        self.showAll.grid(row=5, column=2, columnspan=1, padx=5, pady=1, sticky="NSEW")

        self.pop_b.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="NSEW")
        
        self.prev.grid(row=7, column=0, columnspan=1, padx=5, pady=1, sticky="NSEW")
        self.next.grid(row=7, column=1, columnspan=1, padx=5, pady=1, sticky="NSEW")
        self.restart.grid(row=7, column=2, columnspan=1, padx=5, pady=1, sticky="NSEW")

        self.pop_c.grid(row=8, column=0, columnspan=3, padx=5, pady=5, sticky="NSEW")
        
        self.close.grid(row=9, column=0, columnspan=3, padx=5, pady=5, sticky="NSEW")

        self.pop_d.grid(row=10, column=0, columnspan=3, padx=5, pady=5, sticky="NSEW")

        self.moveList.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky='NSEW')
        self.scroller.grid(row=0, column=2, columnspan=1, padx=5, pady=5, sticky='NSEW')
        self.moveListFrame.grid(row=11, column=0, columnspan=3, padx=5, pady=5, sticky="NSEW")

        self.pop_e.grid(row=8, column=0, columnspan=3, padx=5, pady=5, sticky="NSEW")
        
        opening, white, black = self.get_tag(self.sheetId.get())

        # self.popOpening["text"] = opening
        self.popPlayers["text"] = white + " vs. " + black

        self.postFirstMove()
        self.loadList()
        self.showFlag.set(0)

        if self.ftype.get() == 0:
            if self.winner.get() == 1:
                self.popOpening["text"] = opening + " (W)"
            else: 
                self.popOpening["text"] = opening + " (B)"
        else:
            if self.advantage.get() == 2:
                self.popOpening["text"] = "White"
            elif self.advantage.get() == 1:
                self.popOpening["text"] = "Even"
            else:
                self.popOpening["text"] = "Black"

        ph = 210
        pw = 410

        self.popMoves.maxsize(pw, ph)
        self.popMoves.minsize(pw, ph)

        ws = self.popMoves.winfo_screenwidth()
        hs = self.popMoves.winfo_screenheight()

        x = (ws/2) - (pw/2)
        y = (hs/2) - (ph/2)

        self.popMoves.geometry('%dx%d+%d+%d' % (pw, ph, x, y))

        
    def loadList(self):

        if self.ftype.get() == 0:
            return self.loadGameList()
        else:
            return self.loadOpenList()
        
    def loadGameList(self):

        self.moveList.delete(0, END)

        wb = load_workbook(self.source)
        ws = wb[self.sheet.get()]

        cola = ['A', 'D', 'G', 'J', 'M', 'P', 'S', 'V']
        colb = ['B', 'E', 'H', 'K', 'N', 'Q', 'T', 'W']
        
        found_tag = False 

        for ca, cb in zip(cola, colb):
            
            cell = ca + '1'
            
            if ws[cell].value == self.sheetId.get()[:7]:    

                found_tag = True 

                count = 2
                
                while True:
                    
                    cell = ca + str(count)
                    
                    if ws[cell].value:
                        white_move = ws[cell].value
                    else:
                        self.winner.set(0)
                        break

                    cell = cb + str(count)
                    
                    if ws[cell].value:
                        black_move = ws[cell].value
                    else:
                        self.moveList.insert(END, white_move)
                        self.winner.set(1)
                        break
                    
                    self.moveList.insert(END, white_move.ljust(6) + '  -  ' + black_move)

                    count += 1

                break
        

    def loadOpenList(self):

        self.moveList.delete(0, END)

        wb = load_workbook(self.source)
        ws = wb[self.sheet.get()]

        cola = ['A', 'D', 'G', 'J', 'M', 'P', 'S', 'V']
        colb = ['B', 'E', 'H', 'K', 'N', 'Q', 'T', 'W']

        found_tag = False 

        count = 1

        while True:

            for ca, cb in zip(cola, colb):
                
                cell = ca + str(count)

                if ws[cell].value == self.sheetId.get()[:6]:
                    
                    found_tag = True 

                    count += 1
                    
                    while True:
                        
                        cell = ca + str(count)
                        
                        if ws[cell].value:
                            white_move = ws[cell].value
                        else:
                            break

                        cell = cb + str(count)
                        
                        if ws[cell].value:
                            black_move = ws[cell].value
                        else:
                            self.moveList.insert(END, white_move)
                            break
                        
                        self.moveList.insert(END, white_move.ljust(6) + '  -  ' + black_move)

                        count += 1

                    break

            if found_tag:
                break

            count += 21
        
    def checkAdvantage(self, fill):

        self.advantage.set(1)
        if fill == "FF00FF00":          # green
            self.advantage.set(2)
        elif fill == "FFFFFF00":        # yellow
            self.advantage.set(1)
        elif fill == "FF00FFFF":        # turquise
            self.advantage.set(0)

    def showAllMoves(self):

        if self.showFlag.get() == 0:
            self.showFlag.set(1)
            self.showAll.configure(text="HIDE ALL")
            ph = 380
        else:
            self.showFlag.set(0)
            self.showAll.configure(text="SHOW ALL")
            ph = 210

        pw = 410

        self.popMoves.maxsize(pw, ph)
        self.popMoves.minsize(pw, ph)

        ws = self.popMoves.winfo_screenwidth()
        hs = self.popMoves.winfo_screenheight()

        x = (ws/2) - (pw/2)
        y = (hs/2) - (ph/2)

        self.popMoves.geometry('%dx%d+%d+%d' % (pw, ph, x, y))

        self.processControl(0)

    def hidePopup(self):
        self.processControl(1)
        self.popMoves.destroy()

    def restartMoves(self):

        self.pointer = 0
        self.postFirstMove()

        messagebox.showinfo("Game restarted", "Game has restarted, first move shown.") 

    def processControl(self, mode):
        ''' enable/disable buttons as needed
        '''

        if mode:
            
            self.start["state"] = NORMAL
            self.exit["state"] = NORMAL

        else:

            self.start["state"] = DISABLED
            self.exit["state"] = DISABLED


    def resetProcess(self):
        ''' reset labels, lists and flags
        '''
        
        os.chdir(self.origin)
        self.sourceLabel["text"] = "None"
        self.source = ""

        self.idList.delete(0, END)
        # menu = self.idList['menu']
        # menu.delete(0, 'end')

        menu = self.selectSheet['menu']
        menu.delete(0, 'end')

        self.sheet.set('')
        self.sheetId.set('')


root = Tk()
root.title("MOVES LIST")

# Set size

wh = 460
ww = 450

# root.resizable(height=False, width=False)

# root.minsize(ww, wh)
# root.maxsize(ww, wh)

# Position in center screen

ws = root.winfo_screenwidth()
hs = root.winfo_screenheight()

# calculate x and y coordinates for the Tk root window
x = (ws/2) - (ww/2)
y = (hs/2) - (wh/2)

root.geometry('%dx%d+%d+%d' % (ww, wh, x, y))

app = Application(root)

root.mainloop()

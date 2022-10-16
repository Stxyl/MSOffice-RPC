# tkinter import
from tkinter import *

# Office Applications python script import
from Word_RPC import *
from Excel_RPC import *
from PowerPoint_RPC import *
from Outlook_RPC import *
from OneNote_RPC import *
from Teams_RPC import *


class GUI:
    def __init__(self, master):
        self.master = master
        master.title("MSOffice RPC")

        self.label = Label(master, text="Click the following buttons to run RPC for your office application")
        self.label.pack()

        self.Word = Button(master, text="Word", command=self.Word)
        self.Word.pack()

        self.PowerPoint = Button(master, text="PowerPoint", command=self.PowerPoint)
        self.PowerPoint.pack()

        self.Excel = Button(master, text="Excel", command=self.Excel)
        self.Excel.pack()

        self.OneNote = Button(master, text="OneNote", command=self.OneNote)
        self.OneNote.pack()

        self.Outlook = Button(master, text="Outlook", command=self.Outlook)
        self.Outlook.pack()

        self.Teams = Button(master, text="Teams", command=self.Teams)
        self.Teams.pack()

        self.exit = Button(master, text="Exit", command=master.quit)
        self.exit.pack()

    def Word(self):
        print("Button Pressed - Word")
        print("Preforming Actions...")

        # checks if Word is installed
        appget_word()

        # checks if Word is running
        if appget_pid_word() != "":
            print("Word is running", appget_pid_word())
            update_presence_word()
        else:
            print("Word is not running")

        while True:
            appget_pid_word()

            time.sleep(1)  # Can only update rich presence every 1 second

    def PowerPoint(self):
        print("Button Pressed - PowerPoint")
        print("Preforming Actions...")

        # checks if PowerPoint is installed
        appget_powerpoint()

        # checks if PowerPoint is running
        if appget_pid_powerpoint() != "":
            print("PowerPoint is running", appget_pid_powerpoint())
            update_presence_powerpoint()
        else:
            print("PowerPoint is not running")

        while True:
            appget_pid_powerpoint()

            time.sleep(1)  # Can only update rich presence every 1 second

    def Excel(self):
        print("Button Pressed - Excel")
        print("Preforming Actions...")

        # checks if Excel is installed
        appget_excel()

        # checks if Excel is running
        if appget_pid_excel() != "":
            print("Excel is running", appget_pid_excel())
            update_presence_excel()
        else:
            print("Excel is not running")

        while True:
            appget_pid_excel()

            time.sleep(1) # Can only update rich presence every 1 second

    def OneNote(self):
        print("Button Pressed - OneNote")
        print("Preforming Actions...")

        # checks if OneNote is installed
        appget_onenote()

        # checks if OneNote is running
        if appget_pid_onenote() != "":
            print("OneNote is running", appget_pid_onenote())
            update_presence_onenote()
        else:
            print("OneNote is not running")

        while True:
            appget_pid_onenote()

            time.sleep(1) # Can only update rich presence every 1 second

    def Outlook(self):
        print("Button Pressed - Outlook")
        print("Preforming Actions...")

        # checks if Outlook is installed
        appget_outlook()

        # checks if Outlook is running
        if appget_pid_outlook() != "":
            print("Outlook is running", appget_pid_outlook())
            update_presence_outlook()
        else:
            print("Outlook is not running")

        while True:
            appget_pid_outlook()

            time.sleep(1) # Can only update rich presence every 1 second

    def Teams(self):
        print("Button Pressed - Teams")
        print("Preforming Actions...")

        # checks if Teams is installed
        appget_teams()

        # checks if Teams is running
        if appget_pid_teams() != "":
            print("Teams is running", appget_pid_teams())
            update_presence_teams()
        else:
            print("Teams is not running")

        while True:
            appget_pid_teams()

            time.sleep(1) # Can only update rich presence every 1 second


root = Tk()
root.geometry("500x500")
my_gui = GUI(root)
root.mainloop()

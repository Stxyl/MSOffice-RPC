import os
import pypresence
import time


# This function is used to connect to Discord
def connect_discord():
    RPC = (pypresence.Presence("1014286705056043129"))
    RPC.connect()
    return RPC


# This function is used to update the presence
def update_presence_word():
    RPC = connect_discord()
    RPC.update(state="Editing a document", large_image="word", large_text="Microsoft Word")
    print("Presence updated")
    connect_discord()


def update_presence_excel():
    RPC = connect_discord()
    RPC.update(state="Editing a spreadsheet", large_image="excel", large_text="Microsoft Excel")
    print("Presence updated")
    connect_discord()


def update_presence_powerpoint():
    RPC = connect_discord()
    RPC.update(state="Editing a presentation", large_image="powerpoint", large_text="Microsoft PowerPoint")
    print("Presence updated")
    connect_discord()


def update_presence_onenote():
    RPC = connect_discord()
    RPC.update(state="Taking notes", large_image="onenote", large_text="Microsoft OneNote")
    print("Presence updated")
    connect_discord()


def update_presence_outlook():
    RPC = connect_discord()
    RPC.update(state="Reading emails", large_image="outlook", large_text="Microsoft Outlook")
    print("Presence updated")
    connect_discord()


def update_presence_teams():
    RPC = connect_discord()
    RPC.update(state="Chatting", large_image="teams", large_text="Microsoft Teams")
    print("Presence updated")
    connect_discord()


# These functions find's where the application is stored and if its running
def appget_word():
    if os.path.exists("/Applications/Microsoft Word.app/Contents/MacOS/Microsoft Word"):
        print("Located Word")
        return True
    else:
        print("Word not found (You sure you located it properly?)")
        return False


# Tells console if the application is found
appget_word()


def appget_excel():
    if os.path.exists("/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel"):
        print("Located Excel")
        return True
    else:
        print("Excel not found (You sure you located it properly?)")
        return False


appget_excel()


def appget_powerpoint():
    if os.path.exists("/Applications/Microsoft PowerPoint.app/Contents/MacOS/Microsoft PowerPoint"):
        print("Located PowerPoint")
        return True
    else:
        print("PowerPoint not found (You sure you located it properly?)")
        return False


appget_powerpoint()


def appget_onenote():
    if os.path.exists("/Applications/Microsoft OneNote.app/Contents/MacOS/Microsoft OneNote"):
        print("Located OneNote")
        return True
    else:
        print("OneNote not found (You sure you located it properly?)")
        return False


appget_onenote()


def appget_outlook():
    if os.path.exists("/Applications/Microsoft Outlook.app/Contents/MacOS/Microsoft Outlook"):
        print("Located Outlook")
        return True
    else:
        print("Outlook not found (You sure you located it properly?)")
        return False


appget_outlook()


def appget_teams():
    if os.path.exists("/Applications/Microsoft Teams.app/Contents/MacOS/Teams"):
        print("Located Teams")
        return True
    else:
        print("Teams not found (You sure you located it properly?)")
        return False


appget_teams()


# This function finds the process ID of the application
def appget_pid_word():
    pid = os.popen("pgrep 'Microsoft Word'").read()
    return pid


# This function finds the process ID of the application and tells you if its running or not with PID Number
if appget_pid_word() != "":
    print("Word is running", appget_pid_word())
    update_presence_word()
else:
    print("Word is not running")


def appget_pid_excel():
    pid = os.popen("pgrep 'Microsoft Excel'").read()
    return pid


if appget_pid_excel() != "":
    print("Excel is running", appget_pid_excel())
    update_presence_excel()
else:
    print("Excel is not running")


def appget_pid_powerpoint():
    pid = os.popen("pgrep 'Microsoft PowerPoint'").read()
    return pid


if appget_pid_powerpoint() != "":
    print("PowerPoint is running", appget_pid_powerpoint())
    update_presence_powerpoint()
else:
    print("PowerPoint is not running")


def appget_pid_onenote():
    pid = os.popen("pgrep 'Microsoft OneNote'").read()
    return pid


if appget_pid_onenote() != "":
    print("OneNote is running", appget_pid_onenote())
    update_presence_onenote()
else:
    print("OneNote is not running")


def appget_pid_outlook():
    pid = os.popen("pgrep 'Microsoft Outlook'").read()
    return pid


if appget_pid_outlook() != "":
    print("Outlook is running", appget_pid_outlook())
    update_presence_outlook()
else:
    print("Outlook is not running")


def appget_pid_teams():
    pid = os.popen("pgrep 'Teams'").read()
    return pid


if appget_pid_teams() != "":
    print("Teams is running", appget_pid_teams())
    update_presence_teams()
else:
    print("Teams is not running")


# loop the function above
while True:
    appget_pid_word()
    appget_pid_excel()
    appget_pid_powerpoint()
    appget_pid_onenote()
    appget_pid_outlook()
    appget_pid_teams()

    time.sleep(1)  # Can only update rich presence every 1 second

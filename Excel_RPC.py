import os
import pypresence
import time


# This function is used to connect to Discord
def connect_discord():
    RPC = (pypresence.Presence("1014791212013735977"))
    RPC.connect()
    return RPC


# This function is used to update the presence
def update_presence_excel():
    RPC = connect_discord()
    RPC.update(state="Editing a spreadsheet", large_image="excel", large_text="Microsoft Excel")
    print("Presence updated")
    connect_discord()


# These functions find's where the application is stored and if its running
def appget_excel():
    if os.path.exists("/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel"):
        print("Located Excel")
        return True
    else:
        print("Excel not found (You sure you located it properly?)")
        return False


# This function finds the process ID of the application
def appget_pid_excel():
    pid = os.popen("pgrep 'Microsoft Excel'").read()
    return pid



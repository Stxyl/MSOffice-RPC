import os
import pypresence
import time


# This function is used to connect to Discord
def connect_discord():
    RPC = (pypresence.Presence("1014791434349576243"))
    RPC.connect()
    return RPC


# This function is used to update the presence
def update_presence_outlook():
    RPC = connect_discord()
    RPC.update(state="Editing a document", large_image="outlook", large_text="Microsoft Outlook")
    print("Presence updated")
    connect_discord()


# These functions find's where the application is stored and if its running
def appget_outlook():
    if os.path.exists("/Applications/Microsoft Outlook.app/Contents/MacOS/Microsoft Outlook"):
        print("Located Outlook")
        return True
    else:
        print("Outlook not found (You sure you located it properly?)")
        return False


def appget_pid_outlook():
    pid = os.popen("pgrep 'Microsoft Outlook'").read()
    return pid




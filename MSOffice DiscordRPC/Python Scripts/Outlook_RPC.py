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


appget_outlook()


def appget_pid_outlook():
    pid = os.popen("pgrep 'Microsoft Outlook'").read()
    return pid


# This function finds the process ID of the application and tells you if its running or not with PID Number
if appget_pid_outlook() != "":
    print("Outlook is running", appget_pid_outlook())
    update_presence_outlook()
else:
    print("Outlook is not running")

while True:
    appget_pid_outlook()

    time.sleep(1)  # Can only update rich presence every 1 second

import os
import pypresence
import time


# This function is used to connect to Discord
def connect_discord():
    RPC = (pypresence.Presence("1014791339436679299"))
    RPC.connect()
    return RPC


# This function is used to update the presence
def update_presence_onenote():
    RPC = connect_discord()
    RPC.update(state="Taking notes", large_image="onenote", large_text="Microsoft OneNote")
    print("Presence updated")
    connect_discord()


# These functions find's where the application is stored and if its running
def appget_onenote():
    if os.path.exists("/Applications/Microsoft onenote.app/Contents/MacOS/Microsoft OneNote"):
        print("Located OneNote")
        return True
    else:
        print("OneNote not found (You sure you located it properly?)")
        return False


# Tells if the application is found or not
appget_onenote()


# This function finds the process ID of the application
def appget_pid_onenote():
    pid = os.popen("pgrep 'Microsoft OneNote'").read()
    return pid


# This function finds the process ID of the application and tells you if its running or not with PID Number
if appget_pid_onenote() != "":
    print("OneNote is running", appget_pid_onenote())
    update_presence_onenote()
else:
    print("OneNote is not running")

while True:
    appget_pid_onenote()

    time.sleep(1)  # Can only update rich presence every 1 second

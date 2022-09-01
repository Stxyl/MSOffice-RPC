import os
import pypresence
import time


# This function is used to connect to Discord
def connect_discord():
    RPC = (pypresence.Presence("1014791041838219274"))
    RPC.connect()
    return RPC


# This function is used to update the presence
def update_presence_powerpoint():
    RPC = connect_discord()
    RPC.update(state="Editing a presentation", large_image="PowerPoint", large_text="Microsoft PowerPoint")
    print("Presence updated")
    connect_discord()


# These functions find's where the application is stored and if its running
def appget_powerpoint():
    if os.path.exists("/Applications/Microsoft PowerPoint.app/Contents/MacOS/Microsoft PowerPoint"):
        print("Located PowerPoint")
        return True
    else:
        print("PowerPoint not found (You sure you located it properly?)")
        return False


# Tells if the application is found or not
appget_powerpoint()


# This function finds the process ID of the application
def appget_pid_powerpoint():
    pid = os.popen("pgrep 'Microsoft PowerPoint'").read()
    return pid


# This function finds the process ID of the application and tells you if its running or not with PID Number
if appget_pid_powerpoint() != "":
    print("PowerPoint is running", appget_pid_powerpoint())
    update_presence_powerpoint()
else:
    print("PowerPoint is not running")

while True:
    appget_pid_powerpoint()

    time.sleep(1)  # Can only update rich presence every 1 second

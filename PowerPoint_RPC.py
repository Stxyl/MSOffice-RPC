import os
import pypresence


# This function is used to connect to Discord
def connect_discord():
    RPC = (pypresence.Presence("1014791041838219274"))
    RPC.connect()
    return RPC


# This function is used to update the presence
def update_presence_powerpoint():
    RPC = connect_discord()
    RPC.update(state="Editing a presentation", large_image="powerpoint", large_text="Microsoft PowerPoint")
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


# This function finds the process ID of the application
def appget_pid_powerpoint():
    pid = os.popen("pgrep 'Microsoft PowerPoint'").read()
    return pid

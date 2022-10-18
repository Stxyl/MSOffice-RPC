import os
import pypresence


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


# This function finds the process ID of the application
def appget_pid_onenote():
    pid = os.popen("pgrep 'Microsoft OneNote'").read()
    return pid

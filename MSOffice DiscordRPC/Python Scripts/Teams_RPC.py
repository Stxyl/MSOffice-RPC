import os
import pypresence
import time


# This function is used to connect to Discord
def connect_discord():
    RPC = (pypresence.Presence("1014791567560691712"))
    RPC.connect()
    return RPC


# This function is used to update the presence
def update_presence_teams():
    RPC = connect_discord()
    RPC.update(state="Chatting", large_image="teams", large_text="Microsoft Teams")
    print("Presence updated")
    connect_discord()


# These functions find's where the application is stored and if its running
def appget_teams():
    if os.path.exists("/Applications/Microsoft teams.app/Contents/MacOS/Teams"):
        print("Located teams")
        return True
    else:
        print("Teams not found (You sure you located it properly?)")
        return False


appget_teams()


def appget_pid_teams():
    pid = os.popen("pgrep 'Teams'").read()
    return pid


# This function finds the process ID of the application and tells you if its running or not with PID Number
if appget_pid_teams() != "":
    print("Teams is running", appget_pid_teams())
    update_presence_teams()
else:
    print("Teams is not running")

while True:
    appget_pid_teams()

    time.sleep(1)  # Can only update rich presence every 1 second

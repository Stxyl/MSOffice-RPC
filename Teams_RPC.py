import os
import pypresence


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


def appget_pid_teams():
    pid = os.popen("pgrep 'Teams'").read()
    return pid

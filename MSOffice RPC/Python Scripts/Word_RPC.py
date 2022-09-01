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


# These functions find's where the application is stored and if its running
def appget_word():
    if os.path.exists("/Applications/Microsoft Word.app/Contents/MacOS/Microsoft Word"):
        print("Located Word")
        return True
    else:
        print("Word not found (You sure you located it properly?)")
        return False


appget_word()


def appget_pid_word():
    pid = os.popen("pgrep 'Microsoft Word'").read()
    return pid


# This function finds the process ID of the application and tells you if its running or not with PID Number
if appget_pid_word() != "":
    print("Word is running", appget_pid_word())
    update_presence_word()
else:
    print("Word is not running")

while True:
    appget_pid_word()

    time.sleep(1)  # Can only update rich presence every 1 second

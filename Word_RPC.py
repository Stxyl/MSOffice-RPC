import os
import pypresence


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


def appget_pid_word():
    pid = os.popen("pgrep 'Microsoft Word'").read()
    return pid

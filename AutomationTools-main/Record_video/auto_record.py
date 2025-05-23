import pyautogui
import subprocess
from time import sleep
import win32gui
import win32con
import ctypes
from os.path import abspath, join, split
from pathlib import Path
import threading, win32process, psutil
import tkinter

Dir = "D:\Game"

# ファイル選択ダイアログの表示
root = tkinter.Tk()
root.withdraw()
fTyp = [("","*")]
#iDir = os.path.abspath(os.path.dirname(__file__))
tkinter.messagebox.showinfo('自動録画用tool', '録画するゲームを指定してください')
game_path = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = Dir)
print(game_path)

# 選択したexeファイルのあるフォルダ名の取得
#game_fold_name = split(abspath(join(game_path, '../..')))[-1]
game_fold_name = Path(game_path).parent.name
print(game_fold_name)



if "製品版" in game_fold_name:
    game_title = game_fold_name.replace("製品版", "")
elif "ver" in game_fold_name:
    game_title = game_fold_name.replace("ver", "") #特定の文字以降消す
else:
    game_title = game_fold_name

print(game_title)

    
def get_app_forground_name(hwnd, extra):
    if win32gui.IsWindowVisible(hwnd):
        if extra in win32gui.GetWindowText(hwnd):
            win32gui.SetForegroundWindow(hwnd)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)


#録画コマンド(Ctrl+R)
def record():
    #ウインドウをアクティブにする。
    pyautogui.click(1366,401)
    pyautogui.keyDown("ctrl")
    pyautogui.keyDown("r")
    pyautogui.keyUp("r")
    pyautogui.keyUp("ctrl")
    
# ゲームタイトルと同じ名前が入っていれば、それを開く

# 一番上で開いているアプリのタイトル取得

# process名から、欲しいアプリのWindowをアクティブにする?
def active_window_process_name():
    pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow()) #This produces a list of PIDs active window relates to
    print(psutil.Process(pid[-1]).name()) #pid[-1] is the most likely to survive last longer


if __name__ == '__main__':
    subprocess.run(game_path)  # Start the game executable
    sleep(5)  # Wait for the game window to appear
    
    print(game_title)
    sleep(3) #click on a window you like and wait 3 seconds 
    active_window_process_name()

    t = threading.Thread(target = to_label)
    t.start()
    root.mainloop()

    windowTile = ""
    while True:
        newWindowTile = win32gui.GetWindowText (win32gui.GetForegroundWindow())
        if newWindowTile != windowTile:
            windowTile = newWindowTile 
            print( windowTile )
            break
            
    record() # recording command twice
    record()
    sleep(5)
    pyautogui.moveTo(919, 400, 2)
    pyautogui.click(button="left", clicks=300, interval=5) # click automatically
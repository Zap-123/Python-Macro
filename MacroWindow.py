import math
import time
from tkinter import *

import keyboard
import pyautogui
import win32api
import win32con

# +--------------------+
#       Vars
# +--------------------+


MainWin = Tk()
MainWin.geometry("300x280")
MainWin.title("Macro")

MacroMoves = []
pyautogui.FAILSAFE = False

GlobStartButton = '+'
GlobStopButton = '-'
GlobOpenKey = '|'
GlobCopyPaste = ''
GlobTimeBetweenIterations = 0
GlobTypingTime = 0
rewrite = 1
CPS = 6.0 * 0.93


# +--------------------+
# Decorative Functions
# +--------------------+


def DecorateMain():
    def Space(Row):
        Spacing = Label(MainWin, text="", bg=BGC)
        Spacing.grid(row=Row)

    BGC = "#99c9ff"
    MainWin.config(bg=BGC)

    Title = Label(MainWin, text="Macro GUI", font=("comfortaa", 20), fg="Dark Green", bg=BGC)
    Title.grid(row=0)

    LB1 = Label(MainWin, text="╔═══━━━━──── • ────━━━━═══╗", fg="red", bg=BGC)
    LB1.grid(row=1)

    Space(2)

    OpenMouseWin = Button(MainWin, text="Open Mouse Macros", borderwidth=4, fg="Yellow", bg="Blue",
                          font=("comfortaa", 10), command=OpenMouseWindow)
    OpenMouseWin.grid(row=3)

    Space(4)

    OpenTypingWin = Button(MainWin, text="Open Keyboard Macros", borderwidth=4, fg="Yellow", bg="Red",
                           font=("comfortaa", 10), command=OpenKeyboardWindow)
    OpenTypingWin.grid(row=5)

    Space(6)

    OpenConfigWin = Button(MainWin, text="Open Settings", width=13, borderwidth=4, fg="Yellow", bg="Green",
                           font=("Confortaa", 10), command=OpenConfigWindow)
    OpenConfigWin.grid(row=7)

    Space(8)

    LB2 = Label(MainWin, text="════━━━━──── • ────━━━━════", fg="red", bg=BGC)
    LB2.grid(row=9)


def OpenMouseWindow():
    BGC = "#5d96da"
    MouseGUI = Toplevel()
    MouseGUI.geometry("300x510")
    MouseGUI.title("MouseGUI")
    MouseGUI.config(bg=BGC)

    def Space(Row):
        Spacing = Label(MouseGUI, text="", bg=BGC)
        Spacing.grid(row=Row)

    def FindMouse():
        time.sleep(1.5)

        XEntry.delete(0, END)
        YEntry.delete(0, END)

        MouseX, MouseY = pyautogui.position()
        MousePos = Label(MouseGUI, text=f"Your mouse is at ({MouseX}, {MouseY})", fg="Yellow", bg=BGC,
                         font=("comfortaa", 10))
        MousePos.grid(row=4, column=0)

        XEntry.insert(0, str(MouseX))
        YEntry.insert(0, str(MouseY))

    def AddLeftClick():
        global MacroMoves
        MacroMoves.append("left")
        print(MacroMoves)

    def AddRightClick():
        global MacroMoves
        MacroMoves.append("right")
        print(MacroMoves)

    def Submit():
        global MacroMoves
        global rewrite
        global CPS
        rewrite = 1

        if XEntry.get() != "X Pos here":
            if XEntry.get() != '':
                if YEntry.get() != "Y Pos here":
                    if YEntry.get() != '':
                        MacroMoves.append(f"{XEntry.get()} {YEntry.get()}")
                        XEntry.delete(0, END)
                        YEntry.delete(0, END)

        if WaitTimeEntry.get() != '':
            if float(WaitTimeEntry.get()) != 0:
                MacroMoves.append(f"|||{WaitTimeEntry.get()}|")
                WaitTimeEntry.delete(0, END)

        if CPSEntry.get() != '':
            if CPSEntry.get() != 'CPS':
                CPS = float(CPSEntry.get()) * 0.93
                CPSEntry.delete(0, END)
                print(CPS)

        print(MacroMoves)

    def CloseWin():
        MouseGUI.destroy()

    def Reset():
        global MacroMoves
        MacroMoves.clear()

    MouseTitle = Label(MouseGUI, text="Mouse GUI", font=("comfortaa", 15), fg="#84c33a", bg=BGC)
    MouseTitle.grid(row=0)

    LB = Label(MouseGUI, text="╔═══━━━━──── • ────━━━━═══╗", fg="red", bg=BGC)
    LB.grid(row=1)

    Space(2)

    FindMouseButton = Button(MouseGUI, text="Find Mouse", borderwidth=4, bg="red", fg="yellow", command=FindMouse)
    FindMouseButton.grid(row=3)

    Space(4)

    XEntry = Entry(MouseGUI, width=20, borderwidth=3, bg="Yellow", fg="Blue")
    XEntry.grid(row=5, column=0)

    XEntry.insert(0, "X Pos here")

    YEntry = Entry(MouseGUI, width=20, borderwidth=3, bg="Blue", fg="Yellow")
    YEntry.grid(row=6, column=0)

    YEntry.insert(0, "Y Pos here")

    Space(7)

    LCButton = Button(MouseGUI, text="Left Click", command=AddLeftClick, bg="Orange", fg="Purple", borderwidth=4)
    LCButton.grid(row=8)

    RCButton = Button(MouseGUI, text="Right Click", command=AddRightClick, bg="Purple", fg="Orange", borderwidth=4)
    RCButton.grid(row=9)

    Space(10)

    CPSEntry = Entry(MouseGUI, width=20, bg="green", fg="red", borderwidth=3)
    CPSEntry.grid(row=11)
    CPSEntry.insert(0, "CPS")

    Space(12)

    LB4 = Label(MouseGUI, text="Add a wait time", bg=BGC, fg="yellow")
    LB4.grid(row=13)

    WaitTimeEntry = Entry(MouseGUI, width=20, bg="yellow", fg="blue", borderwidth=3)
    WaitTimeEntry.grid(row=14)

    Space(15)

    SubmitButton = Button(MouseGUI, text="Submit", bg="Green", fg="red", borderwidth=4, command=Submit)
    SubmitButton.grid(row=16)

    CloseButton = Button(MouseGUI, text="Close/Done", bg="Red", fg="green", borderwidth=4, command=CloseWin)
    CloseButton.grid(row=17)

    ResetButton = Button(MouseGUI, text="Reset All", bg="Orange", fg="Dark Red", borderwidth=4, command=Reset)
    ResetButton.grid(row=18)

    Space(19)

    LB = Label(MouseGUI, text="════━━━━──── • ────━━━━════", fg="red", bg=BGC)
    LB.grid(row=20)


def OpenKeyboardWindow():
    BGC = "#ef4c4f"
    KeyboardGUI = Toplevel()
    KeyboardGUI.geometry("300x470")
    KeyboardGUI.title("KeyboardGUI")
    KeyboardGUI.config(bg=BGC)

    def Space(Row):
        Spacing = Label(KeyboardGUI, text="", bg=BGC)
        Spacing.grid(row=Row)

    def Submit():
        global MacroMoves
        global GlobTypingTime
        global rewrite

        rewrite = 1
        Keypress = PressEntry.get()
        Letter = TypeEntry.get()
        TimeDiff = TimeBetweenTypingEntry.get()

        if Letter != '':
            MacroMoves.append(f"|{Letter}")
            TypeEntry.delete(0, END)
        if Keypress != '':
            MacroMoves.append(f"||{Keypress}")
            PressEntry.delete(0, END)
        if TimeDiff != '':
            GlobTypingTime = TimeDiff
            TimeBetweenTypingEntry.delete(0, END)
        if WaitTimeEntry.get() != '':
            if float(WaitTimeEntry.get()) != 0:
                MacroMoves.append(f"|||{WaitTimeEntry.get()}|")

        print(MacroMoves)

    def CloseWin():
        KeyboardGUI.destroy()

    def Reset():
        global MacroMoves
        MacroMoves.clear()

    KeyboardTitle = Label(KeyboardGUI, text="Keyboard GUI", font=("comfortaa", 15), fg="#0d88a9", bg=BGC)
    KeyboardTitle.grid(row=0)

    LB = Label(KeyboardGUI, text="╔═══━━━━──── • ────━━━━═══╗", fg="#003CFF", bg=BGC)
    LB.grid(row=1)

    Space(2)

    LB1 = Label(KeyboardGUI, text="Type words here to auto type", fg="Orange", bg=BGC)
    LB1.grid(row=3)

    TypeEntry = Entry(KeyboardGUI, width=20, bg="Orange", fg="blue", borderwidth=3)
    TypeEntry.grid(row=4)

    Space(5)

    LB2 = Label(KeyboardGUI, text="Type keys to auto press down", fg="Green", bg=BGC)
    LB2.grid(row=6)

    PressEntry = Entry(KeyboardGUI, width=20, bg="green", fg="orange", borderwidth=3)
    PressEntry.grid(row=7)

    Space(8)

    LB3 = Label(KeyboardGUI, text="Type seconds between typing here", fg="Blue", bg=BGC)
    LB3.grid(row=9)

    TimeBetweenTypingEntry = Entry(KeyboardGUI, width=20, bg="Blue", fg="Yellow", borderwidth=3)
    TimeBetweenTypingEntry.grid(row=10)

    Space(11)

    LB4 = Label(KeyboardGUI, text="Add a wait time", bg=BGC, fg="yellow")
    LB4.grid(row=12)

    WaitTimeEntry = Entry(KeyboardGUI, width=20, bg="yellow", fg="blue", borderwidth=3)
    WaitTimeEntry.grid(row=13)

    Space(14)

    SubmitButton = Button(KeyboardGUI, text="Submit", bg="Green", fg="red", borderwidth=4, command=Submit)
    SubmitButton.grid(row=15)

    CloseButton = Button(KeyboardGUI, text="Close/Done", bg="Red", fg="green", borderwidth=4, command=CloseWin)
    CloseButton.grid(row=16)

    ResetButton = Button(KeyboardGUI, text="Reset All", bg="Orange", fg="Dark Red", borderwidth=4, command=Reset)
    ResetButton.grid(row=17)

    Space(18)

    LB = Label(KeyboardGUI, text="════━━━━──── • ────━━━━════", fg="#003CFF", bg=BGC)
    LB.grid(row=19)


def OpenConfigWindow():
    BGC = "#009900"
    ConfigGUI = Toplevel()
    ConfigGUI.geometry("300x550")
    ConfigGUI.title("ConfigGUI")
    ConfigGUI.config(bg=BGC)

    def Space(Row):
        Spacing = Label(ConfigGUI, text="", bg=BGC)
        Spacing.grid(row=Row)

    def Submit():
        global MacroMoves
        global GlobTimeBetweenIterations
        global GlobStopButton
        global GlobStartButton
        global GlobOpenKey
        global GlobCopyPaste
        global rewrite
        rewrite = 1
        if StartKey.get() != '':
            if StartKey.get() != GlobStartButton:
                GlobStartButton = StartKey.get()

        if StopKey.get() != '':
            if StartKey.get() != GlobStopButton:
                GlobStopButton = StopKey.get()

        if TimeBetweenIterations.get() != '':
            GlobTimeBetweenIterations = TimeBetweenIterations.get()

        if OpenKey.get() != '':
            if OpenKey.get() != GlobOpenKey:
                GlobOpenKey = OpenKey.get()

        if CopyPasteEntry.get() != '':
            GlobCopyPaste = CopyPasteEntry.get()

        print(MacroMoves)

    def CloseWin():
        ConfigGUI.destroy()

    def Reset():
        global MacroMoves
        MacroMoves.clear()

    KeyboardTitle = Label(ConfigGUI, text="Settings", font=("comfortaa", 15), fg="Orange", bg=BGC)
    KeyboardTitle.grid(row=0)

    LB = Label(ConfigGUI, text="╔═══━━━━──── • ────━━━━═══╗", fg="#ff3399", bg=BGC)
    LB.grid(row=1)

    LB2 = Label(ConfigGUI, text="Start & Stop Keys", bg=BGC, fg="Black", font=("comfortaa", 13))
    LB2.grid(row=7)

    StartKey = Entry(ConfigGUI, width=2, bg="green", fg="Yellow", borderwidth=3)
    StartKey.grid(row=8)

    StopKey = Entry(ConfigGUI, width=2, bg="red", fg="Yellow", borderwidth=3)
    StopKey.grid(row=9)

    Space(10)

    LB3 = Label(ConfigGUI, text="Re-open The Program", bg=BGC, fg="black", font=("comfortaa", 13))
    LB3.grid(row=11)

    OpenKey = Entry(ConfigGUI, width=2, bg="Blue", fg="Yellow")
    OpenKey.grid(row=12)

    StartKey.insert(0, GlobStartButton)
    StopKey.insert(0, GlobStopButton)
    OpenKey.insert(0, GlobOpenKey)

    Space(13)

    LB4 = Label(ConfigGUI, text="Seconds Between Iterations", bg=BGC, fg="black", font=("comfortaa", 13))
    LB4.grid(row=14)

    TimeBetweenIterations = Entry(ConfigGUI, width=10, bg="Yellow", fg="Blue", borderwidth=3)
    TimeBetweenIterations.grid(row=15)
    TimeBetweenIterations.insert(0, str(GlobTimeBetweenIterations))

    Space(16)

    LB5 = Label(ConfigGUI, text="Copy & Paste", bg=BGC, fg="black", font=("comfortaa", 13))
    LB5.grid(row=17)

    CopyPasteEntry = Entry(ConfigGUI, width=10, borderwidth=3, bg="purple", fg="orange")
    CopyPasteEntry.grid(row=18)

    Space(19)

    LB6 = Label(ConfigGUI, text="Code to copy vvv", bg=BGC, fg="black", font=("comfortaa", 13))
    LB6.grid(row=20)

    CopyMacro = Entry(ConfigGUI, width=20, borderwidth=3, bg="orange", fg="purple")
    CopyMacro.grid(row=21)
    CopyMacro.insert(0, str(MacroMoves))

    Space(22)

    SubmitButton = Button(ConfigGUI, text="Submit", bg="Green", fg="red", borderwidth=4, command=Submit)
    SubmitButton.grid(row=23)

    CloseButton = Button(ConfigGUI, text="Close/Done", bg="Red", fg="green", borderwidth=4, command=CloseWin)
    CloseButton.grid(row=24)

    ResetButton = Button(ConfigGUI, text="Reset All", bg="Orange", fg="Dark Red", borderwidth=4, command=Reset)
    ResetButton.grid(row=25)

    Space(26)

    LB = Label(ConfigGUI, text="════━━━━──── • ────━━━━════", fg="#ff3399", bg=BGC)
    LB.grid(row=27)


# +--------------------+
#       Macro >:)
# +--------------------+

def LeftClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(1 / (CPS * (4 / 3)))


def RightClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)
    time.sleep(1 / (CPS * (4 / 3)))


def GoTo(x, y):
    win32api.SetCursorPos((x, y))


def StringToMacList(string):
    global MacroMoves
    MacroMoves = []
    closed = 1
    wordchunk = ''
    for i in string:
        if i == '\'':
            closed += 1
            if closed == 2:
                closed = 0
            if closed == 1:
                if i != ',':
                    if i != ' ':
                        MacroMoves.append(wordchunk)
                        wordchunk = ''

        if closed == 0:
            if i != '\'':
                wordchunk += i



def Macro():
    global MacroMoves
    iterate = 0

    if GlobCopyPaste != '':
        StringToMacList(GlobCopyPaste)

    for i in MacroMoves:
        if int(GlobTimeBetweenIterations) > 0:
            time.sleep(int(GlobTimeBetweenIterations))

        if keyboard.is_pressed(f'{GlobStopButton}'):
            break

        if i[0] != '|':

            # left & right click

            if i == "left":
                LeftClick()
                iterate += 1
                continue

            elif i == "right":
                RightClick()
                iterate += 1
                continue

            # Mos Pos

            else:
                PlaceHolder = 0
                x = ""
                y = ""

                for i2 in i:
                    if i2 == ' ':
                        PlaceHolder += 1

                    else:
                        if PlaceHolder == 0:
                            x = str(x) + str(i2)

                        else:
                            y = str(y) + str(i2)

                x = int(x)
                y = int(y)

                GoTo(x, y)
                iterate += 1
                continue

        # Typing

        if keyboard.is_pressed(f'{GlobStopButton}'):
            break

        if i[0] == '|':
            if i[1] != '|':
                write = i[1:-1]
                if GlobTypingTime == '':
                    time1 = 0
                else:
                    time1 = float(GlobTypingTime)

                if time1 != 0:
                    for a in write:
                        pyautogui.write(a)
                        time.sleep(time1)

                        if keyboard.is_pressed(f'{GlobStopButton}'):
                            break
                    iterate += 1
                    continue
                else:
                    pyautogui.write(write)
                    iterate += 1
                    continue

        if i[0] == '|':
            if i[2] == '|':
                if i[-1] == '|':
                    sleep = ''
                    for i2 in i:
                        if i2 != '|':
                            sleep = sleep + str(i2)

                    for i4 in range(math.floor(float(sleep)) * 4):
                        time.sleep(0.25)
                        if keyboard.is_pressed(f'{GlobStopButton}'):
                            break

                    time.sleep(float(sleep) - math.floor(float(sleep)))

                    if keyboard.is_pressed(f'{GlobStopButton}'):
                        break

                    iterate += 1
                    continue

        # Pressing

        if str(i[0]) == '|':
            if str(i[1]) == '|':
                if str(i[2]) != '|':
                    first = f"{i[2:-1]}" + f"{i[-1]}"
                    pyautogui.press(first)
                    iterate += 1
                    continue

        iterate += 1


# +--------------------+
#       Startup
# +--------------------+


DecorateMain()
MainWin.mainloop()

while True:
    if rewrite == 1:
        if GlobCopyPaste != '':
            MacroMoves = []
            in_out = 0
            append = ''
            for i3 in GlobCopyPaste:
                if i3 == '\'':
                    if in_out == 0:
                        in_out = 1
                    else:
                        in_out = 0
                        MacroMoves.append(append)
                        append = ''

                if in_out == 1:
                    if i3 != '\'':
                        append = append + str(i3)

        rewrite = 0

    if keyboard.is_pressed(GlobOpenKey):
        MainWin = Tk()
        MainWin.geometry("300x280")
        MainWin.title("Macro")

        DecorateMain()
        rewrite = 0
        MainWin.mainloop()

    if keyboard.is_pressed(f'{GlobStartButton}'):
        while True:
            if keyboard.is_pressed(f'{GlobStopButton}'):
                break

            Macro()

            if keyboard.is_pressed(f'{GlobStopButton}'):
                break

            if keyboard.is_pressed(GlobOpenKey):
                MainWin = Tk()
                MainWin.geometry("300x280")
                MainWin.title("Macro")

                DecorateMain()
                rewrite = 0
                MainWin.mainloop()
                break

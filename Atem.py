import PyATEMMax
from pynput.keyboard import Key, Listener, KeyCode
from pynput import keyboard
import win32gui, win32process, os
import sys, time, threading
from colored import fg

PVW_color = fg('green_1')
PGM_color = fg('red_1')
sColor = fg('white')

cmd = 'mode 24,9'
os.system(cmd)

atemMini = PyATEMMax.ATEMMax()
atem4K = PyATEMMax.ATEMMax()

import tkinter as tk
from tkinter import messagebox


def exit_application():
    msg_box = tk.messagebox.askquestion('Exit', 'Are you sure you want to exit?',
                                        icon='question')
    if msg_box == 'yes':
        atem4K.disconnect()
        atemMini.disconnect()
        os.system("cls")
        print("Disconnected")
        exit()
        quit()
        return False  # stop listener
    # else:
    #     tk.messagebox.showinfo('Return', 'You will now return to the application screen')


def restart_application():
    global stop_flag
    msg_box = tk.messagebox.askquestion('Restart', 'Are you sure you want to restart the application?',
                                        icon='question')
    if msg_box == 'yes':
        stop_flag = True
        atem4K.disconnect()
        atemMini.disconnect()
        os.startfile(__file__)
        quit()
        # os.execl(sys.executable, '"{}"'.format(sys.executable), *sys.argv)
    # else:
    #     atem4K.disconnect()
    #     atemMini.disconnect()
    #     os.system("cls")
    #     print("Disconnected")
    #     exit()
    #     quit()
    #     return False  # stop listener


def connectionFailed(key):
    try:
        k = key.char  # single-char keys
    except:
        k = key.name  # other keys
    if k == 'r':  # restart_application
        restart_application()
        # os.execl(sys.executable, '"{}"'.format(sys.executable), *sys.argv)
    if key == keyboard.Key.esc:
        exit_application()


def connectToAtem():
    i = 0
    while not atem4K.connected and not atemMini.connected and i < 1:
        atemMini.connect("192.168.1.223")
        atemMini.waitForConnection(infinite=False)

        atem4K.connect("192.168.1.221")
        atem4K.waitForConnection(infinite=False)
        i += 1


def loadingAnimation(process):
    while process.is_alive():
        chars = ['', '.', '..', '...', '   ']
        for char in chars:
            sys.stdout.write('\r' + 'Connecting' + char)
            time.sleep(.2)
            sys.stdout.flush()


loading_process = threading.Thread(target=connectToAtem)
loading_process.start()

loadingAnimation(loading_process)
loading_process.join()

pressed = False
mode = "8 CAMS"
sugo = "M/E 3 PGM"  # M/E 3 PGM / M/E 1 PGM #atem4K.auxSource[2].input.value
lock = "Locked"
pc_vid = "OFF"
CurrentME = "ME1"


lockColor = fg('grey_19') # grey_23 egész jó #grey_0 túl halvány #

def consoleText():
    KeyerDisplay()
    os.system('cls')

    print(f"Functions: ", end="")
    if lock == "Locked":
        print(PGM_color, end="")
    else:
        print(PVW_color, end="")
    print(lock)

    if lock == "Locked":
        print(lockColor, end="")
    else:
        print(sColor,end="")
    print(f" - SÚGÓ: {sugo}")
    print(f" - PC_VID: {pc_vid}")
    #print(f" - {CurrentME}:  {keysDisplay}")

    #print(f"MODE: {mode}")
    print("\n"+PVW_color, end="")
    print(f"PVW: ", end="")
    print(sColor, end="")
    print(f"{PVW}")
    print(PGM_color, end="")
    print(f"PGM: ", end="")
    print(sColor, end="")
    print(f"{PGM}")



def setPGMandPVWtoCurrent():
    global PGM
    global PVW
    global last_PGW_MINI
    last_PGW_MINI = atemMini.programInput[0].videoSource.value
    if PVW == 5 and atemMini.connected:
        PVW = 4 + last_PGW_MINI
    if PGM == 5 and atemMini.connected:
        PGM = 4 + last_PGW_MINI

stop_flag = False
def PGM_and_PVW_has_chaned():
    global PVW
    global PGM
    while not stop_flag:
        last_PGM = atem4K.programInput[1].videoSource.value
        last_PVW = atem4K.previewInput[1].videoSource.value

        # if last_PGM != PGM:
        #     PGM = last_PGM
        #     setPGMandPVWtoCurrent()
        # if last_PVW != PVW:
        #     PVW = last_PVW
        #     setPGMandPVWtoCurrent()

        if last_PVW != PVW or last_PGM != PGM:
            PVW = last_PVW
            PGM = last_PGM
            setPGMandPVWtoCurrent()
            consoleText()

        #print(11111)
        time.sleep(0.01)

PGM_PVW_Listener = threading.Thread(target=PGM_and_PVW_has_chaned)



me1A = "O"
me1B = "O"
me1C = "X"
me2A = "X"
me2B = "X"
me2C = "O"
me3A = "X"
me3B = "X"
me3C = "O"

def KeyerDisplay():
    global keysDisplay
    if CurrentME == "ME1":
        keysDisplay = f"{me1A}  {me1B}  {me1C}"
    elif CurrentME == "ME2":
        keysDisplay = f"{me2A}  {me2B}  {me2C}"
    elif CurrentME == "ME3":
        keysDisplay = f"{me3A}  {me3B}  {me3C}"







def on_press(key):
    global lock
    global stop_flag
    global mode
    global sugo
    global pressed
    global PVW
    global PGM
    global pc_vid

    focus_window_pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[1]
    current_process_pid = os.getppid()

    if focus_window_pid == current_process_pid:
        if key == keyboard.Key.esc:
            stop_flag = True
            exit_application()
            # atem4K.disconnect()
            # atemMini.disconnect()
            # os.system("cls")
            # print("Disconnected")
            # return False  # stop listener
        try:
            k = key.char  # single-char keys
        except:
            k = key.name  # other keys
        if k in ['1', '2', '3', '4'] and k != PVW and k != PGM and pressed == False:  # keys of interest #ATEM4k
            # self.keys.append(k)  # store it in global-like variable
            pressed = True
            kInt = int(k)
            # print(type(kInt))
            atem4K.setPreviewInputVideoSource(1, kInt)  # set PVW on Atem4K M/E 2

            #PVW = k
            PVW = atem4K.previewInput[1].videoSource.value  # k

            setPGMandPVWtoCurrent()
            consoleText()

        if k in ['5', '6', '7', '8'] and k != PVW and k != PGM and mode == "8 CAMS" and pressed == False:  # ATEMmini

            pressed = True
            kInt = int(k) - 4
            atemMini.setProgramInputVideoSource(0, kInt)  # set PGM on AtemMini M/E 1
            atem4K.setPreviewInputVideoSource(1, 5)  # set PVW on Atem4K M/E 2

            #PVW = k
            PVW = atemMini.previewInput[1].videoSource.value

            setPGMandPVWtoCurrent()
            consoleText()

        while key == keyboard.Key.space and pressed == False:  ## and PVW != "-1" and PVW != "0": #CUT
            atem4K.execCutME(1)  # Cut on Atem4K M/E 2

            #temp = PGM
            #PGM = PVW
            #PVW = temp

            PGM = atem4K.programInput[1].videoSource.value
            PVW = atem4K.previewInput[1].videoSource.value

            setPGMandPVWtoCurrent()
            consoleText()


            print("CUT")
            pressed = True
        while key == keyboard.Key.enter and pressed == False:  ## and PVW != "0" and PVW != "-1": #FADE
            atem4K.execAutoME(1)

            # temp = PGM
            # PGM = PVW
            # PVW = temp

            PGM = atem4K.programInput[1].videoSource.value
            PVW = atem4K.previewInput[1].videoSource.value  # "-1" #temp

            setPGMandPVWtoCurrent()
            consoleText()


            print("FADE")
            pressed = True

        # MP1 (Hamarosan kezdünk)
        if k in ['0'] and k != PVW and k != PGM and pressed == False and lock == "Unlocked":
            pressed = True
            kInt = '3010'
            # print(type(kInt))
            atem4K.setPreviewInputVideoSource(1, 3010)  # set PVW on Atem4K M/E 2       atem4K.atem.videoSources.mediaPlayer1.value

            PVW = k
            #PVW = atem4K.previewInput[1].videoSource.value  # k
            setPGMandPVWtoCurrent()
            consoleText()




        # MP2 (Adásunk véget ért!)
        if k in ['9'] and k != PVW and k != PGM and pressed == False and lock == "Unlocked":
            pressed = True
            kInt = '3020'
            # print(type(kInt))
            atem4K.setPreviewInputVideoSource(1, atem4K.atem.videoSources.mediaPlayer2.value)  # set PVW on Atem4K M/E 2

            PVW = k
            #PVW = atem4K.previewInput[1].videoSource.value  # k
            setPGMandPVWtoCurrent()
            consoleText()




        # LOCK
        if key == Key.end and pressed == False:
            if lock == "Locked":
                lock = "Unlocked"
            else:
                lock = "Locked"

            setPGMandPVWtoCurrent()
            consoleText()
            pressed = True


        # Súgó
        if key == Key.home and pressed == False and lock == "Unlocked":  # k == "s" and pressed == False:
            if sugo == "M/E 3 PGM":  # and mode == "Prédikáció":
                sugo = "M/E 1 PGM"
                atem4K.setAuxSourceInput(1, "mE1Prog")
            elif sugo == "M/E 1 PGM":
                sugo = "M/E 3 PGM"
                atem4K.setAuxSourceInput(1, "mE3Prog")

            setPGMandPVWtoCurrent()
            consoleText()
            pressed = True



        # PC_VID KIVETÍTÉS
        if key == Key.page_up and pressed == False and lock == "Unlocked":
            if pc_vid == "OFF":
                pc_vid = "ON"
                atem4K.setPreviewInputVideoSource(1, 11) #ME2 Stream #2nd number is the number of PC_VID input
                atem4K.setPreviewInputVideoSource(2, 11) #ME3 Közönség
                atem4K.execAutoME(1)
            else:
                pc_vid = "OFF"
                atem4K.execAutoME(1)

            setPGMandPVWtoCurrent()
            consoleText()
            if pc_vid =="ON":
                print("PC_VID ON LIVE")
            pressed = True

        if k == 'r':  # restart_application
            restart_application()

        #Mode selector
        # if k == 'm' and pressed == False:  # key == Key.page_up and pressed == False:
        #     if mode == "8 CAMS":
        #         mode = "4 CAMS"
        #     else:
        #         mode = "8 CAMS"
        #
        #     setPGMandPVWtoCurrent()
        #     consoleText()
        #
        #     pressed = True


        # UPSTREAM KEYS
        # if key == keyboard.Key.up and pressed == False and lock == "Unlocked":
        #     pressed = True
        #     global CurrentME
        #
        #     if CurrentME == "ME1":
        #         CurrentME = "ME2"
        #     elif CurrentME == "ME2":
        #         CurrentME = "ME3"
        #     elif CurrentME == "ME3":
        #         CurrentME = "ME1"
        #     consoleText()
        #
        #
        #
        # if key == keyboard.Key.left and lock == "Unlocked" and pressed == False:
        #     pressed = True
        #     global me1A
        #     global me1A
        #     global me1B
        #     global me1C
        #     global me2A
        #     global me2B
        #     global me2C
        #     global me3A
        #     global me3B
        #     global me3C
        #     if CurrentME == "ME1":
        #         if me1A == "O":
        #             me1A = "X"
        #             atem4K.setKeyerOnAirEnabled(0, 0, 1)
        #         elif me1A == "X":
        #             me1A = "O"
        #             atem4K.setKeyerOnAirEnabled(0, 0, 0)
        #
        #     if CurrentME == "ME2":
        #         if me2A == "O":
        #             me2A = "X"
        #             atem4K.setKeyerOnAirEnabled(1, 0, 1)
        #         elif me2A == "X":
        #             me2A = "O"
        #             atem4K.setKeyerOnAirEnabled(1, 0, 0)
        #
        #     if CurrentME == "ME3":
        #         if me3A == "O":
        #             me3A = "X"
        #             atem4K.setKeyerOnAirEnabled(2, 0, 1)
        #         elif me3A == "X":
        #             me3A = "O"
        #             atem4K.setKeyerOnAirEnabled(2, 0, 0)
        #
        #     consoleText()
        # if key == keyboard.Key.down and lock == "Unlocked" and pressed == False:
        #     pressed = True
        #     if CurrentME == "ME1":
        #         if me1B == "O":
        #             me1B = "X"
        #             atem4K.setKeyerOnAirEnabled(0, 1, 1)
        #         elif me1B == "X":
        #             me1B = "O"
        #             atem4K.setKeyerOnAirEnabled(0, 1, 0)
        #
        #     if CurrentME == "ME2":
        #         if me2B == "O":
        #             me2B = "X"
        #             atem4K.setKeyerOnAirEnabled(1, 1, 1)
        #         elif me2B == "X":
        #             me2B = "O"
        #             atem4K.setKeyerOnAirEnabled(1, 1, 0)
        #
        #     if CurrentME == "ME3":
        #         if me3B == "O":
        #             me3B = "X"
        #             atem4K.setKeyerOnAirEnabled(2, 1, 1)
        #         elif me3B == "X":
        #             me3B = "O"
        #             atem4K.setKeyerOnAirEnabled(2, 1, 0)
        #     consoleText()
        #
        # if key == keyboard.Key.right and lock == "Unlocked" and pressed == False:
        #     pressed = True
        #     if CurrentME == "ME1":
        #         if me1C == "O":
        #             me1C = "X"
        #             atem4K.setKeyerOnAirEnabled(0, 2, 1)
        #         elif me1C == "X":
        #             me1C = "O"
        #             atem4K.setKeyerOnAirEnabled(0, 2, 0)
        #
        #     if CurrentME == "ME2":
        #         if me2C == "O":
        #             me2C = "X"
        #             atem4K.setKeyerOnAirEnabled(1, 2, 1)
        #         elif me2C == "X":
        #             me2C = "O"
        #             atem4K.setKeyerOnAirEnabled(1, 2, 0)
        #
        #     if CurrentME == "ME3":
        #         if me3C == "O":
        #             me3C = "X"
        #             atem4K.setKeyerOnAirEnabled(2, 2, 1)
        #         elif me3C == "X":
        #             me3C = "O"
        #             atem4K.setKeyerOnAirEnabled(2, 2, 0)
        #     consoleText()




PVW = atem4K.previewInput[1].videoSource.value  # "-1"
PGM = atem4K.programInput[1].videoSource.value  # "0"


def on_release(key):  # The function that's called when a key is released
    global pressed
    pressed = False


if atem4K.connected or atemMini.connected:
    setPGMandPVWtoCurrent()
    consoleText()
    PGM_PVW_Listener.start()
    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()
else:
    os.system('cls')
    print("Connection Failed")
    import ctypes  # An included library with Python install.

    # atem4K.disconnect()
    # atemMini.disconnect()
    # ctypes.windll.user32.MessageBoxW(0, "Failed to connect to Atem", "ERROR", 5)
    with Listener(on_press=connectionFailed, on_release=on_release) as listener:
        listener.join()
    # failedConnectRetry()


# Original
# print(f"MODE: {mode}\n")
# print(f"PVW: {PVW}")
# print(f"PGM: {PGM}")

# listener = keyboard.Listener(on_press=on_press)
# listener.start()  # start to listen on a separate thread
# listener.join()  # remove if main thread is polling self.keys

# Collect events until released
# with Listener(on_press=on_press) as listener:listener.join()

# 1 2 3 4 5 6 7 8 Keys
# 1 2 3 4 5 5 5 5 Atem4K PVW
# x x x x 1 2 3 4 AtemMini PGM

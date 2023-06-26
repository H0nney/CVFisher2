import os
import glob
import pywintypes
import win32gui
import win32ui
import win32con
import threading
import cv2 as cv
import numpy as np
import time
import queue
from pydirectinput import keyDown, keyUp, press, click
from keyboard import is_pressed
from playsound import playsound

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def loadNeedles():
    items = {}
    items_dir = './needles'
    data_path = os.path.join(items_dir,'*g') 
    files = glob.glob(data_path) 
    if files:
        for file in files: 
            basename = os.path.basename(file)
            index = os.path.splitext(basename)[0]
            img = cv.imread(file, cv.IMREAD_COLOR) 
            items[index] = img
        
        # Clear console
        os.system('cls')
        ##################
        print(bcolors.OKGREEN + 'Załadowano zrzuty liczb' + bcolors.ENDC)
        ##################
        return items
    else:
        # Clear console
        os.system('cls')
        ##################
        print(bcolors.FAIL + '######################## ERROR ########################' + bcolors.ENDC)
        print(bcolors.WARNING + 'Umieść wycięte screeny (png) liczb które wyskakują przy łowieniu.\nNie zapisuj wycinków w paincie (ważne)' + bcolors.ENDC)
        print(bcolors.FAIL + '######################## ERROR ########################' + bcolors.ENDC)
        ##################
        exit()

def prepareClient(client):
    press('ALT')
    win32gui.SetForegroundWindow(client)
    time.sleep(0.1)
    keyDown('f')
    time.sleep(1.5)
    keyUp('f')
    keyDown('r')
    time.sleep(1.5)
    keyUp('r')
    keyDown('f')
    time.sleep(0.3)
    keyUp('f')

def windowCapture(client, x, y):
    w = 100 # set this
    h = 90 # set this

    hwnd = client
    wDC = win32gui.GetWindowDC(hwnd)
    dcObj = win32ui.CreateDCFromHandle(wDC)
    cDC = dcObj.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(dcObj, w, h)
    cDC.SelectObject(bmp)
    cDC.BitBlt((0, 0), (w, h), dcObj, (x, y), win32con.SRCCOPY)

    signedIntsArray = bmp.GetBitmapBits(True)
    img = np.frombuffer(signedIntsArray, dtype='uint8')
    img.shape = (h, w, 4)

    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(bmp.GetHandle())

    img = img[...,:3]
    img = np.ascontiguousarray(img)

    return img

    
clickQueue = queue.Queue()
lockedClients = []

def queueWorker(clients: list, sequenceCount: int):
    global lockedClients
    
    i = 0
    j = 0
    throwList = queue.Queue()
    delay = np.random.uniform(0.02, 0.08)
    state = 0
    while True:
        match state:
            case 0:
                item = clickQueue.get()
                if not item is None:
                    win32gui.SetForegroundWindow(item['client'])
                    time.sleep(0.1+delay)
                    for j in range(int(item['count'])):
                        keyDown("space")
                        time.sleep(0.1+delay)
                        keyUp("space")
            
                    i += 1
                    
                if i == sequenceCount:
                    time.sleep(10)
                    lockedClients = []
                    state = 1
                    
            case 1:
                for client in clients:
                    win32gui.SetForegroundWindow(client)
                    time.sleep(0.1+delay)
                    keyDown("space")
                    time.sleep(0.1+delay)
                    
                state = 0
                i = 0
    
def lookForNumbers(clients: list, needles):
    global lockedClients0
    clientCount = len(clients)
    threading.Thread(target=queueWorker, args=(clients, clientCount), daemon=True).start()
    
    max_found_val = 0.0
    while not is_pressed('0'):
        for client in clients:
            capture = windowCapture(client, 350, 120)
            capture = np.array(capture)
            # capture = cv.cvtColor(capture, cv.COLOR_)
            # cv.imshow("Bot vision", capture)
            # cv.waitKey(1)


            # os.system('cls')
            pickedNeedle = None
            threshold = 0.40
            highestVal = 0.0
            for index, needle in needles.items():
                result = cv.matchTemplate(capture, needle, cv.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv.minMaxLoc(result)
                
                # print(f'{index} - {max_val}')
                if max_val > highestVal:
                    highestVal = max_val
                    pickedNeedle = index
            
            if highestVal > 0 and highestVal > threshold and client not in lockedClients:
                clickQueue.put({"client": client, "count": pickedNeedle})
                lockedClients.append(client)
                print(bcolors.OKBLUE + 'Klikam ', pickedNeedle , ' razy, pewność: ', highestVal, bcolors.ENDC)
                

def prepare(clients: list):
    os.system('cls')
    needles = loadNeedles()
    fisher_state = 0
    
    for client in clients:
        print(bcolors.OKGREEN + 'Przygotowuję klienta' + bcolors.ENDC)
        prepareClient(client)
        
    for client in clients:
        print(bcolors.OKGREEN + 'Zarzucam wędke' + bcolors.ENDC)
        win32gui.SetForegroundWindow(client)
        time.sleep(0.1)
        press('space')
        time.sleep(0.1)
        
    threading.Thread(target=lookForNumbers, args=(clients, needles)).start()

def showMenu():
    print('##############################################################################')
    print('Sterowanie:\nNUM7 - Dodaj klienta\nNUM8 - Usuń klienta\nNUM9 - Start\nNUM0 - PANIC Stop\n')

def main():
    main_run = True
    main_state = 0
    clients: list = []
    
    showMenu()
    while main_run:
        if is_pressed('0'):
            main_run = False
            
        match main_state:
            case 0:
                if is_pressed('7'):
                    window = win32gui.GetForegroundWindow()
                    if window not in clients:
                        clients.append(window)
                        os.system('cls')
                        showMenu()
                        print("Lista klientów:")
                        print(clients)
                        print(bcolors.OKGREEN + 'Dodano klienta' + bcolors.ENDC)

                if is_pressed('8'):
                    window = win32gui.GetForegroundWindow()
                    if window in clients:
                        clients.remove(window)
                        os.system('cls')
                        showMenu()
                        print("Lista klientów:")
                        print(clients)
                        print(bcolors.FAIL + 'Usunięto klienta' + bcolors.ENDC)

                if is_pressed('9'):
                    os.system('cls')
                    main_state = 1
                
            case 1:
                main_run = False
                prepare(clients)
            
if __name__ == '__main__':
    main()
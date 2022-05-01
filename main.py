import cv2
import time
import numpy as np
from cvzone.HandTrackingModule import HandDetector
import pyautogui
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# variables
wCam, hCam = 640, 480
pTime = 0
wScr, hScr = pyautogui.size()
detector = HandDetector(detectionCon=0.8, maxHands=1)
frameR = 150  # frame reduction , bigger the smaller rectangle less precise
smoothening = 5  # for smoothening value x y of cursor
previousLocationX, previousLocationY = 0, 0
currentLocationX, currentLocationY = 0, 0

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
# volume.GetMasterVolumeLevel()
volumeRange = volume.GetVolumeRange()

minVol = volumeRange[0]
maxVol = volumeRange[1]

print('screen size', wScr, hScr)
print('camera size', wCam, hCam)
print('smoothening', smoothening)

# camera setup
cam = cv2.VideoCapture(0)
cam.set(3, wCam)
cam.set(4, hCam)

while True:
    # import images
    success, img = cam.read()
    # find the hand in image
    img = detector.findHands(img)
    lmList, _ = detector.findPosition(img)
    # check if we have hand
    if len(lmList) != 0:
        # which finger is up
        fingers = detector.fingersUp()
        cv2.rectangle(img, (frameR, frameR), (wCam - frameR, hCam - frameR), (255, 0, 255), 1)
        # 1 Only index finger for moving the cursor
        if fingers[1] == 1 and fingers[2] == 0 and fingers[0] == 0:
            # convert coordinates
            x1, y1 = lmList[8]  # get index finger coordination
            x3 = np.interp(x1, (frameR, wCam - frameR), (0, wScr))
            y3 = np.interp(y1, (frameR, hCam - frameR), (0, hScr))
            # smoothen the values
            currentLocationX = previousLocationX + (x3 - previousLocationX) / smoothening
            currentLocationY = previousLocationY + (y3 - previousLocationY) / smoothening
            # move mouse
            pyautogui.moveTo(wScr - currentLocationX, currentLocationY)
            # update values
            previousLocationX, previousLocationY = currentLocationX, currentLocationY
        # 2 Both index and middle are up  : click mode
        if fingers[1] == 1 and fingers[2] == 1:
            length, img, lineInfo = detector.findDistance(8, 12, img)
            if length < 40:
                pyautogui.click()
        # 3 thumb and index finger : Adjust volume
        if fingers[1] == 1 and fingers[0] == 1:
            x1, y1 = lmList[8]  # get index finger coordination
            xT, yT = lmList[4]  # thumb finger coordination
            lengthV, img, lineInfo = detector.findDistance(8, 4, img)
            vol = np.interp(lengthV, [10, 160],
                            [minVol, maxVol])  # max and min distance between thumb and fingers is [10 , 160]
            volume.SetMasterVolumeLevel(vol, None)
    # frame rate fps
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, str(int(fps)), (20, 50), cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 0), 3)
    # display
    cv2.imshow("image", img)
    cv2.waitKey(1)

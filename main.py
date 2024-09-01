import os
import cv2
from cvzone.HandTrackingModule import HandDetector
import win32com.client
import time
 
#variables

width=640
height=480

#camera setup
cap=cv2.VideoCapture(0)
cap.set(3,width)
cap.set(4,height)

#powerpoint access

powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = True


# Open the PowerPoint file
presentation = powerpoint.Presentations.Open(r'MINI PROJECT PRESENTATION.ppt', WithWindow=True)


presentation.SlideShowSettings.Run()

#hand detector

detector=HandDetector(detectionCon=0.7,maxHands=1)



while True:
    
    #accessing web cam as well as slides
    
    success, img = cap.read()
    img=cv2.flip(img,1)
    hands, img =detector.findHands(img)
    #hand=hands[0]


    if hands :
        hand=hands[0]
        fingers = detector.fingersUp(hand) # how many fingers are up and gives o/p as list
        print(fingers)
        cx,cy=hand['center']


        if fingers == [1,0,0,0,0] :     # gesture to go backward
            print("Left")
            buttonpress=True
            presentation.SlideShowWindow.View.Previous()
            time.sleep(1)

        if fingers == [0,0,0,0,1] :     # gesture to go forward
            print("Right")
            buttonpress=True
            presentation.SlideShowWindow.View.Next()
            time.sleep(1)

        if fingers == [1,1,0,0,1]:     #gestures to stop the presentation
            print("close")
            buttonpress=True
            presentation.Close()
            time.sleep(1)
        if fingers == [0,0,0,1,1]:
            print("last slide")
            presentation.SlideShowWindow.View.Last()
            time.sleep(2)

        if fingers == [0,1,0,0,0]:
            print("firstslide")
            presentation.SlideShowWindow.View.First()
            time.sleep(2)
    
    cv2.imshow("Image",img)
    key=cv2.waitKey(1)
    if(key==ord('e')):
        break

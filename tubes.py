import sys
import speech_recognition as speech
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import cv2
import mediapipe as mp
from math import hypot
import numpy as np


# define function

cap = cv2.VideoCapture(0) #Checks for camera

mpHands = mp.solutions.hands #detects hand/finger
hands = mpHands.Hands() #complete the initialization configuration of hands
mpDraw = mp.solutions.drawing_utils   #declaration landmark (hand)  
 
#To access speaker through the library pycaw 
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volbar=400
volper=0
 
volMin,volMax = volume.GetVolumeRange()[:2]
 
while True:
    success,img = cap.read() #If camera works capture an image
    imgRGB = cv2.cvtColor(img,cv2.COLOR_BGR2RGB) #Convert to rgb
    
    #Collection of gesture information
    results = hands.process(imgRGB) #completes the image processing.
 
    lmList = [] #empty list
    if results.multi_hand_landmarks: #list of all hands detected.
        #By accessing the list, we can get the information of each hand's corresponding flag bit
        for handlandmark in results.multi_hand_landmarks:
            for id,lm in enumerate(handlandmark.landmark): #adding counter and returning it
                # Get finger joint points
                h,w,_ = img.shape
                cx,cy = int(lm.x*w),int(lm.y*h)
                lmList.append([id,cx,cy]) #adding to the empty list 'lmList'
            mpDraw.draw_landmarks(img,handlandmark,mpHands.HAND_CONNECTIONS)
    
    if lmList != []:
        #getting the value at a point
                        #x      #y
        x1,y1 = lmList[4][1],lmList[4][2]  #thumb
        x2,y2 = lmList[8][1],lmList[8][2]  #index finger
        #creating circle at the tips of thumb and index finger
        cv2.circle(img,(x1,y1),13,(255,0,0),cv2.FILLED) #image #fingers #radius #rgb
        cv2.circle(img,(x2,y2),13,(255,0,0),cv2.FILLED) #image #fingers #radius #rgb
        cv2.line(img,(x1,y1),(x2,y2),(255,0,0),3)  #create a line b/w tips of index finger and thumb
 
        length = hypot(x2-x1,y2-y1) #distance b/w tips using hypotenuse
        # from numpy we find our length,by converting hand range in terms of volume range ie b/w -63.5 to 0
        vol = np.interp(length,[30,350],[volMin,volMax]) 
        volbar = np.interp(length,[30,350],[400,150])
        volper = np.interp(length,[30,350],[0,100])
        
        
        print(vol,int(length))
        volume.SetMasterVolumeLevel(vol, None)
        
        # Hand range 30 - 350
        # Volume range -63.5 - 0.0
        #creating volume bar for volume level 
        cv2.rectangle(img,(50,150),(85,400),(0,0,255),4) # vid ,initial position ,ending position ,rgb ,thickness
        cv2.rectangle(img,(50,int(volbar)),(85,400),(0,0,255),cv2.FILLED)
        cv2.putText(img,f"{int(volper)}%",(10,40),cv2.FONT_ITALIC,1,(0, 255, 98),3)
        #tell the volume percentage ,location,font of text,length,rgb color,thickness

    def fxn():
        # voice recognizer object
        voice = speech.Recognizer()
 
        # use microphone
        with speech.Microphone() as source:
            print("Say command: ")
            voice_command = voice.listen(source)
    
        # check input
        try:
            command = voice.recognize_google(voice_command)
            print(command)
            
        # handle the exceptions
        except speech.UnknownValueError:
            print("Google Speech Recognition system could not\
            understand your instructions please give instructions carefully")
            fxn()
            
        except speech.RequestError as e:
            print(
                "Could not request results from Google Speech Recognition \
                service; {0}".format(e))
            fxn()
            
        # validate input
        if command == "set volume 100" or command == "volume 100":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(5.9, None) #set volume headset
            # volume.SetMasterVolumeLevel(0.0, None)
            fxn()
                
        elif command == "set volume 90" or command == "volume 90":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(3.4, None) #set volume headset
            # volume.SetMasterVolumeLevel(-1.5, None) 
            fxn()

        elif command == "set volume 80" or command == "volume 80":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(1.0, None) #set volume headset
            # volume.SetMasterVolumeLevel(-3.4, None) 
            fxn()

        elif command == "set volume 70" or command == "volume 70":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(-1.0, None) #set volume headset
            # volume.SetMasterVolumeLevel(-5.4, None) 
            fxn()
            
        elif command == "set volume 60" or command == "volume 60":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(-3.4, None) #set volume headset
            # volume.SetMasterVolumeLevel(-7.6, None) 
            fxn()

        elif command == "set volume 50" or command == "volume 50":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(-6.2, None) #set volume headset
            # volume.SetMasterVolumeLevel(-10.2, None) 
            fxn()

        elif command == "set volume 40" or command == "volume 40":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(-9.4, None) #set volume headset
            # volume.SetMasterVolumeLevel(-13.6, None) 
            fxn()
            
        elif command == "set volume 30" or command == "volume 30":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevel(-9.-5, None) #set volume headset
            # volume.SetMasterVolumeLevel(-17.6, None) 
            fxn()

        elif command == "set volume mute" or command == "silent" or command == "mute" or command == "myut":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMute(1, None)
            fxn()
            
        elif command == "set volume unmute" or command == "set volume on mute" or command == "unmute" or command == "unmyut":
            device = AudioUtilities.GetSpeakers()
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMute(0, None)
            fxn()

        elif command == "stop" or command == "exit":
            sys.exit(0)
                
        else: 
            print("Not a command.")
            fxn()


    cv2.imshow('Image',img) #Show the video 
    if cv2.waitKey(1) & 0xff==ord(' '): #By using spacebar delay will stop
        break
 
# execute
fxn()
cap.release()#stop cam
cv2.destroyAllWindows() #close window

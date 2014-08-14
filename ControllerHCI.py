import sys
from cv2 import cv
import math
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")
resetCounter=0;   
action=None;
frame1 = {'pos':None}
frame2 = frame1.copy()
currentframe = frame1.copy()


def galleryDisplay(shell,delay):
    shell.Run("WLXPhotoGallery")
    shell.AppActivate("PhotoViewer")
    i=0
    while i < delay:
        i+=1;


haarFace = cv.Load('face.xml')
haarEyes = cv.Load('eye.xml')
haarleftEar = cv.Load('left_ear.xml')
haarrightEar = cv.Load('right_ear.xml')


def detect(imcolor):-,
    eyesList=[]
    faceList=[]
    leftEarList=[]
    rightEarList=[]

    
    image_size = cv.GetSize(imcolor)
    
    # create grayscale version
    grayscale = cv.CreateImage(image_size, 8, 1)
    cv.CvtColor(imcolor, grayscale, cv.CV_BGR2GRAY)

    # create storage
    storage = cv.CreateMemStorage(0)

    # equalize histogram
    cv.EqualizeHist(grayscale, grayscale)

    # detect objects
    detectedFace = cv.HaarDetectObjects(imcolor, haarFace, storage)
    detectedLeftEar = cv.HaarDetectObjects(imcolor, haarleftEar, storage)
    detectedRightEar = cv.HaarDetectObjects(imcolor, haarrightEar, storage)

    detectedEyes = cv.HaarDetectObjects(imcolor, haarEyes, storage)

    if detectedEyes:
        finalEyesList=[]
        for eye in detectedEyes:
            pt1=(eye[0][0],eye[0][1]);
            pt2=(eye[0][0]+eye[0][2],eye[0][1]+eye[0][3])
            eyesList.append((pt1,pt2));
        lengthElList = len(eyesList)
        if lengthElList >2:
            DiffMatrix  = [[0 for i in xrange(lengthElList)] for i in xrange(lengthElList)]
            for eye in eyesList:
                for eye2 in eyesList:
                    if eye != eye2:
                        x1eye1=eye[0][0]
                        x2eye1=eye[1][0]
                        y1eye1=eye[0][1]
                        y2eye1=eye[1][1]
                        x1eye2=eye2[0][0]
                        x2eye2=eye2[1][0]
                        y1eye2=eye2[0][1]
                        y2eye2=eye2[1][1]
                        Area1 = round(abs(x1eye1-x2eye1)*abs(y1eye1-y2eye1),3)
                        Area2 = round(abs(x1eye2-x2eye2)*abs(y1eye2-y2eye2),3)
                        diffy = round(math.pow((y1eye1-y1eye2),2)*math.pow((y2eye1-y2eye2),2),2)
                        diffx =round(min((x1eye1-x1eye2),(x1eye1-x2eye2),(x2eye1-x1eye2),(x2eye1-x2eye2)),2)
                        diffyfor =round(abs(min((y1eye1-y1eye2),(y1eye1-y2eye2),(y2eye1-y1eye2),(y2eye1-y2eye2))),2)
                        rule1= diffx < -1
                        rule2= max(Area1,Area2) > round(1.5*min(Area1,Area2),2)
                        rule3=(diffx> round(1*max(abs(x1eye1-x2eye1),abs(x1eye2-x2eye2)),2))
                        rule4 =(diffyfor> round(2*max(abs(y1eye1-y2eye1),abs(y1eye2-y2eye2)),2))
                        
                        if(rule1 or rule2 or rule3 or rule4 ):
                            Sum=0
                        else:
                            Sum =diffy+7*abs(Area1-Area2)
                            
             
                        DiffMatrix[eyesList.index(eye)][eyesList.index(eye2)]=Sum;
            try:
                valuex = min(x for x in DiffMatrix if max(x) > 0)
                indexX = DiffMatrix.index(valuex)
                valuey = min(x for x in valuex if x > 0)
                indexY = valuex.index(valuey)
                newList =[]
                newList.append(eyesList[indexX])
                newList.append(eyesList[indexY])
                finalEyesList=newList
            except:
                finalEyesList=None
            
        else:
            finalEyesList=None
    else:
        finalEyesList=None
        
    if finalEyesList:
        for eye in finalEyesList:
            cv.Rectangle(imcolor,eye[0],
            eye[1],cv.RGB(0, 255, 0),2)
            eyesList.append((pt1,pt2));
            
        eye1=finalEyesList[0];
        eye2=finalEyesList[1];
        x1eye1=eye1[0][0]
        x2eye1=eye1[1][0]
        y1eye1=eye1[0][1]
        y2eye1=eye1[1][1]
        x1eye2=eye2[0][0]
        x2eye2=eye2[1][0]
        y1eye2=eye2[0][1]
        y2eye2=eye2[1][1]
        Area1 = round(abs(x1eye1-x2eye1)*abs(y1eye1-y2eye1),3)
        Area2 = round(abs(x1eye2-x2eye2)*abs(y1eye2-y2eye2),3)
        Area = round((image_size[0]*image_size[1]),2)
        percentage = round((Area1+Area2)*100.0/(Area),2)
        
        if percentage > 10:
            currentframe['pos']="zoomin"
   
        elif( x1eye1<x1eye2 ): 
            if abs(y1eye1-y1eye2) < abs(y1eye1-y2eye1)/2:
                currentframe['pos']="straight"
            elif y1eye1<y1eye2:
                 currentframe['pos']="right"
            else:
                 currentframe['pos']="left"
            
        elif (x1eye1>x1eye2):
            if abs(y1eye2-y1eye1) < abs(y1eye2-y2eye2)/2:
                currentframe['pos']="straight"

            elif y1eye1>y1eye2:
                currentframe['pos']="right"

            else:
                currentframe['pos']="left"

            
        
    else:
        
        currentframe['pos']=None
        
        
        
        
    if detectedFace:
        face = detectedFace[0]
        pt1=(face[0][0],face[0][1]);
        pt2=(face[0][0]+face[0][2],face[0][1]+face[0][3])
        x1=pt1[0]
        y1=pt1[1]
        x2=pt2[0]
        y2=pt2[1]
        Area1 = round(abs(x1-x2)*abs(y1-y2),3)
        Area = image_size[0]*image_size[1]
        percentage = Area1*100.0/(Area)
        
        if percentage > 35:
            currentframe['pos']="zoomin"
        elif percentage < 10 and  currentframe['pos']!="zoomin":
            currentframe['pos']="zoomout"
        cv.Rectangle(imcolor,pt1,
        pt2,cv.RGB(255, 0, 0),2)
        faceList.append((pt1,pt2));

    if detectedLeftEar:
        for face in detectedLeftEar:
            pt1=(face[0][0],face[0][1]);
            pt2=(face[0][0]+face[0][2],face[0][1]+face[0][3])
            cv.Rectangle(imcolor,pt1,
            pt2,cv.RGB(255, 255, 255),2)
            leftEarList.append((pt1,pt2));
            currentframe['pos']="r2r"
            
    if detectedRightEar:
        for face in detectedRightEar:
            pt1=(face[0][0],face[0][1]);
            pt2=(face[0][0]+face[0][2],face[0][1]+face[0][3])
            cv.Rectangle(imcolor,pt1,
            pt2,cv.RGB(0, 0, 0),2)
            rightEarList.append((pt1,pt2));
            currentframe['pos']="r2l"

         
    

if __name__ == "__main__":
    print "Press ESC to exit ..."
    
    print "Welcome to our awesome Project"

    # create windows
    cv.NamedWindow('Raw', cv.CV_WINDOW_NORMAL)
    
    # create capture device
    device = 0 # assume we want first device
    capture = cv.CaptureFromCAM(0)
    cv.SetCaptureProperty(capture, cv.CV_CAP_PROP_FRAME_WIDTH, 640)
    cv.SetCaptureProperty(capture, cv.CV_CAP_PROP_FRAME_HEIGHT, 480)

    # check if capture device is OK
    if not capture:
        print "Error opening capture device"
        sys.exit(0)
    galleryDisplay(shell,10000000)
    shell.SendKeys("{ENTER}")
    
    while 1:
        # do forever

        # capture the current frame
        frame = cv.QueryFrame(capture)
        if frame is None:
            break

        # mirror
        cv.Flip(frame, None, 1)

        # face detection
        detect(frame)
        
       
        frame2 = currentframe.copy()
        if frame2['pos'] is not None :
            if action != frame2['pos']:
                action = frame2['pos'];
                print action
                if action == 'left':
                #goLeft
                    galleryDisplay(shell,1000000)
                    shell.SendKeys("{LEFT}")
                elif action == 'right':
                    #goRight
                    galleryDisplay(shell,1000000)
                    shell.SendKeys("{RIGHT}")
                elif action == 'zoomout':
                #ZoomOut
                    galleryDisplay(shell,1000000)
                    shell.SendKeys("-")
                elif action == 'zoomin':
                #ZoomIn
                    galleryDisplay(shell,1000000)
                    shell.SendKeys("=")
                elif action == 'r2r':
                #RotateRight
                    galleryDisplay(shell,1000000)
                    shell.SendKeys("^.")
                elif action == 'r2l':
                    #RotateLeft
                    galleryDisplay(shell,1000000)
                    shell.SendKeys("^,")
            else:
                if(resetCounter!=5):
                    resetCounter+=1;
                else:
                    resetCounter=0;
                    action=None
        
        
        # display webcam image
        cv.ShowImage('Raw', frame)
        
        # handle events
        cv.WaitKey(10)

        
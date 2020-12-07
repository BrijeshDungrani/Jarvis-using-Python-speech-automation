import cv2
import numpy as np

cap = cv2.VideoCapture("eye_recording.flv")

while True : 
    ret,frame = cap.read()

    cv2.imshow("Frame",frame)
    key = cv2.waitKey(30)

    if key == 27:
        break
    
cv2.destroyAllWindows()
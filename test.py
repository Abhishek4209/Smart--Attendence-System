from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch
def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('Data/haarcascade_frontalface_default.xml')

with open('Data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('Data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)


print('Shape of Faces matrix --> ', FACES.shape)


knn = KNeighborsClassifier(n_neighbors=5)


knn.fit(FACES, LABELS)
COL_NAMES = ['NAME', 'TIME']
while True:
    
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    # Iterate over detected faces
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]

        # Resize the cropped face image to 50x50 pixels and flatten it
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)

        # Get current timestamp
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")

        # Check if an attendance file for the current date already exists
        exist = os.path.isfile("Attendance_" + date + ".csv")

        # Draw rectangles and text on the frame for visualization
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 1)

        # Create an attendance record with predicted identity and timestamp
        attendance = [str(output[0]), str(timestamp)]

    # Display the current frame with annotations
    cv2.imshow("Frame", frame)


    k = cv2.waitKey(1)

    # If 'o' is pressed, announce attendance and save it to a CSV file
    if k == ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            # If file exists, append attendance to it
            with open("Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()
        else:
            # If file doesn't exist, create it and write column names and attendance
            with open("Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            csvfile.close()

    # If 'q' is pressed, exit the loop
    if k == ord('q'):
        break

video.release()
cv2.destroyAllWindows()

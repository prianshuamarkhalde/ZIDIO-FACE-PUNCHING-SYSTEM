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
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str1)

# Initialize the video capture and face detector
video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

# Load labels and faces data from pickled files
with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)

# Train the KNN classifier
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Load the background image
imgBackground = cv2.imread("background.png")

# Define the column names for the attendance CSV
COL_NAMES = ['NAME', 'TIME']

# Main loop for capturing video and processing faces
while True:
    ret, frame = video.read()  # Capture video frame by frame
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)  # Convert frame to grayscale
    faces = facedetect.detectMultiScale(gray, 1.3, 5)  # Detect faces in the frame

    # Process each face detected
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]  # Crop the detected face from the frame
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)  # Resize and flatten the face
        output = knn.predict(resized_img)  # Predict the name using KNN classifier

        # Get current timestamp for attendance logging
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")

        # Draw rectangles and labels around the detected face
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)

        # Store attendance details
        attendance = [str(output[0]), str(timestamp)]

    # Resize the frame to fit the red box (coordinates: 162:410, 55:695)
    frame_resized = cv2.resize(frame, (320, 248))

    # Insert the resized frame into the red box on the background image
    imgBackground[102:102 + 248, 55:55 + 320] = frame_resized

    # Display the frame
    cv2.imshow("Frame", imgBackground)

    # Wait for keypresses
    k = cv2.waitKey(1)
    if k == ord('o'):  # 'o' key to take attendance
        speak("Attendance Taken..")
        time.sleep(5)
        # Append or create a new CSV file for attendance
        if exist:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
        else:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)

    if k == ord('q'):  # 'q' key to quit
        break

# Release video capture and close windows
video.release()
cv2.destroyAllWindows()

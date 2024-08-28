
import cv2
import os
import pickle
import face_recognition
import pandas as pd

# Importing student images
folderPath = 'Images'
modePathList = os.listdir(folderPath)
print(modePathList)
imgList = []

# Valid image extensions
valid_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff')
studentIds = []
for path in modePathList:
    if path.lower().endswith(valid_extensions):
        try:
            img = cv2.imread(os.path.join(folderPath, path))
            studentIds.append(os.path.splitext(path)[0])
            print(studentIds)
            if img is not None:
                imgList.append(img)
            else:
                print(f"Warning: {path} could not be read as an image.")
        except Exception as e:
            print(f"Error reading {path}: {e}")
    else:
        print(f"Skipping non-image file: {path}")

print(f"Total number of images loaded: {len(imgList)}")

def findEncodings(imagesList):
    encodeList = []
    for img in imagesList:
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        encode = face_recognition.face_encodings(img)[0]
        encodeList.append(encode)
    return encodeList

print("Encoding started")
encodeListKnown = findEncodings(imgList)
encodeListKnownWithIds = [encodeListKnown, studentIds]
print("Encoding complete")

file = open("EncodeFile.p", 'wb')
pickle.dump(encodeListKnownWithIds, file)
file.close()
print("File Saved")
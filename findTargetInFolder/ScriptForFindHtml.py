import os
import socket

dir = "D:/Programming/UnityDocumentation/Manual"
keyWord = input("Finding:")
for root, dirs,files in os.walk(dir):
    for fi in files:
        if keyWord in fi:
            file_path = os.path.join(dir,fi)
            print(file_path)

            os.startfile(file_path)




# Face Recognition Attendance System

This is a simple face recognition attendance system implemented in Python using OpenCV and face_recognition library. It captures faces through the webcam, recognizes them, and maintains attendance records in an Excel file.

## Features

- **Real-time Face Recognition**: Detect and recognize faces in real-time using a webcam.
- **Attendance Tracking**: Keep track of daily attendance in an Excel file with a new sheet for each day.
- **Dynamic Excel Sheets**: Sheets are named according to the current date.
- **Easy Setup**: Load known faces from a specified picture folder.

## Prerequisites

Make sure to install the required libraries before running the code:

```bash
pip install opencv-python
pip install face_recognition
pip install openpyxl

import cv2
import numpy as np
from PIL import ImageGrab
import os
import base64
import requests
import face_recognition
import re
import time
import pandas as pd
from datetime import datetime
import logging
import threading
from tkinter import Tk, Canvas, Toplevel

# Configuration parameters
API_KEY = ""
BBOX = (100, 100, 800, 600)
CSV_FILE = 'output.csv'
UPDATE_INTERVAL = 2  # seconds

# Logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to encode the image from frame
def encode_image_from_frame(frame):
    _, buffer = cv2.imencode('.jpeg', frame)
    return base64.b64encode(buffer).decode('utf-8')

# Function to save image with detected attributes
def save_image_with_attributes(image, attributes, face_location, frame_count, save_dir="output"):
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    new_name = f"frame{frame_count}.jpeg"
    new_path = os.path.join(save_dir, new_name)

    # Draw a box around the face
    top, right, bottom, left = face_location
    cv2.rectangle(image, (left, top), (right, bottom), (0, 255, 0), 2)

    # Put the attribute text on the box
    for i, (key, value) in enumerate(attributes.items()):
        cv2.putText(image, f"{key}: {value}", (left, top - 10 - (i * 20)), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)

    cv2.imwrite(new_path, image)
    logging.info(f"Saved image with attributes: {new_path}")

# Function to analyze attributes using OpenAI API
def analyze_attributes(encoded_image):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }

    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            "Please analyze the face in this image and provide the following information:\n"
                            "- Emotion in one word\n"
                            "- Gender\n"
                            "- Age as a single number, don't put things like unknown, just one number\n"
                            "- Smile (yes or no)\n"
                            "- Beard (yes or no)\n"
                            "- Glasses (yes or no)\n"
                            "- Makeup (yes or no)\n"
                            "- Hair color\n"
                            "- Moustache (yes or no)\n"
                            "- Baldness (yes or no)"
                        )
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{encoded_image}"
                        }
                    }
                ]
            }
        ],
        "max_tokens": 300
    }

    try:
        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
        response_data = response.json()
        logging.debug(response_data)

        if "choices" in response_data and response_data["choices"]:
            content = response_data["choices"][0]["message"]["content"]
            logging.debug(f"API response content: {content}")
            # Extract the attributes using regular expressions with checks
            attributes = {
                'emotion': extract_attribute(r'Emotion: (\w+)', content),
                'gender': extract_attribute(r'Gender: (\w+)', content),
                'age': extract_attribute(r'Age: (\d+)', content),
                'smile': extract_attribute(r'Smile: (\w+)', content),
                'beard': extract_attribute(r'Beard: (\w+)', content),
                'glasses': extract_attribute(r'Glasses: (\w+)', content),
                'makeup': extract_attribute(r'Makeup: (\w+)', content),
                'hair_color': extract_attribute(r'Hair color: (\w+)', content),
                'moustache': extract_attribute(r'Moustache: (\w+)', content),
                'baldness': extract_attribute(r'Baldness: (\w+)', content)
            }
            logging.info(f"Detected attributes: {attributes}")
            return attributes
    except Exception as e:
        logging.error(f"Error analyzing attributes: {e}")

    return None

def extract_attribute(pattern, content):
    match = re.search(pattern, content, re.IGNORECASE)
    return match.group(1) if match else 'Unknown'

# Function to process each detected face
def process_face(frame, face_location, face_id, frame_count):
    encoded_image = encode_image_from_frame(frame)
    attributes = analyze_attributes(encoded_image)

    if attributes:
        save_image_with_attributes(frame, attributes, face_location, frame_count)
        return attributes

    return None

class OverlayWindow:
    def __init__(self):
        self.root = Toplevel()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-transparent", True)
        self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")
        self.canvas = Canvas(self.root, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.root.attributes("-alpha", 0.3)

    def show_rectangle(self, left, top, right, bottom):
        self.canvas.create_rectangle(left, top, right, bottom, outline="green", width=2)
        self.root.after(1000, self.clear_rectangles)  # Remove rectangles after 1 second

    def clear_rectangles(self):
        self.canvas.delete("all")

    def run(self):
        self.root.mainloop()

overlay_window = OverlayWindow()

def detect_faces():
    frame_count = 0
    face_id = 0

    while True:
        # Capture the screen
        screen = np.array(ImageGrab.grab(bbox=BBOX))
        frame = cv2.cvtColor(screen, cv2.COLOR_RGB2BGR)
        rgb_frame = frame[:, :, ::-1]

        # Detect faces in the frame
        face_locations = face_recognition.face_locations(rgb_frame)

        if face_locations:
            logging.info("Faces detected, drawing rectangles and proceeding with OpenAI analysis...")
            for face_location in face_locations:
                top, right, bottom, left = face_location
                left_screen = BBOX[0] + left
                top_screen = BBOX[1] + top
                right_screen = BBOX[0] + right
                bottom_screen = BBOX[1] + bottom
                overlay_window.show_rectangle(left_screen, top_screen, right_screen, bottom_screen)

                # Process each face and get attributes
                attributes = process_face(frame, face_location, face_id, frame_count)
                if attributes:
                    now = datetime.now()
                    data = {
                        'employee_id': '',  # Fill in the employee_id if available
                        'date': now.strftime("%Y-%m-%d"),
                        'time': now.strftime("%H:%M:%S"),
                        'faceId': face_id,
                        'top': face_location[0],
                        'left': face_location[3],
                        'width': face_location[1] - face_location[3],
                        'height': face_location[2] - face_location[0],
                        'smile': 1 if attributes['smile'].lower() == 'yes' else 0,
                        'gender': attributes['gender'],
                        'age': attributes['age'],
                        'beard': attributes['beard'],
                        'glasses': attributes['glasses'],
                        'makeup': attributes['makeup'],
                        'hair_color': attributes['hair_color'],
                        'emotion': attributes['emotion'],
                        'moustache': attributes['moustache'],
                        'baldness': attributes['baldness']
                    }

                    # Append data to the CSV file
                    pd.DataFrame([data]).to_csv(CSV_FILE, mode='a', header=not os.path.exists(CSV_FILE), index=False)
                    face_id += 1
                    frame_count += 1  # Increment frame count here

        # Add a delay to avoid overwhelming the API
        time.sleep(UPDATE_INTERVAL)

# Start the face detection in a separate thread
threading.Thread(target=detect_faces, daemon=True).start()

# Run the overlay window
overlay_window.run()

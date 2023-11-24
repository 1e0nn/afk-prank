import time
import pygame
from pynput import mouse, keyboard as pynput_keyboard
import os
import cv2
import subprocess
import sys

# Constants
SOUND_FILE = "C:\\the\\path\\Downloads\\sound.mp3"
PHOTO_FOLDER = "C:\\the\\path\\prank_photo_folder"
CAPTURE_COUNT_THRESHOLD = 4
quit_program_key = pynput_keyboard.Key.page_up #change 'page_up' to whatever key you want to use to quit the program

# Initialize Pygame for playing sounds
pygame.init()

# Load the sound file
pygame.mixer.music.load(SOUND_FILE)

# Create a folder to store photos
os.makedirs(PHOTO_FOLDER, exist_ok=True)

# Access the webcam (0 for the default webcam)
capture = cv2.VideoCapture(0)

# Check if the webcam initialization is successful
if not capture.isOpened():
    print("Error: Unable to access the webcam. Photo capture will be disabled.")
    webcam_available = False
else:
    webcam_available = True
    # Read an image from the webcam (necessary to get correct properties)
    _, _ = capture.read()

# Counter for captures
captures_count = 0

# Flag variable to indicate if the program should exit
exit_program = False

# Callback function when a key is pressed
def on_key_press(key):
    try:
        if key == quit_program_key:
            quit_program()
        else:
            perform_actions()
    except AttributeError:
        pass

# Function to play the sound and capture a photo with the webcam
def play_sound_and_capture_photo():
    global captures_count

    # Read an image from the webcam
    _, image = capture.read()

    # Save the image in the photos folder
    photo_path = os.path.join(PHOTO_FOLDER, f"captured_photo_{captures_count}.jpg")
    cv2.imwrite(photo_path, image, [cv2.IMWRITE_JPEG_QUALITY, 80])
    print(f"Photo captured: {photo_path}")

# Callback function when a mouse button is clicked
def on_mouse_event(x, y, button, pressed):
    if pressed and button in [mouse.Button.left, mouse.Button.right]:
        perform_actions()

# Additional function for actions
def perform_actions():
    global captures_count
    try:
        if webcam_available is False:
            pygame.mixer.music.play()
            captures_count += 1
            if captures_count >= CAPTURE_COUNT_THRESHOLD:
                lock_screen()
        else:
            pygame.mixer.music.play()
            play_sound_and_capture_photo()
            captures_count += 1
            if captures_count >= CAPTURE_COUNT_THRESHOLD:
                lock_screen()
    except AttributeError:
        pass

# Function to lock the screen
def lock_screen():
    global exit_program
    # Use rundll32.exe command to lock the screen on Windows
    subprocess.run(["rundll32.exe", "user32.dll,LockWorkStation"])
    # Set the flag to indicate that the program should exit
    exit_program = True

# Function to quit the program gracefully
def quit_program():
    global exit_program
    exit_program = True

# Register the callback functions
keyboard_listener = pynput_keyboard.Listener(on_press=on_key_press)
mouse_listener = mouse.Listener(on_click=on_mouse_event)
keyboard_listener.start()
mouse_listener.start()

# Loop to keep the program running
while not exit_program:
    time.sleep(1)  # Add a small pause to reduce CPU usage

# Clean up the program
keyboard_listener.stop()
mouse_listener.stop()
keyboard_listener.join()
mouse_listener.join()
pygame.mixer.quit()
capture.release()  # Release the webcam resource
sys.exit()

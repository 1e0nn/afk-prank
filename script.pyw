import time
import pygame
from pynput import mouse, keyboard as pynput_keyboard
import os
import cv2
import subprocess
import sys
import getpass
import requests
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL, cast, CoInitialize
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Constants
PHOTO_FOLDER = f"C:\\Users\\{getpass.getuser()}\\Documents\\Captured_Photos"
CAPTURE_COUNT_THRESHOLD = 3
quit_program_key = pynput_keyboard.Key.shift_r  # change 'shift_r' to whatever key you want to use to quit the program
sound_volume = 1.0  # 0.0 means minimum volume and 1.0 means maximum volume
# URL of the Canary token : Web bug / URL token
url_canary_token = "http://canarytokens.com/traffic/...../contact.php"  
    

CoInitialize()

def set_system_volume(sound_volume=sound_volume):
    try:
        # Get the default audio device manager
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)

        # Convert the interface into an IAudioEndpointVolume object
        volume = cast(interface, POINTER(IAudioEndpointVolume))

        # Set the volume to maximum
        volume.SetMasterVolumeLevelScalar(sound_volume, None)
        print("System volume set to", round(sound_volume * 100), "%.")
        return True
    except Exception as e:
        print("Error while setting volume:", str(e))
        return False
    finally:
        if 'interface' in locals():
            interface.Release()

def download_mp3(url):
    # Full path of the downloaded file
    filepath = f"C:\\Users\\{getpass.getuser()}\\Music\\sound_alert.mp3"
    try:
        # Download the file
        response = requests.get(url,verify=False)
        if response.status_code == 200:
            with open(filepath, 'wb') as f:
                f.write(response.content)
            return filepath
        else:
            print("Error while downloading: Code", response.status_code)
            return None
    except Exception as e:
        print("An error occurred while downloading:", str(e))
        return None

# Function to trigger a Canary token
def trigger_canary_token(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            print("Canary token triggered successfully!")
        else:
            print("Failed to trigger Canary token. Status code:", response.status_code)
    except requests.RequestException as e:
        print("Error:", e)

def find_windows_ding_sound():
    # Default path for the "Windows Ding" sound
    default_sound_path = "C:\\Windows\\Media\\Windows Critical Stop.wav"
    if os.path.exists(default_sound_path):
        return default_sound_path
    else:
        print("The 'Windows Ding' sound file was not found.")
        return None

# Test the download_mp3 function
mp3_path = download_mp3("http://cdn.pixabay.com/download/audio/2021/08/09/audio_1818162f91.mp3?filename=cartoon-scream-1-6835.mp3")
ding_sound_path = find_windows_ding_sound()

if mp3_path:
    SOUND_FILE = mp3_path
    set_system_volume()
    print("Downloaded sound will be used:", mp3_path)
elif ding_sound_path:
    SOUND_FILE = ding_sound_path
    set_system_volume()
    print("The 'Windows Ding' sound will be used:", ding_sound_path)
else:
    print("Unable to find a sound to play.")


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
    photo_path = os.path.join(PHOTO_FOLDER, f"captured_photo_{time.strftime('%Y_%m_%d_%H_%M_%S')}.jpg")
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
            trigger_canary_token(url_canary_token)
            captures_count += 1
            if captures_count >= CAPTURE_COUNT_THRESHOLD:
                lock_screen()
        else:
            pygame.mixer.music.play()
            trigger_canary_token(url_canary_token)
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


# Remove the downloaded sound
if mp3_path:
    os.remove(mp3_path)
capture.release()  # Release the webcam resource
sys.exit()

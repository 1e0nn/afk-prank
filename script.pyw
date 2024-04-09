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
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


# Constants
PHOTO_FOLDER = f"C:\\Users\\{getpass.getuser()}\\Documents\\Captured_Photos"
CAPTURE_COUNT_THRESHOLD = 3
quit_program_key = pynput_keyboard.Key.shift_r #change 'page_up' to whatever key you want to use to quit the program
sound_volume=0.70 # 0.0 signifie le volume minimum et 1.0 signifie le volume maximum

def set_system_volume_to_max(sound_volume=sound_volume):
    try:
        # Obtenir le gestionnaire de périphériques audio par défaut
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)

        # Convertir l'interface en objet IAudioEndpointVolume
        volume = cast(interface, POINTER(IAudioEndpointVolume))

        # Régler le volume au maximum
        volume.SetMasterVolumeLevelScalar(sound_volume, None)  
        print("Volume système réglé à", round(sound_volume * 100), "%.")
        return True
    except Exception as e:
        print("Erreur lors du réglage du volume :", str(e))
        return False

def download_mp3(url):

    # Chemin complet du fichier téléchargé
    filepath = f"C:\\Users\\{getpass.getuser()}\\Downloads\\sound_alert.mp3"

    try:
        # Téléchargement du fichier
        response = requests.get(url)
        if response.status_code == 200:
            with open(filepath, 'wb') as f:
                f.write(response.content)
            return filepath
        else:
            #print("Erreur lors du téléchargement : Code", response.status_code)
            return None
    except Exception as e:
        #print("Une erreur est survenue lors du téléchargement :", str(e))
        return None

def find_windows_ding_sound():
    # Chemin par défaut pour le son "Windows Ding"
    default_sound_path = "C:\\Windows\\Media\\Windows Critical Stop.wav"
    if os.path.exists(default_sound_path):
        return default_sound_path
    else:
        #print("Le son 'Windows Ding' n'a pas été trouvé.")
        return None

# Test de la fonction download_mp3
mp3_path = download_mp3("https://cdn.pixabay.com/download/audio/2021/08/09/audio_1818162f91.mp3?filename=cartoon-scream-1-6835.mp3")
ding_sound_path = find_windows_ding_sound()


if mp3_path:
    SOUND_FILE = mp3_path
    set_system_volume_to_max()
    print("Le son téléchargé sera utilisé :", mp3_path)
elif ding_sound_path:
    SOUND_FILE = ding_sound_path
    set_system_volume_to_max()
    print("Le son 'Windows Ding' sera utilisé :", ding_sound_path)
else:
    print("Impossible de trouver un son à jouer.")

sys.exit()
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
#supprimer le son téléchargé
if mp3_path:
    os.remove(mp3_path)
capture.release()  # Release the webcam resource
sys.exit()

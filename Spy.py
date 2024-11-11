import threading
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from winotify import Notification
from datetime import datetime

def notify(title, message):
    toast = Notification(
        app_id="Microphone Monitor",
        title=title,
        msg=message
    )
    toast.show()

def check_audio():
    global running
    last_status = None
    consecutive_samples = 0  # Counter for consecutive readings
    THRESHOLD = 0.01  # Mic activation threshold
    REQUIRED_SAMPLES = 3  # Number of consecutive samples needed to trigger status change
    
    # Get default microphone
    devices = AudioUtilities.GetMicrophone()
    interface = devices.Activate(
        IAudioMeterInformation._iid_, CLSCTX_ALL, None)
    meter = cast(interface, POINTER(IAudioMeterInformation))
    
    while running:
        try:
            # Get peak value
            peak = meter.GetPeakValue()
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Check if microphone is active with debouncing
            if peak > THRESHOLD:
                consecutive_samples += 1
                if consecutive_samples >= REQUIRED_SAMPLES and last_status != "in_use":
                    print(f"[{current_time}] Microphone is actively recording")
                    notify("Audio Alert", f"Microphone is actively recording\nStarted at: {current_time}")
                    last_status = "in_use"
            else:
                consecutive_samples -= 1
                if consecutive_samples <= 0:
                    consecutive_samples = 0
                    if last_status != "available" and last_status is not None:
                        print(f"[{current_time}] Microphone is not in use")
                        notify("Audio Alert", f"Microphone stopped recording\nStopped at: {current_time}")
                        last_status = "available"
                    elif last_status is None:
                        last_status = "available"
                        print(f"[{current_time}] Microphone is not in use")
                    
        except Exception as e:
            print(f"Error checking microphone: {e}")
            
        time.sleep(0.1)  # Reduced sleep time for more accurate sampling

if __name__ == "__main__":
    running = True
    print("Monitoring microphone. Press Ctrl+C to stop...")
    
    try:
        check_audio()
    except KeyboardInterrupt:
        print("\nStopping monitoring...")
        running = False

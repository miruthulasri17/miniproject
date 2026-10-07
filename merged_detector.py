import cv2
import numpy as np
import threading
import time
import os
import re
import queue
from ultralytics import YOLO

# Import the FaceDetector class from the standalone script
from standalone_face_detector import FaceDetector

# Change current working directory to the script's directory
os.chdir(os.path.dirname(os.path.abspath(__file__)))

print("Starting merged detector script... please wait.")

# --- Configuration ---
ENABLE_SPEECH = True

# --- Thread-Safe Audio System ---
speech_queue = queue.Queue()

def speech_worker():
    """Worker thread that processes speech requests one by one using win32com (most reliable on Windows)"""
    print("Speech worker starting...")

    # Initialize COM for this thread (required on Windows)
    try:
        import pythoncom
        pythoncom.CoInitialize()
        print("COM library initialized.")
    except Exception as e:
        print(f"COM init note: {e}")

    # Try win32com first (most reliable on Windows)
    speaker = None
    use_win32 = False
    try:
        import win32com.client
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        speaker.Rate = -1   # Slightly slower for clarity
        speaker.Volume = 100
        use_win32 = True
        print("Speech engine: Windows SAPI (win32com) - READY")
    except Exception as e:
        print(f"win32com not available ({e}), falling back to pyttsx3...")

    # Fallback: pyttsx3
    if not use_win32:
        try:
            import pyttsx3
            speaker = pyttsx3.init('sapi5')
            speaker.setProperty('rate', 150)
            speaker.setProperty('volume', 1.0)
            print("Speech engine: pyttsx3 - READY")
        except Exception as e:
            print(f"pyttsx3 also failed: {e}")
            speaker = None

    if speaker is None:
        print("ERROR: No speech engine available. Speech disabled.")
        return

    while True:
        try:
            text = speech_queue.get()
            if text is None:
                break

            print(f"[SPEECH] Speaking: '{text}'")
            try:
                if use_win32:
                    speaker.Speak(text)
                else:
                    speaker.say(text)
                    speaker.runAndWait()
            except Exception as e:
                print(f"[SPEECH] Speak error: {e}")
                # Try to re-init win32com on failure
                try:
                    import win32com.client
                    speaker = win32com.client.Dispatch("SAPI.SpVoice")
                    speaker.Speak(text)
                except Exception as e2:
                    print(f"[SPEECH] Re-init also failed: {e2}")

            speech_queue.task_done()
            time.sleep(0.2)
        except Exception as e:
            print(f"[SPEECH] Loop error: {e}")
            time.sleep(1.0)

# Start the speech worker thread
if ENABLE_SPEECH:
    t = threading.Thread(target=speech_worker, daemon=True)
    t.start()
    print("Speech worker thread started.")

def speak_async(text):
    """Adds a message to the speech queue"""
    if not ENABLE_SPEECH:
        return
    # Clear queue if backed up (avoid stale announcements)
    while speech_queue.qsize() > 1:
        try:
            speech_queue.get_nowait()
            speech_queue.task_done()
        except:
            break
    print(f"[QUEUE] Adding: '{text}'")
    speech_queue.put(text)

# --- Currency label to speech mapping ---
WORD_MAP = {
    '10':   'it is ten rupees',
    '20':   'it is twenty rupees',
    '50':   'it is fifty rupees',
    '100':  'it is one hundred rupees',
    '200':  'it is two hundred rupees',
    '500':  'it is five hundred rupees',
    '2000': 'it is two thousand rupees',
}

def label_to_speech(label):
    """Convert a model label like '100Rs' to a spoken phrase"""
    raw_val = re.sub(r'[^0-9]', '', label)
    return WORD_MAP.get(raw_val, f'it is {raw_val} rupees')

# --- Detection Logic ---
def run_merged_detector():
    model_path = 'best.pt'

    if not os.path.exists(model_path):
        print(f"Error: {model_path} not found! Please ensure best.pt is in the same folder.")
        return

    print(f"Loading Currency model from {model_path}...")
    currency_model = YOLO(model_path)
    print("Currency Model loaded.")

    print("Initializing Face Detector...")
    face_detector = FaceDetector()

    # Give speech worker time to start before first announcement
    time.sleep(1.5)
    speak_async("System ready. Currency and face detection active.")

    # Robust Camera setup
    print("Initializing camera... please wait.")
    cap = None
    for index in [0, 1, 2]:
        for backend in [cv2.CAP_DSHOW, cv2.CAP_MSMF, cv2.CAP_ANY]:
            print(f"  Trying camera {index}, backend {backend}...")
            temp_cap = cv2.VideoCapture(index, backend)
            if temp_cap.isOpened():
                ret, frame = temp_cap.read()
                if ret and frame is not None and np.sum(frame) > 0:
                    cap = temp_cap
                    print(f"Camera {index} is working.")
                    break
                else:
                    temp_cap.release()
        if cap:
            break

    if not cap:
        print("CRITICAL: No working camera found.")
        speak_async("Camera error. Please check your webcam.")
        return

    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

    last_spoken_time = 0
    last_label = None
    no_detection_start_time = time.time()
    last_no_detection_speak_time = 0

    # Stability: require same detection for N consecutive frames before speaking
    stability_buffer = []
    STABILITY_FRAMES = 3

    VOICE_COOLDOWN = 4.0          # seconds between repeating same currency
    NO_DETECTION_TIMEOUT = 2.0   # seconds before saying "no currency"
    CONFIDENCE_THRESHOLD = 0.20

    print(f"Model Labels: {currency_model.names}")
    print("Merged Detection started. Press 'q' to quit.")

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        # 1. PROCESS FACE DETECTION
        frame = face_detector.process_frame(frame)

        # 2. PROCESS CURRENCY DETECTION
        results = currency_model(frame, stream=True, conf=CONFIDENCE_THRESHOLD, verbose=False)

        detections_found = False
        raw_best_label = None
        highest_conf = 0

        for r in results:
            for box in r.boxes:
                detections_found = True
                conf = float(box.conf[0])
                cls = int(box.cls[0])
                label_raw = currency_model.names[cls]

                if conf > highest_conf:
                    highest_conf = conf
                    raw_best_label = label_raw

                # Draw bounding box
                x1, y1, x2, y2 = map(int, box.xyxy[0])
                cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 255, 0), 3)
                cv2.putText(frame, f"{label_raw} ({conf*100:.0f}%)",
                            (x1, y1 - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)

        # Stability buffer
        stability_buffer.append(raw_best_label if detections_found else None)
        if len(stability_buffer) > STABILITY_FRAMES:
            stability_buffer.pop(0)

        stable_label = None
        if (len(stability_buffer) == STABILITY_FRAMES and
                all(x == stability_buffer[0] for x in stability_buffer) and
                stability_buffer[0] is not None):
            stable_label = stability_buffer[0]

        # Audio feedback
        current_time = time.time()

        if stable_label:
            no_detection_start_time = current_time
            speech_text = label_to_speech(stable_label)
            is_new_note = (speech_text != last_label)
            time_since_last = current_time - last_spoken_time

            if is_new_note or time_since_last > VOICE_COOLDOWN:
                speak_async(speech_text)
                last_spoken_time = current_time
                last_label = speech_text

        elif not detections_found:
            if (current_time - no_detection_start_time) > NO_DETECTION_TIMEOUT:
                if last_label is not None:
                    speak_async("No currency detected")
                    last_label = None

                if (current_time - last_no_detection_speak_time) > 30.0:
                    speak_async("Scanning")
                    last_no_detection_speak_time = current_time

        # UI status
        if stable_label:
            status_text = label_to_speech(stable_label).upper()
            color = (0, 255, 0)
        elif detections_found:
            status_text = "Scanning Currency..."
            color = (0, 165, 255)
        else:
            status_text = "No Currency"
            color = (0, 0, 255)

        cv2.putText(frame, status_text, (10, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.9, color, 2)
        
        # Display the merged frame
        cv2.imshow("Merged Currency & Face Detector", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    run_merged_detector()

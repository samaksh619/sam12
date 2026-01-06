import cv2
import mediapipe as mp
import numpy as np
from math import hypot

from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from comtypes.client import CreateObject
from comtypes import CoInitialize, CoUninitialize

# Prefer the canonical pycaw entrypoints; provide a clear error if import fails
try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except Exception:
    try:
        # Fallback for some pycaw installs/layouts
        from pycaw.utils import AudioUtilities  # type: ignore
        from pycaw.api.endpointvolume import IAudioEndpointVolume  # type: ignore
    except Exception:
        raise ImportError(
            "pycaw import failed. Install pycaw and comtypes: `pip install pycaw comtypes`"
        )

# ---------------- Camera ----------------
cap = cv2.VideoCapture(0)

# ---------------- MediaPipe ----------------
mpHands = mp.solutions.hands
hands = mpHands.Hands(
    max_num_hands=1,
    min_detection_confidence=0.7,
    min_tracking_confidence=0.7
)
mpDraw = mp.solutions.drawing_utils

# ---------------- Audio (CORRECT WAY) ----------------
device = AudioUtilities.GetSpeakers()

# Some pycaw versions/wrappers return an object that exposes a direct
# COM `Activate` method, while others wrap the underlying COM object.
# Try several common access patterns so this works across installs.
interface = None

# Direct call if available
if hasattr(device, "Activate"):
    try:
        interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    except Exception:
        interface = None

# Common wrapper attribute names that may hold the underlying COM object
if interface is None:
    for attr in ("_comobj", "_com_object", "_device", "device", "_mmdevice", "_ptr"):
        obj = getattr(device, attr, None)
        if obj and hasattr(obj, "Activate"):
            try:
                interface = obj.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                break
            except Exception:
                interface = None

if interface is None:
    # Print lightweight debug info to help diagnose wrapper variations
    # Some pycaw wrappers expose the endpoint directly as `EndpointVolume`.
    ev = getattr(device, "EndpointVolume", None)
    if ev is not None:
        # Use the provided EndpointVolume object directly
        volume = ev
        try:
            volMin, volMax = volume.GetVolumeRange()[:2]
        except Exception:
            # If the EndpointVolume wrapper is a null/invalid COM pointer,
            # fall back to creating an MMDeviceEnumerator and activating
            # the default endpoint directly via COM.
            try:
                # Ensure COM is initialized on this thread before creating COM objects
                try:
                    CoInitialize()
                    initialized = True
                except Exception:
                    initialized = False

                enum = CreateObject("MMDeviceEnumerator.MMDeviceEnumerator")
                dev = enum.GetDefaultAudioEndpoint(0, 1)  # eRender(0), eMultimedia(1)
                interface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                if interface:
                    volume = cast(interface, POINTER(IAudioEndpointVolume))
                    volMin, volMax = volume.GetVolumeRange()[:2]
                else:
                    raise RuntimeError("Activate returned NULL COM pointer")
            except Exception:
                try:
                    if initialized:
                        CoUninitialize()
                except Exception:
                    pass
                # Last resort: try casting the existing object
                try:
                    volume = cast(volume, POINTER(IAudioEndpointVolume))
                    volMin, volMax = volume.GetVolumeRange()[:2]
                except Exception:
                    # Give helpful debug info and raise
                    try:
                        print("DEBUG: EndpointVolume repr:", repr(ev))
                        print("DEBUG: EndpointVolume type:", type(ev))
                    except Exception:
                        pass
                    raise
    else:
        # Print lightweight debug info to help diagnose wrapper variations
        try:
            print("DEBUG: device type:", type(device))
            public_attrs = [a for a in dir(device) if not a.startswith("_")]
            print("DEBUG: device public attributes:", public_attrs)
        except Exception as e:
            print("DEBUG: failed to inspect device:", e)

        raise AttributeError(
            "Could not activate audio endpoint: the pycaw AudioDevice object has no 'Activate' method.\n"
            "I printed debug info above — please paste that output here.\n"
            "Confirm pycaw/comtypes are installed and compatible: `pip install --upgrade pycaw comtypes`"
        )

if interface is not None:
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volMin, volMax = volume.GetVolumeRange()[:2]
else:
    # If we fell back to using an `EndpointVolume` attribute earlier,
    # `volume` and `volMin`/`volMax` should already be set. Verify that.
    if 'volume' not in locals():
        raise RuntimeError(
            "Failed to obtain audio endpoint volume interface (no interface and no EndpointVolume)."
        )
    # Ensure volMin/volMax are available
    try:
        volMin, volMax
    except NameError:
        try:
            volMin, volMax = volume.GetVolumeRange()[:2]
        except Exception as e:
            raise RuntimeError("Could not read volume range from EndpointVolume") from e

# ---------------- Main Loop ----------------
while True:
    success, img = cap.read()
    if not success:
        break

    imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = hands.process(imgRGB)

    lmList = []

    if results.multi_hand_landmarks:
        for handLms in results.multi_hand_landmarks:
            for id, lm in enumerate(handLms.landmark):
                h, w, _ = img.shape
                cx, cy = int(lm.x * w), int(lm.y * h)
                lmList.append([id, cx, cy])

            mpDraw.draw_landmarks(img, handLms, mpHands.HAND_CONNECTIONS)

    if lmList:
        x1, y1 = lmList[4][1], lmList[4][2]   # Thumb tip
        x2, y2 = lmList[8][1], lmList[8][2]   # Index tip

        cv2.circle(img, (x1, y1), 8, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 8, (255, 0, 255), cv2.FILLED)
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)

        length = hypot(x2 - x1, y2 - y1)

        vol = np.interp(length, [15, 220], [volMin, volMax])
        volume.SetMasterVolumeLevel(vol, None)

    cv2.imshow("Hand Volume Control", img)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
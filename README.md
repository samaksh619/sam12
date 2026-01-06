# Hand Volume Control 🎚️

A real-time hand gesture recognition application that allows you to control your system volume using hand gestures captured via webcam.

## Features

- **Hand Detection**: Real-time hand tracking using MediaPipe
- **Gesture Recognition**: Distance-based volume control (thumb-to-index finger distance)
- **Visual Feedback**: Live camera feed with hand landmarks and distance line displayed
- **Cross-Platform**: Windows-compatible audio control via pycaw

## Requirements

- Python 3.7+
- Webcam
- Windows OS (for audio control via pycaw)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/samaksh619/sam12.git
   cd sam12
   ```

2. Install dependencies:
   ```bash
   pip install opencv-python mediapipe numpy pycaw comtypes
   ```

## Usage

Run the application:
```bash
python index.py
```

### Controls

- **Volume Control**: 
  - Move your thumb and index finger closer to decrease volume
  - Move them farther apart to increase volume
  - Distance range: 15–220 pixels maps to min–max volume

- **Exit**: Press `q` to quit the application

## How It Works

1. **Hand Detection**: MediaPipe Hands detects hand landmarks (21 keypoints per hand)
2. **Distance Calculation**: Calculates Euclidean distance between thumb tip (landmark 4) and index tip (landmark 8)
3. **Volume Mapping**: Maps the distance to system volume using linear interpolation
4. **Audio Control**: pycaw interfaces with Windows audio endpoints to set master volume level

## Technical Details

- **Framework**: MediaPipe (hand pose estimation)
- **Computer Vision**: OpenCV (camera feed & visualization)
- **Audio API**: pycaw (Windows audio endpoint control via COM)
- **Math**: NumPy (distance calculation & interpolation)

## Troubleshooting

### "NULL COM pointer access" error
- Ensure `pycaw` and `comtypes` are installed: `pip install --upgrade pycaw comtypes`
- This can occur if Windows audio endpoints are misconfigured

### Camera not detected
- Verify your webcam is connected and not in use by another application
- Check OpenCV installation: `python -c "import cv2; print(cv2.__version__)"`

### Hand detection not working
- Ensure good lighting and a clear background
- Keep your hand in view of the webcam at a distance of 30cm–80cm

## Architecture

- **Hand Detection Model**: MediaPipe Hands (TensorFlow Lite XNNPACK delegate)
- **Hand Tracking**: 1 hand max, 70% detection confidence, 70% tracking confidence
- **Audio Initialization**: Robust COM initialization with multiple fallback strategies

## Files

- `index.py` – Main application
- `README.md` – This file

## License

MIT License – Feel free to use and modify for your projects

## Author

[samaksh619](https://github.com/samaksh619)

---

**Enjoy gesture-based volume control!** 🎵

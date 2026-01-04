# pywin32cap

A Python utility for capturing application windows on Windows, with support for background/minimized window capture.

## Features

- Capture specific windows by title or handle (HWND).
- **Background Capture**: Capture windows even when they are minimized or behind other windows (using `PrintWindow` API).
- **Client Area Support**: Option to capture only the client area (content) or the full window (including title bar and borders).
- **Robust Fallbacks**: Automatically tries multiple capture methods (PrintWindow, BitBlt) to ensure success.
- **Window Management**: Utilities to find windows, check minimization state, and restore windows without stealing focus.

## Installation

1.  Clone the repository.
2.  Create a virtual environment (optional but recommended):
    ```bash
    python -m venv .venv
    .\.venv\Scripts\activate
    ```
3.  Install the package and dependencies:
    ```bash
    pip install -e .
    ```

## Usage

### Basic Usage

```python
from pywin32cap.capture_window import WindowCapture

# Initialize
capturer = WindowCapture()

# Find a window
windows = capturer.find_windows_by_title_partial("Notepad")
if windows:
    hwnd, title = windows[0]
    
    # Capture the client area (content only)
    image = capturer.capture_window_client(hwnd)
    
    if image:
        image.save("capture.png")
```

### Running Examples

The `examples/` directory contains demo scripts to showcase functionality.

1.  **Basic Demo**: Lists visible windows and attempts to capture Notepad (or the first visible window).
    ```bash
    python examples/demo_window_capture.py
    ```

2.  **Animation Capture**: Launches a simple animated Pygame app and captures frames from it in a loop.
    ```bash
    python examples/demo_capture_animation.py
    ```
    This demonstrates the ability to capture a moving application. The results are saved in `examples/captured_frames/`.

## Dependencies

- `pywin32`: For Windows API access.
- `pillow`: For image handling.
- `pygame`: Used only for the animation example.

## How it Works

The core logic resides in `pywin32cap/capture_window.py`. It uses `ctypes` and `pywin32` to interface with Windows GDI and User32 APIs.

- **`capture_window`**: The main entry point. It handles window state (restoring if minimized), creates a memory device context, and uses `PrintWindow` or `BitBlt` to copy the window pixels to a bitmap.
- **`capture_window_client`**: A wrapper that specifically targets the client area. If the direct capture fails or includes borders, it can fall back to capturing the full window and cropping it based on client coordinates.

import time
import logging
import os
import sys

# Add parent directory to path so we can import pywin32cap
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from pywin32cap.capture_window import WindowCapture

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    # 1. Initialize the WindowCapture class
    capturer = WindowCapture()
    
    print("--- Window Capture Demo ---")

    # 2. List some available windows (for demonstration purposes)
    print("\nScanning for windows...")
    windows = capturer.find_windows_by_title_partial("") # Empty string matches all
    visible_windows = []
    
    # Filter for visible windows with titles
    import win32gui
    for hwnd, title in windows:
        if win32gui.IsWindowVisible(hwnd) and title.strip():
            visible_windows.append((hwnd, title))
    
    print(f"Found {len(visible_windows)} visible windows.")
    for i, (hwnd, title) in enumerate(visible_windows[:5]): # Show first 5
        print(f"  [{i}] {title} (HWND: {hwnd})")
    if len(visible_windows) > 5:
        print("  ...")

    if not visible_windows:
        print("No visible windows found to capture.")
        return

    # 3. Select a window to capture
    # For this demo, we'll try to find "Notepad" or fallback to the first visible window
    target_title = "Notepad"
    target_hwnd = None
    
    print(f"\nAttempting to find window with title containing '{target_title}'...")
    notepad_windows = capturer.find_windows_by_title_partial(target_title)
    
    if notepad_windows:
        target_hwnd, found_title = notepad_windows[0]
        print(f"Found target window: '{found_title}' (HWND: {target_hwnd})")
    else:
        # Fallback to the first visible window from our list
        target_hwnd, found_title = visible_windows[0]
        print(f"Target '{target_title}' not found. Using first visible window: '{found_title}' (HWND: {target_hwnd})")

    # 4. Capture the window
    # We can capture just the client area (content) or the full window (with borders/title bar)
    output_filename = "captured_window.png"
    print(f"\nCapturing window to '{output_filename}'...")
    
    # capture_window_client is robust and handles fallbacks
    image = capturer.capture_window_client(target_hwnd, save_file=output_filename)
    
    if image:
        print(f"Success! Image captured. Size: {image.size}")
        print(f"Saved to: {os.path.abspath(output_filename)}")
    else:
        print("Failed to capture window.")

    # 5. Demonstrate other useful methods
    print("\n--- Other Capabilities ---")
    
    # Check if minimized
    is_minimized = capturer.is_window_minimized(target_hwnd)
    print(f"Is window minimized? {is_minimized}")
    
    # Get dimensions
    dims = capturer.get_window_dimensions(target_hwnd, client_only=True)
    if dims:
        # dims format for client_only=True is (0, 0, width, height, width, height)
        print(f"Client dimensions: {dims[2]}x{dims[3]}")

if __name__ == "__main__":
    main()

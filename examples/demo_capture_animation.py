import time
import logging
import os
import sys
import subprocess
import shutil
from datetime import datetime

# Add parent directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from pywin32cap.capture_window import WindowCapture

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    # Setup output directory
    output_dir = os.path.join(os.path.dirname(__file__), "captured_frames")
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    
    print(f"Output directory: {output_dir}")

    # 1. Start the animated app
    app_script = os.path.join(os.path.dirname(__file__), "animated_app.py")
    print(f"Starting app: {app_script}")
    
    # Use the same python interpreter
    process = subprocess.Popen([sys.executable, app_script])
    
    try:
        capturer = WindowCapture()
        target_title = "Window Capture Test App"
        target_hwnd = None
        
        # 2. Wait for window to appear
        print("Waiting for window...")
        for _ in range(10):
            windows = capturer.find_windows_by_title_partial(target_title)
            if windows:
                target_hwnd, title = windows[0]
                print(f"Found window: '{title}' (HWND: {target_hwnd})")
                break
            time.sleep(0.5)
            
        if not target_hwnd:
            print("Could not find target window.")
            return

        # 3. Capture loop
        print("Starting capture loop (5 seconds)...")
        start_time = time.time()
        frame_count = 0
        
        while time.time() - start_time < 5:
            timestamp = datetime.now().strftime("%H%M%S_%f")
            filename = f"frame_{frame_count:04d}_{timestamp}.png"
            filepath = os.path.join(output_dir, filename)
            
            # Capture client area
            img = capturer.capture_window_client(target_hwnd)
            
            if img:
                # Save individual frame
                img.save(filepath)
                print(f"Saved {filename}")
                frame_count += 1
            else:
                print("Failed to capture frame")
                
            # Cap capture rate roughly
            time.sleep(0.1)
            
        print(f"\nCaptured {frame_count} frames.")
        print(f"Check {output_dir} for results.")

    finally:
        # Cleanup
        if process:
            print("Stopping animated app...")
            process.terminate()
            process.wait()

if __name__ == "__main__":
    main()

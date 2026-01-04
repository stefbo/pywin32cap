import win32gui
import win32ui
import win32con
import win32api
from PIL import Image
import ctypes
from ctypes import wintypes
import time
import logging

logger = logging.getLogger(__name__)

# Windows API constants
SW_HIDE = 0
SW_SHOWNORMAL = 1
SW_SHOWMINIMIZED = 2
SW_SHOWMAXIMIZED = 3
SW_SHOWNOACTIVATE = 4  # Key flag!
SW_RESTORE = 9
WS_EX_LAYERED = 0x80000
LWA_ALPHA = 0x2
PW_RENDERFULLCONTENT = 0x00000002
PW_CLIENTONLY = 0x00000001

# SetWindowPos flags
SWP_NOACTIVATE = 0x0010
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOZORDER = 0x0004
SWP_SHOWWINDOW = 0x0040


class WindowCapture:
    def __init__(self):
        self.user32 = ctypes.windll.user32
        self.gdi32 = ctypes.windll.gdi32

    def get_foreground_window(self):
        """Get currently focused window"""
        return win32gui.GetForegroundWindow()

    def restore_focus(self, hwnd):
        """Restore focus to the specified window"""
        if hwnd and hwnd != 0:
            try:
                win32gui.SetForegroundWindow(hwnd)
            except:
                # Fallback - sometimes SetForegroundWindow fails
                self.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)

    def is_window_minimized(self, hwnd):
        """Check if window is minimized"""
        return win32gui.IsIconic(hwnd)

    def make_window_transparent(self, hwnd, alpha=1):
        """Make window transparent (1 = almost invisible, 255 = opaque)"""
        try:
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style | WS_EX_LAYERED)
            self.user32.SetLayeredWindowAttributes(hwnd, 0, alpha, LWA_ALPHA)
            return True
        except Exception as e:
            logger.warning("Failed to make window transparent: %s", e)
            return False

    def restore_window_opacity(self, hwnd):
        """Restore window to normal opacity"""
        try:
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style & ~WS_EX_LAYERED)
            return True
        except Exception as e:
            logger.warning("Failed to restore opacity: %s", e)
            return False

    def restore_window_no_focus(self, hwnd):
        """Restore minimized window without stealing focus"""
        try:
            # Method 1: Use ShowWindow with SW_SHOWNOACTIVATE
            result1 = win32gui.ShowWindow(hwnd, SW_SHOWNOACTIVATE)

            # Method 2: Also use SetWindowPos for extra assurance
            # Get current window rect to maintain size/position
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top

            result2 = self.user32.SetWindowPos(
                hwnd,  # Window handle
                0,  # Insert after (ignored with SWP_NOZORDER)
                left,
                top,  # Position
                width,
                height,  # Size
                SWP_SHOWWINDOW  # Show the window
                | SWP_NOACTIVATE  # Don't activate/focus
                | SWP_NOZORDER,  # Don't change Z-order
            )

            return result1 or result2

        except Exception as e:
            logger.warning("Failed to restore window without focus: %s", e)
            return False

    def get_window_dimensions(self, hwnd, client_only=False):
        """Get window dimensions - either full window or client area only"""
        if client_only:
            # Get client area dimensions (in client coordinates)
            client_rect = win32gui.GetClientRect(hwnd)
            width = client_rect[2] - client_rect[0]
            height = client_rect[3] - client_rect[1]

            # For client area, we don't need screen coordinates
            # Just return the dimensions
            return 0, 0, width, height, width, height
        else:
            # Get full window dimensions including decorations
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            return left, top, right, bottom, width, height

    def capture_window(self, hwnd, client_only=True, save_file=None):
        """
        Capture window and return PIL Image

        Args:
            hwnd: Window handle
            client_only: If True, capture only client area (without decorations)
            save_file: Optional filename to save the image

        Returns:
            PIL.Image object or None if capture failed
        """
        try:
            window_title = win32gui.GetWindowText(hwnd)
            capture_type = "client area" if client_only else "full window"
            logger.info("Starting capture of %s: '%s' (HWND: %s)", capture_type, window_title, hwnd)

            # Store the currently focused window
            original_foreground = self.get_foreground_window()

            # Check window state
            was_minimized = self.is_window_minimized(hwnd)
            was_visible = win32gui.IsWindowVisible(hwnd)

            if was_minimized:
                logger.info("Window is minimized, temporarily restoring without activation")

            # If minimized, restore without activation
            if was_minimized:
                # Make it transparent first so user doesn't see it
                self.make_window_transparent(hwnd, alpha=1)

                # Restore without stealing focus
                if not self.restore_window_no_focus(hwnd):
                    logger.error("Failed to restore window without focus")
                    return None

                # Small delay for window to restore
                time.sleep(0.1)

            # Get window dimensions
            left, top, right, bottom, width, height = self.get_window_dimensions(hwnd, client_only)

            if width <= 0 or height <= 0:
                logger.error("Invalid dimensions: %sx%s", width, height)
                return None

            # Create device contexts
            if client_only:
                hwnd_dc = win32gui.GetDC(hwnd)  # Client area DC
            else:
                hwnd_dc = win32gui.GetWindowDC(hwnd)  # Full window DC

            if not hwnd_dc or hwnd_dc == 0:
                logger.error("Failed to get device context")
                return None

            try:
                mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
                save_dc = mfc_dc.CreateCompatibleDC()
            except Exception as dc_e:
                logger.error("Failed to create compatible DC: %s", dc_e)
                win32gui.ReleaseDC(hwnd, hwnd_dc)
                return None

            # Create bitmap
            try:
                save_bitmap = win32ui.CreateBitmap()
                save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
                old_bitmap = save_dc.SelectObject(save_bitmap)

            except Exception as bmp_e:
                logger.error("Bitmap creation failed: %s", bmp_e)
                save_dc.DeleteDC()
                mfc_dc.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwnd_dc)
                return None

            # Capture attempts with different methods
            success = False
            method_used = None

            if client_only:
                # Method 1: PrintWindow with PW_CLIENTONLY
                result = self.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_CLIENTONLY)
                if result:
                    success = True
                    method_used = "PrintWindow PW_CLIENTONLY"
                else:
                    # Method 2: Standard PrintWindow on client DC
                    result = self.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 0)
                    if result:
                        success = True
                        method_used = "Standard PrintWindow"
                    else:
                        # Method 3: BitBlt fallback
                        result = save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)
                        if result:
                            success = True
                            method_used = "BitBlt"
            else:
                # Method 1: PrintWindow with full content
                result = self.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)
                if result:
                    success = True
                    method_used = "PrintWindow PW_RENDERFULLCONTENT"
                else:
                    # Method 2: BitBlt fallback
                    result = save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)
                    if result:
                        success = True
                        method_used = "BitBlt"

            # Convert to PIL Image if successful
            img = None
            if success:
                logger.info("Capture successful using %s", method_used)

                try:
                    bmpinfo = save_bitmap.GetInfo()
                    bmpstr = save_bitmap.GetBitmapBits(True)

                    # Check for empty/black bitmap
                    if len(bmpstr) == 0:
                        logger.error("Bitmap data is empty")
                        success = False
                    else:
                        img = Image.frombuffer(
                            "RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), bmpstr, "raw", "BGRX", 0, 1
                        )

                        # Optional: save to file
                        if save_file:
                            img.save(save_file)
                            logger.info("Image saved to: %s", save_file)

                except Exception as img_e:
                    logger.error("PIL Image creation failed: %s", img_e)
                    success = False
            else:
                logger.error("All capture methods failed for %s", capture_type)

            # Cleanup with error handling
            try:
                if "old_bitmap" in locals() and old_bitmap:
                    save_dc.SelectObject(old_bitmap)
                if "save_bitmap" in locals():
                    win32gui.DeleteObject(save_bitmap.GetHandle())
                if "save_dc" in locals():
                    save_dc.DeleteDC()
                if "mfc_dc" in locals():
                    mfc_dc.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwnd_dc)
            except Exception as cleanup_e:
                logger.warning("Cleanup error (non-fatal): %s", cleanup_e)

            # Restore original window state
            if was_minimized:
                win32gui.ShowWindow(hwnd, SW_SHOWMINIMIZED)
                self.restore_window_opacity(hwnd)

            # Restore focus
            if original_foreground and original_foreground != 0:
                self.restore_focus(original_foreground)

            return img

        except Exception as e:
            logger.error("Fatal error in capture_window: %s", e)

            # Emergency focus restoration
            if "original_foreground" in locals() and original_foreground:
                try:
                    self.restore_focus(original_foreground)
                except:
                    logger.error("Emergency focus restoration failed")

            return None

    def capture_window_full(self, hwnd, save_file=None):
        """Capture full window including decorations"""
        return self.capture_window(hwnd, client_only=False, save_file=save_file)

    def capture_window_client(self, hwnd, save_file=None):
        """Capture only client area (without decorations) with automatic fallback"""
        # Try direct client area capture first
        img = self.capture_window(hwnd, client_only=True, save_file=None)

        if img is None:
            logger.info("Direct client capture failed, falling back to crop method")
            # Fall back to crop method
            return self.capture_window_client_crop(hwnd, save_file)
        else:
            # Success with direct method
            if save_file:
                img.save(save_file)
            return img

    def capture_window_client_crop(self, hwnd, save_file=None):
        """
        Reliable client area capture by cropping full window capture
        This method is more reliable but slightly slower
        """
        try:
            window_title = win32gui.GetWindowText(hwnd)
            logger.info("Crop-based client area capture of: '%s' (HWND: %s)", window_title, hwnd)

            # Capture full window first
            full_img = self.capture_window_full(hwnd)
            if not full_img:
                logger.error("Could not capture full window for cropping")
                return None

            # Get window and client rectangles
            window_rect = win32gui.GetWindowRect(hwnd)
            client_rect = win32gui.GetClientRect(hwnd)

            # Calculate client area position within window
            client_top_left = win32gui.ClientToScreen(hwnd, (0, 0))

            # Calculate offsets from window top-left to client area top-left
            left_border = client_top_left[0] - window_rect[0]
            top_border = client_top_left[1] - window_rect[1]

            client_width = client_rect[2] - client_rect[0]
            client_height = client_rect[3] - client_rect[1]

            # Validate crop boundaries
            if (
                left_border < 0
                or top_border < 0
                or left_border + client_width > full_img.width
                or top_border + client_height > full_img.height
            ):
                logger.error(
                    "Invalid crop dimensions: borders=(%s, %s), client=%sx%s, window=%s",
                    left_border,
                    top_border,
                    client_width,
                    client_height,
                    full_img.size,
                )
                return None

            # Crop the full window image to get client area
            client_img = full_img.crop(
                (left_border, top_border, left_border + client_width, top_border + client_height)
            )

            logger.info("Crop method succeeded: %s", client_img.size)

            if save_file:
                client_img.save(save_file)

            return client_img

        except Exception as e:
            logger.error("Crop capture error: %s", e)
            return None

    def find_window_by_title(self, title):
        """Find window by exact title"""
        hwnd = win32gui.FindWindow(None, title)
        return hwnd if hwnd != 0 else None

    def find_windows_by_title_partial(self, partial_title):
        """Find windows containing partial title"""
        windows = []

        def enum_windows_proc(hwnd, lParam):
            if win32gui.IsWindow(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if partial_title.lower() in window_title.lower():
                    windows.append((hwnd, window_title))
            return True

        win32gui.EnumWindows(enum_windows_proc, None)
        return windows

    def find_window_by_pid(self, pid):
        """Find window by process ID"""
        windows = []

        def enum_windows_proc(hwnd, lParam):
            if win32gui.IsWindow(hwnd):
                _, window_pid = win32gui.GetWindowThreadProcessId(hwnd)
                if window_pid == pid:
                    title = win32gui.GetWindowText(hwnd)
                    windows.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_windows_proc, None)
        return windows

    def list_all_windows(self):
        """List all windows for debugging"""
        logger.info("Listing all windows")

        current_foreground = self.get_foreground_window()
        window_count = 0

        def enum_windows_proc(hwnd, lParam):
            nonlocal window_count
            if win32gui.IsWindow(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title.strip():
                    _, pid = win32gui.GetWindowThreadProcessId(hwnd)
                    is_visible = win32gui.IsWindowVisible(hwnd)
                    is_minimized = win32gui.IsIconic(hwnd)
                    is_foreground = hwnd == current_foreground

                    status_flags = []
                    if is_visible:
                        status_flags.append("visible")
                    if is_minimized:
                        status_flags.append("minimized")
                    if is_foreground:
                        status_flags.append("foreground")

                    status = ", ".join(status_flags) if status_flags else "none"
                    logger.info("Window: '%s' (HWND: %s, PID: %s, Status: %s)", title, hwnd, pid, status)
                    window_count += 1
            return True

        win32gui.EnumWindows(enum_windows_proc, None)
        logger.info("Total windows with titles: %s", window_count)

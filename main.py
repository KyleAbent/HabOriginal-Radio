import time
import uiautomation as auto
import psutil
import ctypes
import tkinter as tk
import threading
import pythoncom


def get_windows_with_timeout(timeout=5):
    result = []
    start_time = time.time()

    def enum_windows_proc(hwnd, lParam):
        if time.time() - start_time > timeout:
            return False
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
            result.append((hwnd, buff.value))
        return True

    ctypes.windll.user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)(enum_windows_proc),
                                     0)
    return result


def get_apple_music_info():
    # Initialize COM for this thread
    pythoncom.CoInitialize()

    try:
        apple_music_process = None
        for process in psutil.process_iter(['name']):
            if process.info['name'].lower() == 'applemusic.exe':
                apple_music_process = process
                break

        if not apple_music_process:
            return None

        windows = get_windows_with_timeout()

        apple_music_window = None
        for hwnd, window_name in windows:
            if "Apple Music" == window_name:
                apple_music_window = auto.ControlFromHandle(hwnd)
                break

        if not apple_music_window:
            return None

        lcd_control = apple_music_window.GroupControl(AutomationId="LCD")
        if not lcd_control.Exists(3):
            return None

        song_name = ""
        artist_name = ""
        album_name = ""

        song_pane = lcd_control.PaneControl(AutomationId="myScrollViewer", foundIndex=1)
        if song_pane.Exists(3):
            song_text_control = song_pane.TextControl(AutomationId="ScrollingText")
            if song_text_control.Exists(1):
                song_name = song_text_control.Name

        artist_album_pane = lcd_control.PaneControl(AutomationId="myScrollViewer", foundIndex=2)
        if artist_album_pane.Exists(3):
            artist_album_text_control = artist_album_pane.TextControl(AutomationId="ScrollingText")
            if artist_album_text_control.Exists(1):
                artist_album_text = artist_album_text_control.Name
                parts = artist_album_text.split(" — ")
                artist_name = parts[0] if len(parts) > 0 else ""
                album_name = parts[1] if len(parts) > 1 else ""

        transport_bar = apple_music_window.GroupControl(AutomationId="TransportBar")
        is_paused = None
        if transport_bar.Exists(3):
            play_pause_button = transport_bar.ButtonControl(AutomationId="TransportControl_PlayPauseStop")
            if play_pause_button.Exists():
                is_paused = play_pause_button.Name == "Play"

        if not song_name or not artist_name:
            return None

        return {
            "song": song_name,
            "artist": artist_name,
            "album": album_name,
            "is_paused": is_paused
        }
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()


def update_title(root):
    while True:
        try:
            info = get_apple_music_info()
            if info:
                # title = f"{info['song']} - {info['artist']} ({'Paused' if info['is_paused'] else 'Playing'})"
                #Artist first
                title = f"{info['artist']} - {info['song']} ({'Paused' if info['is_paused'] else 'Playing'})"
                root.title(title)
            else:
                root.title("HabOriginal Radio - Not playing or unable to fetch info")
        except Exception as e:
            print(f"Error updating title: {e}")
            root.title("Apple Music - Error fetching info")
        time.sleep(1)  # Update every second


def main():
    root = tk.Tk()
    root.geometry("300x100")  # Set a small window size
    root.title("HabOriginal Radio - Initializing...")

    label = tk.Label(root,text="Apple Music Metadata Extractor\nMinimize this window to keep it running in the background.")
    label.pack(expand=True)

    # Start the update thread
    update_thread = threading.Thread(target=update_title, args=(root,), daemon=True)
    update_thread.start()

    root.mainloop()


if __name__ == "__main__":
    main()
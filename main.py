import time
import uiautomation as auto
import psutil
import ctypes


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

    ctypes.windll.user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)(enum_windows_proc), 0)
    return result


def print_control_structure(control, depth=0):
    print("  " * depth + f"{control.ControlTypeName} '{control.Name}' (AutomationId: {control.AutomationId})")
    for child in control.GetChildren():
        print_control_structure(child, depth + 1)


def get_apple_music_info():
    print("Starting to find Apple Music process...")
    apple_music_process = None
    for process in psutil.process_iter(['name']):
        if process.info['name'].lower() == 'applemusic.exe':
            apple_music_process = process
            break

    if not apple_music_process:
        print("Apple Music is not running.")
        return None

    print(f"Apple Music process found (PID: {apple_music_process.pid})")

    print("Fetching windows...")
    windows = get_windows_with_timeout()
    print(f"Number of windows: {len(windows)}")

    apple_music_window = None
    for hwnd, window_name in windows:
        if "Apple Music" == window_name:
            print(f"Found Apple Music window: {window_name}")
            apple_music_window = auto.ControlFromHandle(hwnd)
            break

    if not apple_music_window:
        print("Apple Music window not found.")
        return None

    print(f"Apple Music window found: {apple_music_window.Name}")
    print("Printing Apple Music window structure:")
    # print_control_structure(apple_music_window)

    # Try to find the song information
    lcd_control = apple_music_window.GroupControl(AutomationId="LCD")
    if not lcd_control.Exists(3):
        print("LCD control not found.")
        return None

    print("LCD control found.")

    # Locate the song name, artist, and album information within the LCD control
    song_name = ""
    artist_name = ""
    album_name = ""

    song_pane = lcd_control.PaneControl(AutomationId="myScrollViewer", foundIndex=1)
    if song_pane.Exists(3):
        song_text_control = song_pane.TextControl(AutomationId="ScrollingText")
        if song_text_control.Exists(1):
            song_name = song_text_control.Name
            print(f"Song name: {song_name}")

    artist_album_pane = lcd_control.PaneControl(AutomationId="myScrollViewer", foundIndex=2)
    if artist_album_pane.Exists(3):
        artist_album_text_control = artist_album_pane.TextControl(AutomationId="ScrollingText")
        if artist_album_text_control.Exists(1):
            artist_album_text = artist_album_text_control.Name
            parts = artist_album_text.split(" — ")
            artist_name = parts[0] if len(parts) > 0 else ""
            album_name = parts[1] if len(parts) > 1 else ""
            print(f"Artist: {artist_name}, Album: {album_name}")

    # Get playback status
    transport_bar = apple_music_window.GroupControl(AutomationId="TransportBar")
    is_paused = None
    if transport_bar.Exists(3):
        play_pause_button = transport_bar.ButtonControl(AutomationId="TransportControl_PlayPauseStop")
        if play_pause_button.Exists():
            is_paused = play_pause_button.Name == "Play"
            print(f"Playback status: {'Paused' if is_paused else 'Playing'}")

    if not song_name or not artist_name:
        print("Could not find complete song information.")
        return None

    return {
        "song": song_name,
        "artist": artist_name,
        "album": album_name,
        "is_paused": is_paused
    }


def main():
    print("Apple Music Metadata Extractor for Windows (Detailed Version)")
    print("Press Ctrl+C to stop")

    try:
        while True:
            print("\n--- Fetching Apple Music info ---")
            info = get_apple_music_info()
            if info:
                print(f"Extracted info: {info}")
                print(f"Now playing: {info['song']} by {info['artist']} from {info['album']}")
                print(f"Paused: {info['is_paused']}")
            else:
                print("Unable to fetch current track info or no track is playing.")
            print("--- Waiting 10 seconds before next update ---")
            time.sleep(10)  # Wait for 10 seconds before checking again
    except KeyboardInterrupt:
        print("\nStopping metadata extraction.")


if __name__ == "__main__":
    main()

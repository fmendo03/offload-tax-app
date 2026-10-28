"""Offload as a desktop program.

Starts the app on this computer and opens it in its own window - no browser tabs, no address bar, no
black console window. Closing the window stops the app. Double-click the Offload shortcut (made by
make-desktop-shortcut.ps1) or this file.
"""
import os
import sys
import time
import json
import socket
import subprocess
import threading
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)

# With no console (pythonw) there is nowhere to print, so everything goes to a log file instead.
LOG_PATH = os.path.join(HERE, 'offload-desktop.log')
try:
    _log = open(LOG_PATH, 'a', encoding='utf-8', buffering=1)
    sys.stdout = _log
    sys.stderr = _log
except OSError:
    pass


def say(message):
    print(time.strftime('%Y-%m-%d %H:%M:%S'), message)


def error_box(text):
    """A plain Windows message box, since there is no console to read an error from."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, 'Offload', 0x10)
    except Exception:
        pass


def port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex(('127.0.0.1', port)) == 0


def offload_answering(port):
    try:
        with urllib.request.urlopen(f'http://127.0.0.1:{port}/auth/status', timeout=2) as r:
            return json.loads(r.read().decode('utf-8')).get('success') is True
    except Exception:
        return False


def window_open_on_profile(profile):
    """Is a Chrome/Edge window still open on Offload's own profile? Asked of Windows itself, because when a
    window is already open there, a new launch hands over to it and returns at once - which must not be
    mistaken for the window having been closed."""
    marker = os.path.basename(os.path.normpath(profile))
    try:
        out = subprocess.run(
            ['powershell', '-NoProfile', '-Command',
             "(Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'chrome.exe' -or $_.Name -eq 'msedge.exe') -and $_.CommandLine -like '*%s*' } | Measure-Object).Count" % marker],
            capture_output=True, text=True, timeout=20, creationflags=0x08000000)
        return int((out.stdout or '0').strip() or 0) > 0
    except Exception:
        return False


def find_browser():
    candidates = [
        os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        os.path.join(os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google', 'Chrome', 'Application', 'chrome.exe'),
        os.path.join(os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'), 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
        os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    ]
    return next((c for c in candidates if os.path.isfile(c)), None)


def main():
    port = int(os.environ.get('PORT', 5000))
    url = f'http://127.0.0.1:{port}'
    started_here = False

    if port_in_use(port):
        if not offload_answering(port):
            error_box(f'Port {port} is being used by something else, so Offload could not start.')
            return
        say('Offload is already running - opening a window to it.')
    else:
        os.environ['FLASK_DEBUG'] = '0'
        try:
            import app as offload
        except Exception as e:
            say(f'Could not load the app: {e!r}')
            error_box(f'Offload could not start:\n\n{e}\n\nDetails are in offload-desktop.log')
            return
        threading.Thread(
            target=lambda: offload.app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False, threaded=True),
            daemon=True
        ).start()
        started_here = True
        for _ in range(120):
            if offload_answering(port):
                break
            time.sleep(0.5)
        else:
            error_box('Offload did not start in time. Details are in offload-desktop.log')
            return
        say(f'Offload is running at {url}')

    browser = find_browser()
    if not browser:
        error_box('Offload needs Google Chrome or Microsoft Edge installed to show its window.')
        return
    # Its own profile, so the window is a separate program-like process we can wait on, and the
    # microphone permission for the lessons is remembered.
    profile = os.path.join(os.environ.get('DATA_DIR', HERE), '.offload-window')
    window = subprocess.Popen([
        browser, f'--app={url}', f'--user-data-dir={profile}', '--window-size=1500,950',
        '--no-first-run', '--no-default-browser-check', '--disable-features=Translate'
    ])
    say(f'Window opened with {os.path.basename(browser)}')
    window.wait()
    say('Window process returned.')
    if started_here:
        # A window left open from before takes the new one over, so the launch returns straight away
        # even though the app is still on screen - keep running until no window is left.
        while window_open_on_profile(profile):
            time.sleep(3)
        say('Window closed.')
        time.sleep(2)   # let the page finish saving what it was doing
        os._exit(0)


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        say(f'Unexpected error: {exc!r}')
        error_box(f'Offload hit a problem:\n\n{exc}')

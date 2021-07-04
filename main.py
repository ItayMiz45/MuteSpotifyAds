import pygetwindow as gw
from pycaw.pycaw import AudioUtilities
import win32process
import psutil
import time


def get_spotify_pid():
    sessions = AudioUtilities.GetAllSessions()

    for ses in sessions:
        proc = ses.Process
        if proc:
            if proc.name() == "Spotify.exe":
                return proc.pid


def mute_proc(pid):
    ses = AudioUtilities.GetProcessSession(pid)
    ses.SimpleAudioVolume.SetMute(1, None)


def unmute_proc(pid):
    ses = AudioUtilities.GetProcessSession(pid)
    ses.SimpleAudioVolume.SetMute(0, None)


def get_spotify_window_handle(spotify_pid):
    for win in gw.getAllWindows():  # loop through all windows
        curr_win_pid = win32process.GetWindowThreadProcessId(win._hWnd)[1]  # returns (tid, pid)
        if curr_win_pid != spotify_pid:  # check if this window is spotify window
            continue

        if win.title in ("", "MSCTFIME UI", "Default IME"):  # bad titles
            continue

        return win


def main():
    while True:
        spotify_pid = get_spotify_pid()
        spotify_win = get_spotify_window_handle(spotify_pid)

        while psutil.pid_exists(spotify_pid):  # check if process is still running
            if spotify_win.title == "Advertisement":
                mute_proc(spotify_pid)
            else:
                unmute_proc(spotify_pid)

            time.sleep(1)

        time.sleep(10)


if __name__ == '__main__':
    main()

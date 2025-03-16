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
from ctypes import wintypes, CFUNCTYPE, POINTER
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QSpinBox, 
                           QLineEdit, QGroupBox, QGridLayout, QDialog, 
                           QListWidget, QListWidgetItem, QMessageBox, QScrollArea,
                           QProgressBar, QTextEdit)
from PyQt6.QtCore import Qt, QTimer

# Keyboard hook için gerekli yapılar
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
LRESULT = ctypes.c_long
WPARAM = ctypes.c_ulonglong
LPARAM = ctypes.c_ulonglong

# Keyboard hook callback tipi
HOOKPROC = CFUNCTYPE(LRESULT, ctypes.c_int, WPARAM, LPARAM)

class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p)
    ]

class WindowsAPI:
    def __init__(self, main_window):
        self.main_window = main_window
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32
        self._hooks = {}  # window_title: hook_id
        self._hook_callbacks = {}  # window_title: callback_function

    def _key_to_vk(self, key):
        # F tuşları için özel işlem
        if key.upper().startswith('F') and len(key) <= 3:
            try:
                f_num = int(key[1:])
                if 1 <= f_num <= 24:  # F1-F24
                    return win32con.VK_F1 + f_num - 1
            except ValueError:
                pass
        
        # Normal tuşlar için
        return ord(key.upper())

    def send_key(self, window_title, key, pid):
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                vk_code = self._key_to_vk(key)
                win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                self.main_window.log(f"Tuş gönderildi: {key} -> {window_title} (PID: {pid})")
                return True
            return False
        except Exception as e:
            self.main_window.log(f"Tuş gönderme hatası: {e}")
            return False

    def set_hook(self, window_title, pid):
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                # Hook callback fonksiyonu
                def hook_callback(nCode, wParam, lParam):
                    if nCode >= 0 and wParam == WM_KEYDOWN:
                        return self.user32.CallNextHookEx(None, nCode, wParam, lParam)
                    return self.user32.CallNextHookEx(None, nCode, wParam, lParam)
                
                # Callback'i sakla (garbage collection'dan korumak için)
                callback = HOOKPROC(hook_callback)
                self._hook_callbacks[window_title] = callback
                
                # Hook'u kur
                hook_id = self.user32.SetWindowsHookExA(
                    WH_KEYBOARD_LL,
                    callback,
                    None,
                    0  # Global hook için 0
                )
                
                if hook_id:
                    self._hooks[window_title] = hook_id
                    self.main_window.log(f"Hook kuruldu: {window_title} (PID: {pid})")
                    return True
                else:
                    error = ctypes.WinError()
                    self.main_window.log(f"Hook kurma hatası: {error}")
            return False
        except Exception as e:
            self.main_window.log(f"Hook kurma hatası: {e}")
            return False

    def remove_hook(self, window_title, pid):
        try:
            hook_id = self._hooks.get(window_title)
            if hook_id:
                if self.user32.UnhookWindowsHookEx(hook_id):
                    del self._hooks[window_title]
                    if window_title in self._hook_callbacks:
                        del self._hook_callbacks[window_title]
                    self.main_window.log(f"Hook kaldırıldı: {window_title} (PID: {pid})")
                    return True
            return False
        except Exception as e:
            self.main_window.log(f"Hook kaldırma hatası: {e}")
            return False

class WowClient:
    def __init__(self, window_title, pid, windows_api):
        self.window_title = window_title
        self.pid = pid
        self.key = '1'
        self.min_delay = 30
        self.max_delay = 60
        self.running = False
        self.thread = None
        self.windows_api = windows_api
        self.current_delay = 0
        self.next_delay = 0
        self._stop_event = threading.Event()
        self.hooked = False

    def send_key(self):
        try:
            if self.hooked:
                self.windows_api.send_key(self.window_title, self.key, self.pid)
                self.current_delay = 0
                self.next_delay = random.uniform(self.min_delay, self.max_delay)
        except Exception as e:
            self.windows_api.main_window.log(f"Tuş gönderme hatası: {e}")
            self.stop()

    def advertise_loop(self):
        try:
            while not self._stop_event.is_set():
                try:
                    self.send_key()
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
                # Hook'u kur
                if self.windows_api.set_hook(self.window_title, self.pid):
                    self.hooked = True
                else:
                    raise Exception("Hook kurulamadı")
                
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
                
                # Sonra hook'u kaldır
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
        self.setFixedSize(400, 220)
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_progress)

    def setup_ui(self):
        self.setTitle(f"Client: {self.client.window_title} (PID: {self.client.pid})")
        layout = QVBoxLayout()
        layout.setSpacing(5)

        # Üst kısım - Tuş ayarı
        key_layout = QHBoxLayout()
        key_layout.addWidget(QLabel("Tuş:"))
        self.key_input = QLineEdit(self.client.key)
        self.key_input.setMaxLength(3)  # F10 gibi tuşlar için 3 karakter
        self.key_input.setFixedWidth(50)  # Biraz daha geniş
        key_layout.addWidget(self.key_input)
        key_layout.addStretch()
        layout.addLayout(key_layout)

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
        self.remove_button = QPushButton("Kaldır")
        
        self.start_button.clicked.connect(self.toggle_client)
        self.remove_button.clicked.connect(self.remove_client)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.remove_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def update_progress(self):
        if self.client.running and self.client.next_delay > 0:
            self.client.current_delay += 0.1
            progress = (self.client.current_delay / self.client.next_delay) * 100
            self.progress.setValue(min(int(progress), 100))

    def toggle_client(self):
        try:
            if not self.client.running:
                self.client.key = self.key_input.text()
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

class ProcessSelector(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("WoW Client Seç")
        self.setModal(True)
        self.setup_ui()
        self.refresh_processes()
        self.resize(400, 500)

    def setup_ui(self):
        layout = QVBoxLayout()
        
        self.process_list = QListWidget()
        self.process_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        layout.addWidget(QLabel("WoW Client'ları seçin (Ctrl ile çoklu seçim yapabilirsiniz):"))
        layout.addWidget(self.process_list)

        button_layout = QHBoxLayout()
        refresh_btn = QPushButton("Yenile")
        select_btn = QPushButton("Ekle")
        refresh_btn.clicked.connect(self.refresh_processes)
        select_btn.clicked.connect(self.accept)
        button_layout.addWidget(refresh_btn)
        button_layout.addWidget(select_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def refresh_processes(self):
        self.process_list.clear()
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    # WoW veya benzer isimli processleri bul
                    if 'wow' in proc.info['name'].lower() or 'world' in proc.info['name'].lower():
                        pid = proc.info['pid']
                        hwnd = None
                        
                        def callback(h, extra):
                            try:
                                if win32process.GetWindowThreadProcessId(h)[1] == pid:
                                    if win32gui.IsWindowVisible(h):
                                        extra[0] = h
                                        return False
                            except Exception:
                                pass
                            return True
                        
                        extra = [None]
                        win32gui.EnumWindows(callback, extra)
                        hwnd = extra[0]
                        
                        if hwnd:
                            title = win32gui.GetWindowText(hwnd)
                            if title:  # Boş başlıklı pencereleri atla
                                item_text = f"{title} (PID: {pid})"
                                item = QListWidgetItem(item_text)
                                item.setData(Qt.ItemDataRole.UserRole, (title, pid))
                                self.process_list.addItem(item)
                except (psutil.NoSuchProcess, psutil.AccessDenied, Exception) as e:
                    continue
        except Exception as e:
            if self.parent():
                self.parent().log(f"Process listesi hatası: {e}")

    def get_selected_processes(self):
        return [item.data(Qt.ItemDataRole.UserRole) 
                for item in self.process_list.selectedItems()]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WoW Advertiser")
        self.client_controls = []
        self.windows_api = WindowsAPI(self)
        self.setup_ui()
        self.setFixedSize(1280, 900)

    def setup_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()

        # Client ekleme butonu
        header_layout = QHBoxLayout()
        add_client_btn = QPushButton("Client Ekle")
        add_client_btn.clicked.connect(self.add_client)
        header_layout.addWidget(add_client_btn)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # Scroll area ve grid layout için container
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        self.grid_widget = QWidget()
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setSpacing(10)
        
        scroll.setWidget(self.grid_widget)
        main_layout.addWidget(scroll)

        # Minimal Log box
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFixedHeight(100)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #999;
                border-radius: 4px;
                padding: 5px;
                font-family: monospace;
                font-size: 11px;
            }
        """)
        main_layout.addWidget(self.log_text)
        
        # Log box'ı alta yapıştır
        main_layout.setStretchFactor(scroll, 1)
        main_layout.setStretchFactor(self.log_text, 0)

        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def log(self, message):
        try:
            self.log_text.append(f"{time.strftime('%H:%M:%S')} - {message}")
            self.log_text.verticalScrollBar().setValue(
                self.log_text.verticalScrollBar().maximum()
            )
        except Exception as e:
            print(f"Log hatası: {e}")

    def add_client(self):
        try:
            dialog = ProcessSelector(self)
            if dialog.exec():
                selected_processes = dialog.get_selected_processes()
                for process in selected_processes:
                    control = ClientControl(process, self.windows_api, self)
                    self.client_controls.append(control)
                    
                    # Grid'e yerleştirme (soldan sağa, yukarıdan aşağı)
                    idx = len(self.client_controls) - 1
                    row = idx // 3
                    col = idx % 3
                    self.grid_layout.addWidget(control, row, col)
                    self.log(f"Client eklendi: {process[0]} (PID: {process[1]})")
        except Exception as e:
            self.log(f"Client ekleme hatası: {e}")
            self.log(traceback.format_exc())

    def remove_client_control(self, control):
        try:
            # Grid'den kaldır
            self.grid_layout.removeWidget(control)
            control.hide()  # Hemen gizle
            
            # Listeden kaldır
            if control in self.client_controls:
                self.client_controls.remove(control)
                self.log(f"Client kaldırıldı: {control.client.window_title} (PID: {control.client.pid})")
            
            # Kalan kontrolleri yeniden düzenle
            for i, ctrl in enumerate(self.client_controls):
                row = i // 3
                col = i % 3
                self.grid_layout.addWidget(ctrl, row, col)
            
        except Exception as e:
            self.log(f"Client kaldırma hatası: {e}")
            self.log(traceback.format_exc())

    def stop_all_clients(self):
        for control in self.client_controls[:]:
            try:
                control.stop_client()
            except Exception as e:
                self.log(f"Client durdurma hatası: {e}")

    def closeEvent(self, event):
        try:
            self.stop_all_clients()
            self.log("Uygulama kapatılıyor...")
            event.accept()
        except Exception as e:
            self.log(f"Uygulama kapatma hatası: {e}")
            event.accept()

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    try:
        window = MainWindow()
        window.show()
        window.log("Uygulama başlatıldı")
        sys.exit(app.exec())
    except Exception as e:
        print(f"Uygulama hatası: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

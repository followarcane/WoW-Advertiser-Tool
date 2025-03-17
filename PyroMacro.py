import sys
import os
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
from datetime import datetime
from ctypes import wintypes
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QSpinBox, 
                           QLineEdit, QGroupBox, QGridLayout, QDialog, 
                           QListWidget, QListWidgetItem, QMessageBox, QScrollArea,
                           QProgressBar, QTextEdit)
from PyQt6.QtCore import Qt, QTimer

# Log dosyası ayarları
LOG_FILE = os.path.join(os.path.expanduser("~"), "Documents", "WoWAdvertiser.log")

def write_log(message):
    """Dosyaya log yazar"""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{timestamp} - {message}\n")
    except Exception as e:
        print(f"Log yazma hatası: {e}")

# Başlangıç logunu yaz
write_log("=== Uygulama başlatılıyor ===")

class WindowsAPI:
    def __init__(self, main_window):
        self.main_window = main_window
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32
        self._target_windows = {}  # window_title: hwnd
        write_log("WindowsAPI başlatıldı")

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
                
                # Pencere thread'ini bul
                target_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
                
                # Tuşu gönder (PostThreadMessage kullanarak)
                result = self.user32.PostThreadMessageW(
                    target_thread_id, 
                    win32con.WM_KEYDOWN, 
                    vk_code, 
                    0
                )
                time.sleep(0.05)
                result2 = self.user32.PostThreadMessageW(
                    target_thread_id, 
                    win32con.WM_KEYUP, 
                    vk_code, 
                    0
                )
                
                log_msg = f"Tuş gönderildi: {key} -> {window_title} (PID: {pid}, ThreadID: {target_thread_id}) - Sonuç: {result},{result2}"
                self.main_window.log(log_msg)
                write_log(log_msg)
                return True
            return False
        except Exception as e:
            error_msg = f"Tuş gönderme hatası: {e}"
            self.main_window.log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())
            return False

    def set_hook(self, window_title, pid):
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                self._target_windows[window_title] = hwnd
                log_msg = f"Client hazır: {window_title} (PID: {pid})"
                self.main_window.log(log_msg)
                write_log(log_msg)
                return True
            return False
        except Exception as e:
            error_msg = f"Client hazırlama hatası: {e}"
            self.main_window.log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())
            return False

    def remove_hook(self, window_title, pid):
        try:
            if window_title in self._target_windows:
                del self._target_windows[window_title]
                log_msg = f"Client durduruldu: {window_title} (PID: {pid})"
                self.main_window.log(log_msg)
                write_log(log_msg)
                return True
            return False
        except Exception as e:
            error_msg = f"Client durdurma hatası: {e}"
            self.main_window.log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())
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
        write_log(f"WowClient oluşturuldu: {window_title} (PID: {pid})")

    def send_key(self):
        try:
            if self.hooked:
                self.windows_api.send_key(self.window_title, self.key, self.pid)
                self.current_delay = 0
                self.next_delay = random.uniform(self.min_delay, self.max_delay)
        except Exception as e:
            error_msg = f"Tuş gönderme hatası: {e}"
            self.windows_api.main_window.log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())
            self.stop()

    def advertise_loop(self):
        try:
            write_log(f"Advertise loop başlatıldı: {self.window_title} (PID: {self.pid})")
            while not self._stop_event.is_set():
                try:
                    self.send_key()
                    # Her 0.1 saniyede bir kontrol et
                    for _ in range(int(self.next_delay * 10)):
                        if self._stop_event.is_set():
                            return
                        time.sleep(0.1)
                except Exception as e:
                    error_msg = f"Loop hatası (PID: {self.pid}): {e}"
                    self.windows_api.main_window.log(error_msg)
                    write_log(error_msg)
                    write_log(traceback.format_exc())
                    break
        except Exception as e:
            error_msg = f"Thread hatası (PID: {self.pid}): {e}"
            self.windows_api.main_window.log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())

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
                error_msg = f"Client başlatma hatası (PID: {self.pid}): {e}"
                self.windows_api.main_window.log(error_msg)
                write_log(error_msg)
                write_log(traceback.format_exc())
                self.hooked = False
                self.running = False

    def stop(self):
        if self.running:
            try:
                # Thread'i durdur
                self._stop_event.set()
                if self.thread and self.thread.is_alive():
                    self.thread.join(0.5)  # 0.5 saniye bekle
                
                # Hook'u kaldır
                if self.hooked:
                    self.windows_api.remove_hook(self.window_title, self.pid)
                    self.hooked = False
                
                self.running = False
            except Exception as e:
                error_msg = f"Client durdurma hatası (PID: {self.pid}): {e}"
                self.windows_api.main_window.log(error_msg)
                write_log(error_msg)
                write_log(traceback.format_exc())
                self.running = False

class ClientControl(QGroupBox):
    def __init__(self, process_info, windows_api, parent=None):
        super().__init__(parent)
        self.client = WowClient(process_info[0], process_info[1], windows_api)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_progress)
        self.setup_ui()

    def setup_ui(self):
        self.setTitle(f"Client: {self.client.window_title} (PID: {self.client.pid})")
        layout = QVBoxLayout()
        layout.setSpacing(5)

        # Tuş ayarı
        key_layout = QHBoxLayout()
        key_layout.addWidget(QLabel("Tuş:"))
        self.key_input = QLineEdit(self.client.key)
        self.key_input.setMaxLength(3)  # F10 gibi tuşlar için 3 karakter
        self.key_input.setFixedWidth(50)
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
        self.start_button.clicked.connect(self.toggle_client)
        self.remove_button = QPushButton("Kaldır")
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
            error_msg = f"Client toggle hatası (PID: {self.client.pid}): {e}"
            self.window().log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())
            self.stop_client()

    def stop_client(self):
        try:
            self.client.stop()
            self.start_button.setText("Başlat")
            self.timer.stop()
            self.progress.setValue(0)
        except Exception as e:
            error_msg = f"Client durdurma hatası (PID: {self.client.pid}): {e}"
            self.window().log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())

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
            error_msg = f"Client kaldırma hatası (PID: {self.client.pid}): {e}"
            self.window().log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())

class ProcessSelector(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("WoW Process Seçici")
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)
        self.setup_ui()
        self.populate_process_list()

    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Process listesi
        self.process_list = QListWidget()
        self.process_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(self.process_list)
        
        # Butonlar
        button_layout = QHBoxLayout()
        select_button = QPushButton("Seç")
        select_button.clicked.connect(self.accept)
        cancel_button = QPushButton("İptal")
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(select_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)

    def populate_process_list(self):
        try:
            self.process_list.clear()
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    # WoW process'lerini bul
                    if 'wow' in proc.info['name'].lower() or 'world of warcraft' in proc.info['name'].lower():
                        pid = proc.info['pid']
                        
                        # Process'in pencere başlığını bul
                        def callback(hwnd, hwnds):
                            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                                if found_pid == pid:
                                    title = win32gui.GetWindowText(hwnd)
                                    if title:  # Boş başlıkları atla
                                        hwnds.append((title, pid))
                            return True
                        
                        hwnds = []
                        win32gui.EnumWindows(callback, hwnds)
                        
                        for title, pid in hwnds:
                            item = QListWidgetItem(f"{title} (PID: {pid})")
                            item.setData(Qt.ItemDataRole.UserRole, (title, pid))
                            self.process_list.addItem(item)
                            write_log(f"Process bulundu: {title} (PID: {pid})")
                except (psutil.NoSuchProcess, psutil.AccessDenied, Exception) as e:
                    write_log(f"Process erişim hatası: {e}")
                    continue
        except Exception as e:
            error_msg = f"Process listesi hatası: {e}"
            if self.parent():
                self.parent().log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())

    def get_selected_processes(self):
        selected = [item.data(Qt.ItemDataRole.UserRole) 
                for item in self.process_list.selectedItems()]
        write_log(f"Seçilen processler: {selected}")
        return selected

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WoW Advertiser")
        self.client_controls = []
        self.windows_api = WindowsAPI(self)
        self.setup_ui()
        self.setFixedSize(1280, 900)
        write_log("MainWindow oluşturuldu")

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
        write_log("UI kuruldu")

    def log(self, message):
        try:
            self.log_text.append(f"{time.strftime('%H:%M:%S')} - {message}")
            self.log_text.verticalScrollBar().setValue(
                self.log_text.verticalScrollBar().maximum()
            )
        except Exception as e:
            error_msg = f"Log hatası: {e}"
            print(error_msg)
            write_log(error_msg)

    def add_client(self):
        try:
            write_log("Client ekleme başlatıldı")
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
                    log_msg = f"Client eklendi: {process[0]} (PID: {process[1]})"
                    self.log(log_msg)
                    write_log(log_msg)
        except Exception as e:
            error_msg = f"Client ekleme hatası: {e}"
            self.log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())

    def remove_client_control(self, control):
        try:
            # Grid'den kaldır
            self.grid_layout.removeWidget(control)
            control.hide()  # Hemen gizle
            
            # Listeden kaldır
            if control in self.client_controls:
                self.client_controls.remove(control)
                log_msg = f"Client kaldırıldı: {control.client.window_title} (PID: {control.client.pid})"
                self.log(log_msg)
                write_log(log_msg)
            
            # Kalan kontrolleri yeniden düzenle
            for i, ctrl in enumerate(self.client_controls):
                row = i // 3
                col = i % 3
                self.grid_layout.addWidget(ctrl, row, col)
            
        except Exception as e:
            error_msg = f"Client kaldırma hatası: {e}"
            self.log(error_msg)
            write_log(error_msg)
            write_log(traceback.format_exc())

    def stop_all_clients(self):
        for control in self.client_controls[:]:
            try:
                control.stop_client()
            except Exception as e:
                error_msg = f"Client durdurma hatası: {e}"
                self.log(error_msg)
                write_log(error_msg)

    def closeEvent(self, event):
        try:
            self.stop_all_clients()
            log_msg = "Uygulama kapatılıyor..."
            self.log(log_msg)
            write_log(log_msg)
            event.accept()
        except Exception as e:
            error_msg = f"Uygulama kapatma hatası: {e}"
            self.log(error_msg)
            write_log(error_msg)
            event.accept()

def main():
    try:
        write_log("QApplication başlatılıyor")
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        
        write_log("MainWindow oluşturuluyor")
        window = MainWindow()
        window.show()
        window.log("Uygulama başlatıldı")
        write_log("MainWindow gösterildi")
        
        sys.exit(app.exec())
    except Exception as e:
        error_msg = f"Uygulama hatası: {e}"
        print(error_msg)
        write_log(error_msg)
        write_log(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = f"Ana program hatası: {e}"
        print(error_msg)
        write_log(error_msg)
        write_log(traceback.format_exc())

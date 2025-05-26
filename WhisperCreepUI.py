import sys
import subprocess
import os
import whisper # OpenAI's Whisper
import torch
import tempfile
import shutil
import logging # For logging
import datetime # For timestamps
import traceback # For error tracebacks
import time # For timing
import threading
import win32api
import win32file
import win32con
import socket
import configparser # For configuration file
from pathlib import Path
# Import video_frame_snatcher module
from video_frame_snatcher import VideoFrameSnatcher
# Import YouTube Caption Fetcher
from youtube_captionfetcher import YouTubeCaptionFetcher

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QMessageBox, QComboBox, QGroupBox,
    QSpacerItem, QSizePolicy, QRadioButton, QButtonGroup, QMenuBar, QMenu,
    QMainWindow, QDialog, QLineEdit, QSystemTrayIcon, QStatusBar, QToolTip
)
from PySide6.QtCore import Qt, QThread, Signal, QObject, QTimer, QEvent
from PySide6.QtGui import QIntValidator, QIcon, QAction

# --- Global Variables & Constants ---
APP_VERSION = "1.4.5_ParanoidLogPath_OriginalGUI"
transcription_in_progress = False

# --- Setup Python Logging (Initial: Console Only) ---
# FileHandler will be added dynamically per run.
logging.basicConfig(
    level=logging.DEBUG,
    format="[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout) 
    ]
)
logger_app = logging.getLogger("WhisperCreepApp.GUI")
logger_worker = logging.getLogger("WhisperCreepApp.Worker")

logger_app.info(f"--- LOGGER INITIALIZED (Console Only). SCRIPT STARTING. Version: {APP_VERSION} ---")
try: SCRIPT_DIR_FOR_INFO = os.path.dirname(os.path.abspath(__file__))
except NameError: SCRIPT_DIR_FOR_INFO = os.getcwd()
logger_app.info(f"Script directory (for info): {SCRIPT_DIR_FOR_INFO}")
logger_app.info(f"Current working directory: {os.getcwd()}")
logger_app.info("Run-specific file logging will be set up in Downloads folder when processing starts.")

# --- Global State Management ---
class TranscriptionStateManager(QObject):
    state_changed = Signal(bool)  # Emits new state when changed
    
    def __init__(self):
        super().__init__()
        self._lock = threading.Lock()
        self._state = False
        self._active_monitors = set()
    
    @property
    def is_transcribing(self):
        with self._lock:
            return self._state
    
    def set_transcribing(self, value):
        with self._lock:
            old_state = self._state
            self._state = value
            if old_state != value:
                self.state_changed.emit(value)
    
    def register_monitor(self, monitor_id):
        with self._lock:
            self._active_monitors.add(monitor_id)
            return len(self._active_monitors)
    
    def unregister_monitor(self, monitor_id):
        with self._lock:
            self._active_monitors.discard(monitor_id)
            return len(self._active_monitors)

# Create global state manager instance
transcription_state = TranscriptionStateManager()

# --- Helper Functions ---
def bring_console_to_front():
    if os.name == 'nt':
        try:
            from ctypes import windll
            hwnd = windll.kernel32.GetConsoleWindow()
            if hwnd: windll.user32.SetForegroundWindow(hwnd); windll.user32.ShowWindow(hwnd, 9)
        except Exception as e: logger_app.warning(f"Could not bring console to front: {e}", exc_info=False)

def format_timestamp_for_transcript(seconds: float) -> str:
    abs_seconds = abs(seconds); hours = int(abs_seconds / 3600); minutes = int((abs_seconds % 3600) / 60)
    secs = int(abs_seconds % 60); millis = int((abs_seconds - int(abs_seconds)) * 1000)
    return f"{'-' if seconds < 0 else ''}{hours:02d}:{minutes:02d}:{secs:02d}.{millis:03d}"

def get_download_folder_path():
    """
    Determines user's Downloads folder path with increased robustness.
    Ensures the path is absolute and the directory exists or can be created.
    Logs its decision-making process.
    Returns an absolute path string or None if determination fails.
    """
    logger_app.info("Attempting to determine absolute Downloads folder path...")
    final_downloads_path = None

    # Strategy 1: USERPROFILE environment variable (primarily for Windows)
    user_profile_env = os.environ.get('USERPROFILE')
    logger_app.debug(f"Value of os.environ.get('USERPROFILE'): '{user_profile_env}'")

    if user_profile_env and user_profile_env.strip(): # Check if not None and not empty/whitespace
        # Ensure user_profile_env itself is an absolute and valid directory
        abs_user_profile = os.path.abspath(user_profile_env)
        if os.path.isdir(abs_user_profile):
            candidate_path = os.path.join(abs_user_profile, 'Downloads')
            logger_app.debug(f"Using USERPROFILE strategy. Candidate Downloads path: '{candidate_path}'")
            # os.path.abspath will ensure it's truly absolute, even if candidate_path was somehow relative
            final_downloads_path = os.path.abspath(candidate_path)
        else:
            logger_app.warning(f"USERPROFILE ('{user_profile_env}', abs: '{abs_user_profile}') is set but is not a valid directory. Falling back.")
    else:
        logger_app.warning(f"USERPROFILE environment variable not found or is empty. Falling back to home directory strategy.")

    # Strategy 2: Home directory (cross-platform fallback)
    if not final_downloads_path: 
        home_dir = os.path.expanduser('~') # This should always return an absolute path
        logger_app.debug(f"Value of os.path.expanduser('~'): '{home_dir}'")
        if home_dir and os.path.isdir(home_dir): # Check if home_dir is valid
            candidate_path = os.path.join(home_dir, 'Downloads')
            logger_app.debug(f"Using home directory strategy. Candidate Downloads path: '{candidate_path}'")
            final_downloads_path = os.path.abspath(candidate_path) # Ensure absolute
        else:
            logger_app.critical(f"CRITICAL: Could not determine a valid base for Downloads. USERPROFILE strategy failed, and home directory ('{home_dir}') is invalid or not found.")
            return None # Cannot proceed

    if not final_downloads_path: # Should be caught by above, but as a safeguard
        logger_app.critical("CRITICAL: Failed to construct any candidate path for Downloads folder.")
        return None

    logger_app.info(f"Target Downloads folder path determined as: '{final_downloads_path}'")

    # Ensure the determined Downloads directory exists or create it
    try:
        if not os.path.exists(final_downloads_path):
            logger_app.info(f"Downloads directory '{final_downloads_path}' does not exist. Creating now.")
            os.makedirs(final_downloads_path, exist_ok=True) # exist_ok=True handles race conditions
            logger_app.info(f"Downloads directory created/ensured: '{final_downloads_path}'")
        elif not os.path.isdir(final_downloads_path): # Path exists but is not a directory
            logger_app.error(f"CRITICAL: Path '{final_downloads_path}' exists but is NOT a directory. Cannot use as Downloads folder.")
            return None
        else: # Path exists and is a directory
            logger_app.debug(f"Downloads directory '{final_downloads_path}' already exists and is a directory.")
        
        return final_downloads_path
    except Exception as e:
        logger_app.error(f"CRITICAL: Exception creating/accessing Downloads directory '{final_downloads_path}': {e}", exc_info=True)
        return None

# --- Processing Indicator Dialog (from v1.4.4) ---
class ProcessingIndicatorDialog(QDialog):
    kill_process_requested = Signal() 
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Working..."); self.setModal(False) 
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint | Qt.WindowStaysOnTopHint)
        main_layout = QVBoxLayout(self); self.spinner_label = QLabel("📀 Processing...")
        self.spinner_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.spinner_label.font(); font.setPointSize(18); self.spinner_label.setFont(font)
        main_layout.addWidget(self.spinner_label)
        self.kill_button_dialog = QPushButton("Kill Process")
        self.kill_button_dialog.setStyleSheet("background-color: orange; color: black; padding: 5px; border-radius: 3px;")
        self.kill_button_dialog.clicked.connect(self._request_kill)
        main_layout.addWidget(self.kill_button_dialog)
        self.spinner_chars = ["📀", "💿", "🎬", "🎵"] 
        self.char_index = 0; self.timer = QTimer(self); self.timer.timeout.connect(self._update_spinner)
        self.setFixedSize(280, 150); logger_app.debug("ProcessingIndicatorDialog initialized.")
    def _update_spinner(self): self.char_index=(self.char_index+1)%len(self.spinner_chars); self.spinner_label.setText(f"{self.spinner_chars[self.char_index]} Processing...")
    def _request_kill(self): logger_app.info("Kill Process on Dialog clicked."); self.kill_process_requested.emit() 
    def start_animation(self): logger_app.info("ProcessingIndicatorDialog animation started."); self.char_index=0; self.spinner_label.setText(f"{self.spinner_chars[self.char_index]} Processing..."); self.timer.start(300); self.show()
    def stop_animation_and_close(self): logger_app.info("ProcessingIndicatorDialog stop_animation_and_close called."); self.timer.stop(); self.close() 
    def closeEvent(self, event): logger_app.debug("ProcessingIndicatorDialog closeEvent."); self.timer.stop(); super().closeEvent(event)

# --- Monitor Folder Dialog (move this above WhisperCreepInterface) ---
class MonitorFolderDialog(QDialog):
    _instances = set()  # Track all instances
    _retry_pause_active = False  # Track if any monitor is in retry pause
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.monitor_id = id(self)  # Unique identifier for this instance
        MonitorFolderDialog._instances.add(self)
        
        # Check for duplicate monitoring of same folder
        self.folder_path = None
        self.output_path = None  # Add output path storage
        self.setWindowTitle("Monitor Folder for Video")
        self.setModal(False)
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint | Qt.WindowMinimizeButtonHint)
        
        # Apply dark theme styling
        self.setStyleSheet("""
            QDialog {
                background-color: black;
                color: white;
            }
            QLabel {
                color: white;
                font-family: Calibri;
                font-size: 11pt;
            }
            QPushButton {
                background-color: #333333;
                color: white;
                padding: 5px;
                border-radius: 3px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #444444;
            }
            QPushButton:disabled {
                background-color: #222222;
                color: #888888;
            }
            QLineEdit {
                background-color: #222222;
                color: white;
                border: 1px solid #444444;
                padding: 3px;
                border-radius: 2px;
            }
        """)
        
        layout = QVBoxLayout(self)

        # Folder selection
        self.folder_label = QLabel("Select folder to monitor:")
        self.folder_path_label = QLabel("No folder selected.")
        self.folder_button = QPushButton("Browse")
        self.folder_button.setStyleSheet("background-color: royalblue;")
        self.folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.folder_label)
        layout.addWidget(self.folder_path_label)
        layout.addWidget(self.folder_button)

        # Output location
        self.output_label = QLabel("Select output location:")
        self.output_path_label = QLabel("No location selected.")
        self.output_button = QPushButton("Browse")
        self.output_button.setStyleSheet("background-color: royalblue;")
        self.output_button.clicked.connect(self.select_output)
        layout.addWidget(self.output_label)
        layout.addWidget(self.output_path_label)
        layout.addWidget(self.output_button)

        # Time increment
        self.time_label = QLabel("Set time increment (seconds):")
        self.time_input = QLineEdit()
        self.time_input.setValidator(QIntValidator(1, 3600))  # 1 second to 1 hour
        layout.addWidget(self.time_label)
        layout.addWidget(self.time_input)

        # Button layout for Apply, Save, and Close
        btn_layout = QHBoxLayout()
        self.apply_button = QPushButton("Apply Settings")
        self.apply_button.setStyleSheet("background-color: #555555;")
        self.apply_button.clicked.connect(self.apply_settings)
        btn_layout.addWidget(self.apply_button)
        
        self.save_button = QPushButton("Save Settings")
        self.save_button.setStyleSheet("background-color: green;")
        self.save_button.clicked.connect(self.save_settings)
        btn_layout.addWidget(self.save_button)
        
        self.close_button = QPushButton("Close")
        self.close_button.setStyleSheet("background-color: crimson;")
        self.close_button.clicked.connect(self.close)
        btn_layout.addWidget(self.close_button)
        layout.addLayout(btn_layout)

        # Start/Stop monitoring
        self.start_button = QPushButton("Start Monitoring")
        self.start_button.setStyleSheet("background-color: purple;")
        self.start_button.clicked.connect(self.start_monitoring)
        self.start_button.setEnabled(False)  # Disabled until settings are saved
        layout.addWidget(self.start_button)

        self.stop_button = QPushButton("Stop Monitoring")
        self.stop_button.setStyleSheet("background-color: orange; color: black;")
        self.stop_button.clicked.connect(self.stop_monitoring)
        self.stop_button.setEnabled(False)
        layout.addWidget(self.stop_button)

        self.monitoring = False
        self.monitor_thread = None
        self.stop_event = threading.Event()
        self.processed_files = set()
        self.staged_settings = {}
        self.settings = None  # Will be set when settings are saved
        self.tray_icon = None
        self.setup_tray_icon()
        self.log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "monitor_system.log")
        self.last_file_size = None
        self.size_check_time = None
        self.temp_output_path = None
        
        # Keep a reference to the parent window
        self.parent_window = parent

    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        # Try to load the custom icon
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "WhisperCreepICO.ico")
            logger_app.info(f"Attempting to load icon from: {icon_path}")
            
            if os.path.exists(icon_path):
                icon = QIcon(icon_path)
                if not icon.isNull():
                    self.tray_icon.setIcon(icon)
                    logger_app.info(f"Successfully loaded custom icon from: {icon_path}")
                else:
                    raise Exception("Icon loaded but is null")
            else:
                raise FileNotFoundError(f"Icon file not found at: {icon_path}")
        except Exception as e:
            logger_app.error(f"Error loading custom icon: {e}")
            self.tray_icon.setIcon(QIcon.fromTheme("system-run", QIcon.fromTheme("applications-system")))
            logger_app.info("Using system default icon")
        
        self.tray_icon.setToolTip("WhisperCreep Folder Monitor")
        menu = QMenu()
        restore_action = QAction("Restore Monitor", self)
        restore_action.triggered.connect(self.showNormal)
        menu.addAction(restore_action)
        quit_action = QAction("Quit Monitoring", self)
        quit_action.triggered.connect(self.stop_monitoring)
        menu.addAction(quit_action)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self.on_tray_activated)

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.showNormal()
            self.raise_()
            self.activateWindow()

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            # Check if this folder is already being monitored
            for instance in MonitorFolderDialog._instances:
                if instance != self and instance.monitoring and instance.folder_path == folder:
                    QMessageBox.warning(self, "Folder Already Monitored",
                        f"This folder is already being monitored by another instance.\n\n"
                        f"Please stop the other monitoring session first.")
                    return
            self.folder_path = folder
            self.folder_path_label.setText(folder)  # Update the QLabel instead of the string

    def select_output(self):
        output = QFileDialog.getExistingDirectory(self, "Select Output Location")
        if output:
            self.output_path = output  # Store the path
            self.output_path_label.setText(output)  # Update the display

    def apply_settings(self):
        # Stage the settings without saving or closing
        if not self.folder_path or not self.output_path or not self.time_input.text():
            QMessageBox.warning(self, "Input Error", "Please fill all fields.")
            return
        
        self.staged_settings = {
            'folder': self.folder_path,
            'output': self.output_path,
            'interval': self.time_input.text()
        }
        QMessageBox.information(self, "Settings Staged", "Settings have been staged but not saved. Click 'Save Settings' to apply them permanently.")

    def save_settings(self):
        if not self.staged_settings:
            QMessageBox.warning(self, "No Settings Staged", "Please click 'Apply Settings' first to stage your settings.")
            return
            
        self.settings = self.staged_settings.copy()
        self.start_button.setEnabled(True)  # Enable start button when settings are saved
        QMessageBox.information(self, "Settings Saved", "Your monitoring settings have been saved. You can now start monitoring.")

    def check_network_path(self, path):
        """Check if a path is a network path and verify its accessibility."""
        try:
            # Check if it's a UNC path or mapped drive
            if path.startswith('\\\\') or (len(path) > 2 and path[1:3] == ':\\' and win32file.WNetGetConnection(path[0] + ':')):
                # Try to get network path info
                try:
                    win32file.WNetGetResourceInformation(path)
                except Exception as e:
                    logger_app.error(f"Network path validation failed: {e}")
                    return False
                
                # Test write access with a small file
                test_file = os.path.join(path, f'.wc_test_{int(time.time())}.tmp')
                try:
                    with open(test_file, 'w') as f:
                        f.write('test')
                    os.remove(test_file)
                except Exception as e:
                    logger_app.error(f"Network path write test failed: {e}")
                    return False
            return True
        except Exception as e:
            logger_app.error(f"Network path check failed: {e}")
            return False

    def check_folder_permissions(self, path):
        """Check if we have read/write permissions for a folder."""
        try:
            # Test read access
            os.listdir(path)
            # Test write access
            test_file = os.path.join(path, f'.wc_perm_test_{int(time.time())}.tmp')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            return True
        except Exception as e:
            logger_app.error(f"Permission check failed for {path}: {e}")
            return False

    def verify_file_consistency(self, file_path, check_interval=2, num_checks=3):
        """Check if file size remains consistent over multiple checks."""
        try:
            current_size = os.path.getsize(file_path)
            if self.last_file_size is None:
                self.last_file_size = current_size
                self.size_check_time = time.time()
                return False
            
            if time.time() - self.size_check_time < check_interval:
                return False
                
            if current_size != self.last_file_size:
                self.last_file_size = current_size
                self.size_check_time = time.time()
                return False
                
            # Size remained consistent
            self.last_file_size = None
            self.size_check_time = None
            return True
        except Exception as e:
            logger_app.error(f"File consistency check failed: {e}")
            return False

    def start_monitoring(self):
        if not self.settings:
            QMessageBox.warning(self, "No Settings", "Please save your settings before starting monitoring.")
            return
            
        if not self.settings['folder'] or not self.settings['output'] or not self.settings['interval']:
            QMessageBox.warning(self, "Input Error", "Please fill all fields and save settings.")
            return

        # Check network paths
        if not self.check_network_path(self.settings['folder']):
            QMessageBox.critical(self, "Network Error", 
                "Cannot access the selected folder. It may be a network path that is unavailable or inaccessible.")
            return
            
        if not self.check_network_path(self.settings['output']):
            QMessageBox.critical(self, "Network Error", 
                "Cannot access the output location. It may be a network path that is unavailable or inaccessible.")
            return

        # Check permissions
        if not self.check_folder_permissions(self.settings['folder']):
            QMessageBox.critical(self, "Permission Error", 
                "Cannot access the selected folder. Please check your permissions.")
            return
            
        if not self.check_folder_permissions(self.settings['output']):
            QMessageBox.critical(self, "Permission Error", 
                "Cannot access the output location. Please check your permissions.")
            return

        # Register with state manager
        active_monitors = transcription_state.register_monitor(self.monitor_id)
        logger_app.info(f"Registered monitor {self.monitor_id}. Total active monitors: {active_monitors}")
        
        if active_monitors > 1:
            QMessageBox.warning(self, "Multiple Monitors",
                "Another monitoring session is already active.\n\n"
                "Please stop the other session first.")
            transcription_state.unregister_monitor(self.monitor_id)
            return
            
        self.monitoring = True
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.save_button.setEnabled(False)
        self.apply_button.setEnabled(False)
        self.stop_event.clear()
        
        # Ensure tray icon is visible before hiding window
        self.tray_icon.show()
        
        # Start monitoring thread
        self.monitor_thread = threading.Thread(target=self.monitor_folder, daemon=True)
        self.monitor_thread.start()
        
        # Hide window but don't close it
        self.hide()
        
        # Show a notification that monitoring has started
        self.tray_icon.showMessage(
            "WhisperCreep Monitor",
            "Folder monitoring has started. The application will continue running in the system tray.",
            QSystemTrayIcon.Information,
            3000
        )
        
        # Log the start of monitoring with details
        self.log_event(f"MONITORING STARTED | Folder: {self.settings['folder']} | Output: {self.settings['output']} | Interval: {self.settings['interval']}s")

    def stop_monitoring(self):
        if not self.monitoring:
            return
            
        self.monitoring = False
        self.stop_event.set()
        
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)
            if self.monitor_thread.is_alive():
                logger_app.warning("Monitor thread did not stop gracefully")
        
        # Unregister from state manager and log the count
        remaining = transcription_state.unregister_monitor(self.monitor_id)
        logger_app.info(f"Unregistered monitor {self.monitor_id}. Remaining active monitors: {remaining}")
        
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(True)
        self.apply_button.setEnabled(True)
        
        self.tray_icon.hide()
        QMessageBox.information(self, "Stopped", "Folder monitoring has been stopped.")
        self.showNormal()
        
        # Log the stop event
        self.log_event("MONITORING STOPPED")

    def monitor_folder(self):
        try:
            folder_to_monitor = self.settings['folder']
            output_location = self.settings['output']
            time_increment = int(self.settings['interval'])
            
            logger_app.info(f"Starting folder monitoring: {folder_to_monitor}")
            logger_app.info(f"Output location: {output_location}")
            logger_app.info(f"Time increment: {time_increment} seconds")
            self.log_event(f"MONITORING ACTIVATED - Will check every {time_increment} seconds")
            
            # Clear processed files list on start
            self.processed_files = set()
            
            while not self.stop_event.is_set():
                try:
                    # Check if a transcription is in progress from the main app
                    if transcription_state.is_transcribing:
                        logger_app.info("Manual transcription in progress, pausing monitoring...")
                        self.tray_icon.setToolTip("WhisperCreep Monitor - Monitoring On Hold")
                        self.log_event("Monitoring paused due to manual transcription")
                        # Wait until transcription is complete
                        while transcription_state.is_transcribing and not self.stop_event.is_set():
                            time.sleep(1)
                        if not self.stop_event.is_set():
                            logger_app.info("Manual transcription completed, waiting full interval before resuming...")
                            self.tray_icon.setToolTip("WhisperCreep Monitor - Waiting to Resume")
                            self.log_event("Manual transcription completed, waiting interval before resuming")
                            # Wait the full interval before resuming
                            for _ in range(time_increment):
                                if self.stop_event.is_set():
                                    break
                                time.sleep(1)
                            if not self.stop_event.is_set():
                                logger_app.info("Resuming monitoring after interval...")
                                self.tray_icon.setToolTip("WhisperCreep Monitor - Active")
                                self.log_event("Monitoring resumed after waiting interval")
                    
                    # Verify folders still exist
                    if not os.path.exists(folder_to_monitor):
                        logger_app.error(f"Monitor folder does not exist: {folder_to_monitor}")
                        self.log_event(f"ERROR: Monitor folder missing: {folder_to_monitor}")
                        self.stop_monitoring()
                        return
                        
                    if not os.path.exists(output_location):
                        logger_app.error(f"Output location does not exist: {output_location}")
                        self.log_event(f"ERROR: Output folder missing: {output_location}")
                        self.stop_monitoring()
                        return
                        
                    # Scan for video files
                    try:
                        logger_app.info(f"Scanning for video files in {folder_to_monitor}")
                        video_files = [f for f in os.listdir(folder_to_monitor) if f.lower().endswith((
                            '.mp4', '.mov', '.avi', '.wmv', '.mkv'))]
                        logger_app.info(f"Found {len(video_files)} video files, {len(self.processed_files)} already processed")
                    except Exception as e:
                        logger_app.error(f"Error scanning folder {folder_to_monitor}: {e}")
                        self.log_event(f"ERROR scanning folder: {str(e)}")
                        time.sleep(time_increment)
                        continue
                    
                    unprocessed = [f for f in video_files if f not in self.processed_files]
                    if unprocessed:
                        logger_app.info(f"Found {len(unprocessed)} unprocessed videos: {', '.join(unprocessed)}")
                    
                    # Process only one file at a time
                    if unprocessed:
                        file_name = unprocessed[0]  # Take the first unprocessed file
                        file_path = os.path.join(folder_to_monitor, file_name)
                        if os.path.isfile(file_path):
                            start_time = time.time()
                            export_name = os.path.splitext(file_name)[0] + "_transcript.txt"
                            export_path = os.path.join(output_location, export_name)
                            
                            self.log_event(f"START: {file_name}", src_path=file_path, dest_path=export_path)
                            self.tray_icon.setToolTip(f"WhisperCreep Monitor - Transcribing: {file_name}")
                            logger_app.info(f"Starting transcription process for {file_path} -> {export_path}")
                            
                            # Try to transcribe with retries
                            transcription_result = self.transcribe_video(file_path, export_path)
                            elapsed = time.time() - start_time
                            
                            if transcription_result:
                                self.log_event(f"DONE: {file_name} | Export: {export_name} | Time: {elapsed:.2f}s", src_path=file_path, dest_path=export_path)
                                logger_app.info(f"Transcription SUCCESSFUL in {elapsed:.2f}s")
                                self.processed_files.add(file_name)
                            else:
                                self.log_event(f"SKIPPED: {file_name} due to transcription failures", src_path=file_path)
                                logger_app.warning(f"Transcription FAILED in {elapsed:.2f}s")
                            
                            # Wait the full interval before processing next file
                            logger_app.info(f"Waiting {time_increment} seconds before next file...")
                            self.tray_icon.setToolTip(f"WhisperCreep Monitor - Waiting {time_increment}s before next file")
                            for _ in range(time_increment):
                                if self.stop_event.is_set():
                                    break
                                time.sleep(1)
                        else:
                            logger_app.warning(f"File {file_path} not found or not a file")
                            self.log_event(f"WARNING: File not found: {file_name}")
                            self.processed_files.add(file_name)  # Mark as processed to avoid retrying
                    else:
                        # No new files, wait the interval before next scan
                        self.tray_icon.setToolTip(f"WhisperCreep Monitor - No new files, checking in {time_increment}s")
                        logger_app.info(f"No unprocessed files found. Checking again in {time_increment}s")
                        for _ in range(time_increment):
                            if self.stop_event.is_set():
                                break
                            time.sleep(1)
                        
                except Exception as e:
                    logger_app.error(f"Error in monitoring loop: {e}", exc_info=True)
                    self.log_event(f"ERROR in monitoring: {str(e)[:150]}")
                    self.tray_icon.setToolTip(f"WhisperCreep Monitor - Error: {str(e)[:50]}...")
                    time.sleep(time_increment)  # Wait before retrying
                    
        except Exception as e:
            logger_app.error(f"Fatal error in monitor_folder: {e}", exc_info=True)
            self.log_event(f"FATAL ERROR: {str(e)[:150]}")
            self.stop_monitoring()

    def check_file_lock(self, file_path, timeout=5):
        """Check if a file is locked by attempting to open it with exclusive access."""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Try to open the file with exclusive access
                with open(file_path, 'a'):
                    return False  # File is not locked
            except IOError:
                # File is locked, wait a bit and try again
                time.sleep(0.5)
        return True  # File is locked after timeout

    def transcribe_video(self, file_path, export_path, max_retries=2):
        retry_count = 0
        import whisper
        import torch
        import shutil
        import tempfile
        model_name = "base"  # You can make this configurable if desired
        device = "cuda" if torch.cuda.is_available() else "cpu"
        if device == "cpu":
            logger_app.warning("CUDA not available, falling back to CPU. This will be slower.")
        else:
            logger_app.info("CUDA is available, using GPU acceleration.")
        while retry_count <= max_retries:
            try:
                if MonitorFolderDialog._retry_pause_active:
                    logger_app.info("Another process is in retry pause state, waiting...")
                    time.sleep(2)
                    continue

                if not self.check_folder_permissions(os.path.dirname(file_path)):
                    logger_app.error(f"Lost read permission for {file_path}")
                    self.log_event(f"Lost read permission: {os.path.basename(file_path)}")
                    return False

                if not self.check_folder_permissions(os.path.dirname(export_path)):
                    logger_app.error(f"Lost write permission for {export_path}")
                    self.log_event(f"Lost write permission: {os.path.basename(export_path)}")
                    return False

                if os.path.exists(export_path) and self.check_file_lock(export_path):
                    logger_app.warning(f"Output file {export_path} is locked. Waiting for access...")
                    self.log_event(f"Waiting for output file access: {os.path.basename(export_path)}")
                    time.sleep(2)
                    retry_count += 1
                    continue

                if self.check_file_lock(file_path):
                    logger_app.warning(f"Input file {file_path} is locked. Waiting for access...")
                    self.log_event(f"Waiting for input file access: {os.path.basename(file_path)}")
                    time.sleep(2)
                    retry_count += 1
                    continue

                if not self.verify_file_consistency(file_path):
                    logger_app.info(f"Waiting for file {file_path} to stabilize...")
                    time.sleep(2)
                    continue

                MonitorFolderDialog._retry_pause_active = True
                self.temp_output_path = export_path + '.partial'
                temp_dir = None
                temp_audio_path = None
                try:
                    # Set global transcription state
                    transcription_state.set_transcribing(True)
                    self.log_event(f"TRANSCRIPTION PROCESS STARTING for {os.path.basename(file_path)}")
                    logger_app.info(f"[MONITOR] === BEGINNING ACTUAL TRANSCRIPTION PROCESS FOR {file_path} ===")
                    
                    # 1. Extract audio from video
                    temp_dir = tempfile.mkdtemp(prefix="wc_temp_")
                    temp_audio_path = os.path.join(temp_dir, "temp_audio.wav")
                    cmd = [
                        "ffmpeg", "-i", file_path, "-y", "-vn", "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1", temp_audio_path
                    ]
                    logger_app.info(f"[MONITOR] FFMPEG cmd: {' '.join(cmd)}")
                    res = subprocess.run(cmd, capture_output=True, text=True, errors='replace', creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
                    if res.returncode != 0:
                        logger_app.error(f"[MONITOR] FFMPEG audio extraction error: {res.stderr}")
                        self.log_event(f"FFMPEG extract error: {res.stderr[:150]}...")
                        return False
                    if not os.path.exists(temp_audio_path) or os.path.getsize(temp_audio_path) == 0:
                        logger_app.error(f"[MONITOR] FFMPEG OK but output missing/empty: '{temp_audio_path}'")
                        self.log_event("FFMPEG produced no audio data.")
                        return False
                    logger_app.info("[MONITOR] Audio extraction OK.")

                    # 2. Load Whisper and transcribe
                    logger_app.info(f"[MONITOR] Loading Whisper model '{model_name}' on '{device}'.")
                    model = whisper.load_model(model_name, device=device)
                    logger_app.info("[MONITOR] Model loaded.")
                    logger_app.info(f"[MONITOR] Starting transcription... (Audio: '{temp_audio_path}')")
                    res_dict = model.transcribe(temp_audio_path, verbose=True)
                    logger_app.info("[MONITOR] transcribe() call completed.")
                    if res_dict is None:
                        logger_app.error("[MONITOR] Transcription result is None!")
                        self.log_event("Transcription returned no result.")
                        return False

                    # 3. Write transcript to temp file, then move to final output
                    segments = res_dict.get("segments")
                    full_txt = res_dict.get("text", "").strip()
                    with open(self.temp_output_path, "w", encoding="utf-8") as f:
                        f.write(f"Transcription of {file_path}\n\n")  # Add header so it's clear it's a real transcript
                        if segments and any(seg['text'].strip() for seg in segments):
                            for i, seg in enumerate(segments):
                                f.write(f"[{format_timestamp_for_transcript(seg['start'])} --> {format_timestamp_for_transcript(seg['end'])}] {seg['text'].strip()}\n")
                                if i % 100 == 0:
                                    f.flush()
                        elif full_txt:
                            f.write(full_txt + ("\n" if full_txt else ""))
                        else:
                            error_msg = "[ERROR] Whisper returned no transcript data. Check logs for details."
                            logger_app.error(f"[MONITOR] {error_msg}")
                            f.write(error_msg + "\n")
                    # Move to final output
                    if os.path.exists(export_path):
                        os.remove(export_path)
                    os.rename(self.temp_output_path, export_path)
                    self.temp_output_path = None
                    logger_app.info(f"[MONITOR] Transcript written successfully to {export_path}")

                    # 4. Move video file to output folder
                    try:
                        dest_video_path = os.path.join(os.path.dirname(export_path), os.path.basename(file_path))
                        if os.path.abspath(file_path) != os.path.abspath(dest_video_path):
                            shutil.move(file_path, dest_video_path)
                            logger_app.info(f"[MONITOR] Video file moved to output folder: {dest_video_path}")
                            self.log_event(f"VIDEO MOVED to output folder: {dest_video_path}")
                    except Exception as e:
                        logger_app.warning(f"[MONITOR] Failed to move video file: {e}")
                        self.log_event(f"Failed to move video file: {str(e)[:150]}")

                    # Clean up temp audio
                    if temp_audio_path and os.path.exists(temp_audio_path):
                        try:
                            os.remove(temp_audio_path)
                        except Exception as e:
                            logger_app.warning(f"[MONITOR] Failed to clean up temp audio: {e}")
                    if temp_dir and os.path.exists(temp_dir):
                        try:
                            shutil.rmtree(temp_dir)
                        except Exception as e:
                            logger_app.warning(f"[MONITOR] Failed to clean up temp dir: {e}")

                    return True
                finally:
                    # Always clear transcription state
                    transcription_state.set_transcribing(False)
                    logger_app.info(f"[MONITOR] === TRANSCRIPTION PROCESS COMPLETED FOR {file_path} ===")
                    MonitorFolderDialog._retry_pause_active = False
                    if self.temp_output_path and os.path.exists(self.temp_output_path):
                        try:
                            os.remove(self.temp_output_path)
                        except Exception as e:
                            logger_app.warning(f"Failed to clean up temp file {self.temp_output_path}: {e}")
                        self.temp_output_path = None
            except Exception as e:
                retry_count += 1
                if retry_count <= max_retries:
                    wait_time = 2 ** retry_count  # Exponential backoff
                    logger_app.warning(f"Transcription attempt {retry_count} failed: {e}. Retrying in {wait_time} seconds...")
                    self.log_event(f"Retry {retry_count}/{max_retries} for {os.path.basename(file_path)}")
                    time.sleep(wait_time)
                else:
                    logger_app.error(f"All transcription attempts failed for {file_path}: {e}")
                    self.log_event(f"FAILED: {os.path.basename(file_path)} after {max_retries} retries")
                    return False

    def log_event(self, message, src_path=None, dest_path=None):
        # Enhanced logging: include src and dest if provided and if START or DONE
        if (message.startswith("START:") or message.startswith("DONE:")) and src_path and dest_path:
            message = f"{message} | Src: {src_path} | Dest: {dest_path}"
        with open(self.log_path, 'a', encoding='utf-8') as logf:
            logf.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

    def closeEvent(self, event):
        if self.monitoring:
            reply = QMessageBox.question(self, 'Confirm Close',
                'Monitoring is active. Are you sure you want to close?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
            self.stop_monitoring()
        
        # Clean up any remaining temp files
        if self.temp_output_path and os.path.exists(self.temp_output_path):
            try:
                os.remove(self.temp_output_path)
            except Exception as e:
                logger_app.warning(f"Failed to clean up temp file {self.temp_output_path}: {e}")
        
        MonitorFolderDialog._instances.discard(self)
        self.tray_icon.hide()
        self.stop_event.set()
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)
        event.accept()

    def showNormal(self):
        super().showNormal()
        self.raise_()
        self.activateWindow()

# --- Worker Class (Based on v1.4.4 logic) ---
class WhisperWorker(QObject):
    finished_with_path = Signal(str) 
    error = Signal(str)
    def __init__(self, mode, source_file, determined_dest_file_path, model_name="base", device=None):
        super().__init__()
        self.mode=mode; self.source_file=source_file; self.dest_file_path=determined_dest_file_path
        self.device = device or ("cuda" if torch.cuda.is_available() else "cpu")
        if self.device == "cpu":
            logger_worker.warning("CUDA not available, falling back to CPU. This will be slower.")
        else:
            logger_worker.info("CUDA is available, using GPU acceleration.")
        self.model_name=model_name
        self._is_running=True; self.temp_dir=None; self.temp_audio_path=None
        logger_worker.info(f"Worker init. Mode:{self.mode}, Src:'{self.source_file}', Dest:'{self.dest_file_path}', Model:{self.model_name}, Dev:{self.device}")
    def run(self):
        op_ok=False; out_path_sig=""; start_time=time.time()
        logger_worker.info("Worker run method started.")
        try:
            if os.name == 'nt': bring_console_to_front()
            logger_worker.debug("--- Worker Log Segment Start ---")
            logger_worker.info("Starting worker process execution...")
            audio_path = self.source_file
            if not self._is_running: logger_worker.warning("Stop at run start."); return
            if self.mode == "video_transcript":
                logger_worker.info(f"Extracting audio from '{self.source_file}' for transcription.")
                self.temp_dir = tempfile.mkdtemp(prefix="wc_temp_")
                self.temp_audio_path = os.path.join(self.temp_dir, "temp_audio.wav")
                cmd = ["ffmpeg","-i",self.source_file,"-y","-vn","-acodec","pcm_s16le","-ar","16000","-ac","1",self.temp_audio_path]
                logger_worker.info(f"FFMPEG cmd: {' '.join(cmd)}")
                res = subprocess.run(cmd, capture_output=True, text=True, errors='replace', creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
                if res.returncode!=0: logger_worker.error(f"FFMPEG audio extraction error: {res.stderr}"); self.error.emit(f"FFMPEG extract error: {res.stderr[:150]}..."); return
                if not os.path.exists(self.temp_audio_path) or os.path.getsize(self.temp_audio_path)==0:
                    logger_worker.error(f"FFMPEG OK but output missing/empty:'{self.temp_audio_path}'"); self.error.emit("FFMPEG produced no audio data."); return
                logger_worker.info("Audio extraction OK."); audio_path = self.temp_audio_path
            if not self._is_running: logger_worker.warning("Stop after extract."); return
            if self.mode == "rip_audio":
                logger_worker.info(f"Ripping audio to '{self.dest_file_path}'.")
                parent_dir = os.path.dirname(self.dest_file_path); os.makedirs(parent_dir, exist_ok=True)
                cmd = ["ffmpeg","-i",self.source_file,"-y","-vn","-acodec","mp3",self.dest_file_path]
                logger_worker.info(f"FFMPEG cmd: {' '.join(cmd)}")
                res = subprocess.run(cmd, capture_output=True, text=True, errors='replace', creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
                if res.returncode!=0: logger_worker.error(f"FFMPEG rip error: {res.stderr}"); self.error.emit(f"FFMPEG rip error: {res.stderr[:150]}..."); return
                if os.path.exists(self.dest_file_path) and os.path.getsize(self.dest_file_path)>0:
                    logger_worker.info(f"Audio rip OK: '{self.dest_file_path}'"); op_ok=True; out_path_sig=self.dest_file_path
                else: logger_worker.error(f"Audio rip output missing/empty:'{self.dest_file_path}'"); self.error.emit("Audio rip output error.")
            elif self.mode in ["video_transcript", "audio_transcript"]:
                logger_worker.info(f"Loading Whisper model '{self.model_name}' on '{self.device}'.")
                try: self.model = whisper.load_model(self.model_name, device=self.device); logger_worker.info("Model loaded.")
                except Exception as e: logger_worker.error("Model load failed",exc_info=True); self.error.emit(f"Model load error: {e}"); return
                if not self._is_running: logger_worker.warning("Stop after model load."); return
                logger_worker.info(f"Starting transcription... (Audio: '{audio_path}')")
                if not os.path.exists(audio_path): logger_worker.error(f"Audio for transcribe missing:'{audio_path}'"); self.error.emit(f"Audio missing: {os.path.basename(audio_path)}"); return
                if os.path.getsize(audio_path)==0: logger_worker.error(f"Audio for transcribe empty:'{audio_path}'"); self.error.emit(f"Audio empty: {os.path.basename(audio_path)}"); return
                res_dict = None
                try: res_dict = self.model.transcribe(audio_path, verbose=True)
                except Exception as e: logger_worker.critical("transcribe() call failed!",exc_info=True); self.error.emit(f"Transcription error: {e}"); return
                logger_worker.info("transcribe() call completed.")
                if res_dict is None: logger_worker.error("Transcription result is None!"); self.error.emit("Transcription returned no result."); return
                parent_dir = os.path.dirname(self.dest_file_path); os.makedirs(parent_dir, exist_ok=True)
                try:
                    segments = res_dict.get("segments")
                    logger_worker.info(f"Preparing to write transcript to '{self.dest_file_path}'. Segments: {len(segments) if segments else 'N/A'}.")
                    with open(self.dest_file_path, "w", encoding="utf-8") as f:
                        if segments:
                            for i, seg in enumerate(segments):
                                if not self._is_running: logger_worker.warning("Stop during segment write."); break
                                f.write(f"[{format_timestamp_for_transcript(seg['start'])} --> {format_timestamp_for_transcript(seg['end'])}] {seg['text'].strip()}\n")
                                if i%100==0: f.flush() 
                        else: 
                            full_txt = res_dict.get("text","").strip()
                            logger_worker.warning(f"No segments. Writing full text (len {len(full_txt)}).")
                            f.write(full_txt + ("\n" if full_txt else ""))
                    if self._is_running: logger_worker.info(f"Transcript saved: '{self.dest_file_path}'"); op_ok=True; out_path_sig=self.dest_file_path
                    else: logger_worker.info(f"Write interrupted. Partial file at '{self.dest_file_path}'"); out_path_sig=self.dest_file_path 
                except Exception as e: logger_worker.error(f"Transcript write failed for '{self.dest_file_path}'",exc_info=True); self.error.emit(f"Transcript write error: {e}")
            else: logger_worker.error(f"Unknown worker mode: {self.mode}"); self.error.emit(f"Internal error: Unknown mode '{self.mode}'")
        except Exception as e: logger_worker.critical("Unhandled worker exception",exc_info=True); self.error.emit(f"General worker error: {e}")
        finally:
            total_t = time.time()-start_time; logger_worker.info(f"Worker 'finally'. Time:{total_t:.2f}s. OpOK:{op_ok}. Running:{self._is_running}")
            self.cleanup()
            if op_ok and out_path_sig and self._is_running:
                logger_worker.info(f"Emitting finished_with_path: '{out_path_sig}' (Success)")
                self.finished_with_path.emit(out_path_sig)
            elif out_path_sig : 
                logger_worker.warning(f"Emitting finished_with_path: '{out_path_sig}' (OpOK:{op_ok}, Running:{self._is_running}. Error/Stop likely.)")
                self.finished_with_path.emit(out_path_sig) 
            logger_worker.info("Worker run method fully completed.\n--- Worker Log Segment End ---")
    def stop(self): logger_worker.info("Worker stop method called."); self._is_running = False
    def cleanup(self):
        logger_worker.info(f"Cleanup. Temp dir:'{self.temp_dir}'")
        if self.temp_dir and os.path.exists(self.temp_dir):
            try: shutil.rmtree(self.temp_dir); logger_worker.info(f"Cleaned temp:'{self.temp_dir}'")
            except Exception as e: logger_worker.warning(f"Failed to clean temp:'{self.temp_dir}':{e}",exc_info=True)
        else: logger_worker.debug("No temp dir to clean.")
        self.temp_dir=None; self.temp_audio_path=None; logger_worker.info("Cleanup finished.")

# --- Subclass QMenu to force tooltips on hover ---
class PsychoMenu(QMenu):
    def __init__(self, title, parent=None):
        super().__init__(title, parent)
        self.setStyleSheet("background-color: #0d0000; color: white; border: none;")
        
    def event(self, event):
        if event.type() == QEvent.ToolTip:
            action = self.actionAt(event.pos())
            if action and action.toolTip():  # Only show if tooltip is set on action
                QToolTip.showText(event.globalPos(), action.toolTip(), self)
                return True
        return super().event(event)

# --- Main GUI Class (Based on User's Original Structure) ---
class WhisperCreepInterface(QMainWindow):
    def __init__(self):
        super().__init__()
        self.app_start_time = time.time()
        logger_app.info(f"UI initializing. Version: {APP_VERSION}. App Start: {datetime.datetime.fromtimestamp(self.app_start_time).strftime('%Y-%m-%d %H:%M:%S')}")
        self.source_file_path = None
        self.whisper_worker = None
        self.whisper_thread = None
        self.processing_dialog = None
        self.was_killed_by_user = False
        self.current_run_file_handler = None
        self.tray_icon = None
        self.setup_tray_icon()
        self.setWindowTitle("WhisperCreep Control Room")
        self.setMinimumWidth(800)
        self.setMinimumHeight(600)
        
        # Set window background to black
        self.setStyleSheet("""
            QMainWindow {
                background-color: black;
                color: white;
                border: none;
            }
            QStatusBar {
                background-color: black;
                color: white;
                border-top: none;
            }
        """)
        
        # Add a status bar
        status_bar = QStatusBar()
        status_bar.setStyleSheet("background-color: black; color: white; border-top: none;")
        self.setStatusBar(status_bar)
        
        menu_bar = self.menuBar()
        menu_bar.setStyleSheet("QMenuBar { background-color: #0d0000; color: white; } QMenuBar::item:selected { background: #333333; } QMenu { background-color: #0d0000; color: white; } QMenu::item:selected { background-color: #333333; }")
        file_menu = menu_bar.addMenu("File")
        
        # Add Change App Icon option to File menu
        change_icon_action = file_menu.addAction("Change App Icon")
        change_icon_action.setToolTip("Choose a custom icon for the application")
        change_icon_action.triggered.connect(self.change_app_icon)
        
        exit_action = file_menu.addAction("Exit")
        exit_action.triggered.connect(self.close)
        
        # Use PsychoMenu for Tools menu to force tooltips
        tools_menu = PsychoMenu("Tools", self)
        tools_menu.setStyleSheet("QMenu { background-color: #0d0000; color: white; } QMenu::item:selected { background-color: #333333; }")
        menu_bar.addMenu(tools_menu)
        
        help_menu = menu_bar.addMenu("Help")
        central_widget = QWidget()
        central_widget.setStyleSheet("background-color: black; color: white;")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        self.button_group = QButtonGroup(self)
        self.button_group.buttonClicked.connect(self.enable_file_buttons)
        def create_option_row(description, help_button, radio_button):
            layout = QHBoxLayout()
            label = QLabel(description)
            label.setStyleSheet("font-family:Calibri;font-size:11pt;")
            layout.addWidget(label)
            layout.addStretch()
            radio_button.setStyleSheet("padding-left:8px;padding-right:8px;")
            layout.addWidget(radio_button)
            layout.addSpacing(20)
            layout.addWidget(help_button)
            return layout
        ops_data = [("video_transcript", "Extract Video Transcript", "Extract spoken content from video.", ".mp4..."),
                    ("rip_audio", "Rip Audio from Video", "Extract audio track (MP3) from video.", ".mp4..."),
                    ("audio_transcript", "Extract Audio Transcript", "Transcribe an audio-only file.", ".mp3...")]
        gb_style = "QGroupBox{background-color:black;color:white;border:1px solid #222;padding:10px;margin-top:1ex;}QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;padding:0 10px;font-weight:bold;}"
        for n, t, d, ft in ops_data:
            rb = QRadioButton()
            rb.setObjectName(n)
            self.button_group.addButton(rb)
            hb = QPushButton("?")
            hb.setFixedWidth(25)
            hb.setStyleSheet("background-color:royalblue;color:white;")
            hb.clicked.connect(lambda c=False, ty=ft, tt=t: QMessageBox.information(self, f"{tt} - Types", ty))
            gb = QGroupBox(t)
            gb.setStyleSheet(gb_style)
            gb.setLayout(create_option_row(d, hb, rb))
            main_layout.addWidget(gb)
        src_gb = QGroupBox("Select Source File")
        src_gb.setStyleSheet(gb_style)
        src_lo = QHBoxLayout()
        self.src_btn = QPushButton("Browse Source File")
        self.src_btn.setStyleSheet("background-color:royalblue;color:white;")
        self.src_btn.setEnabled(False)
        self.src_btn.clicked.connect(self.browse_source_file)
        src_lo.addWidget(self.src_btn)
        self.src_file_lbl = QLabel("No file selected.")
        src_lo.addWidget(self.src_file_lbl)
        src_lo.addStretch()
        src_gb.setLayout(src_lo)
        main_layout.addWidget(src_gb)
        dst_gb = QGroupBox("Output Location (Automatic)")
        dst_gb.setStyleSheet(gb_style)
        dst_lo = QHBoxLayout()
        self.dst_info_lbl = QLabel("Output will be saved to your Downloads folder.")
        dst_lo.addWidget(self.dst_info_lbl)
        dst_gb.setLayout(dst_lo)
        main_layout.addWidget(dst_gb)
        mdl_gb = QGroupBox("Select Whisper Model")
        mdl_gb.setStyleSheet(gb_style)
        mdl_lo = QHBoxLayout()
        self.mdl_combo = QComboBox()
        self.mdl_combo.addItems(["tiny", "base", "small", "medium", "large"])
        self.mdl_combo.setCurrentText("base")
        mdl_lo.addWidget(QLabel("Model Size:"))
        mdl_lo.addWidget(self.mdl_combo)
        mdl_gb.setLayout(mdl_lo)
        main_layout.addWidget(mdl_gb)
        btm_lo = QHBoxLayout()
        self.run_btn = QPushButton("Run")
        self.run_btn.setFixedWidth(200)
        self.run_btn.setStyleSheet("background-color:purple;color:white;padding:5px;")
        self.run_btn.clicked.connect(self.start_processing)
        self.run_btn.setEnabled(False)
        btm_lo.addWidget(self.run_btn)
        self.rst_btn = QPushButton("Start Over")
        self.rst_btn.setStyleSheet("background-color:gray;color:black;padding:5px;")
        self.rst_btn.clicked.connect(self.reset_form)
        btm_lo.addWidget(self.rst_btn)
        btm_lo.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        min_btn = QPushButton("Minimize")
        min_btn.setStyleSheet("background-color:royalblue;color:white;padding:5px;")
        min_btn.clicked.connect(self.minimize_to_tray)
        btm_lo.addWidget(min_btn)
        cls_btn = QPushButton("Close")
        cls_btn.setStyleSheet("background-color:crimson;color:white;padding:5px;")
        cls_btn.clicked.connect(self.close)
        btm_lo.addWidget(cls_btn)
        main_layout.addLayout(btm_lo)
        logger_app.info("UI widgets created (original structure).")

        # New menu option
        monitor_action = tools_menu.addAction("Monitor Folder for Video")
        monitor_action.setToolTip("Monitor a folder and auto transcribe video files.")
        monitor_action.triggered.connect(self.open_monitor_dialog)
        
        # Add Frame Snatcher to tools menu
        frame_snatcher_action = tools_menu.addAction("Frame Snatcher")
        frame_snatcher_action.setToolTip("Choose to create still frames from a video for analysis")
        frame_snatcher_action.triggered.connect(self.open_frame_snatcher)

        # Add YouTube Caption Fetcher to tools menu
        youtube_caption_action = tools_menu.addAction("YouTube Caption Fetcher")
        youtube_caption_action.setToolTip("Download and clean captions from YouTube videos")
        youtube_caption_action.triggered.connect(self.open_youtube_caption_fetcher)

        # Add Website to PDF to tools menu
        webtopdf_action = tools_menu.addAction("WebPage to PDF")
        webtopdf_action.setToolTip("Convert websites to PDF documents")
        webtopdf_action.triggered.connect(lambda: os.system(f'python "{os.path.join(os.path.dirname(os.path.abspath(__file__)), "webtopdf_gui.py")}"'))

    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        # Try to load the custom icon from config first
        icon_loaded = False
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(script_dir, "config.ini")
            
            if os.path.exists(config_path):
                config = configparser.ConfigParser()
                config.read(config_path)
                
                if config.has_section('UI') and config.has_option('UI', 'icon_path'):
                    icon_path = os.path.join(script_dir, config.get('UI', 'icon_path'))
                    logger_app.info(f"Attempting to load custom icon from config: {icon_path}")
                    
                    if os.path.exists(icon_path):
                        icon = QIcon(icon_path)
                        if not icon.isNull():
                            self.tray_icon.setIcon(icon)
                            self.setWindowIcon(icon)
                            app = QApplication.instance()
                            app.setWindowIcon(icon)
                            logger_app.info(f"Successfully loaded custom icon from: {icon_path}")
                            icon_loaded = True
                        else:
                            raise Exception("Custom icon loaded but is null")
                    else:
                        logger_app.warning(f"Custom icon file not found at: {icon_path}")
        except Exception as e:
            logger_app.error(f"Error loading custom icon from config: {e}")
        
        # Fall back to default icon if custom wasn't loaded
        if not icon_loaded:
            try:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                icon_path = os.path.join(script_dir, "WhisperCreepICO.ico")
                logger_app.info(f"Attempting to load default icon from: {icon_path}")
                
                if os.path.exists(icon_path):
                    icon = QIcon(icon_path)
                    if not icon.isNull():
                        self.tray_icon.setIcon(icon)
                        self.setWindowIcon(icon)
                        app = QApplication.instance()
                        app.setWindowIcon(icon)
                        logger_app.info(f"Successfully loaded default icon from: {icon_path}")
                    else:
                        raise Exception("Icon loaded but is null")
                else:
                    raise FileNotFoundError(f"Icon file not found at: {icon_path}")
            except Exception as e:
                logger_app.error(f"Error loading default icon: {e}")
                self.tray_icon.setIcon(QIcon.fromTheme("system-run", QIcon.fromTheme("applications-system")))
                logger_app.info("Using system default icon for main window")
        
        self.tray_icon.setToolTip("WhisperCreep Main Window")
        menu = QMenu()
        restore_action = QAction("Restore Main Window", self)
        restore_action.triggered.connect(self.showNormal)
        menu.addAction(restore_action)
        quit_action = QAction("Quit Application", self)
        quit_action.triggered.connect(self.close)
        menu.addAction(quit_action)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self.on_tray_activated)

    def minimize_to_tray(self):
        if not self.tray_icon.isVisible():
            self.tray_icon.show()
        self.hide()
        self.tray_icon.setToolTip("WhisperCreep Main Window (Minimized)")

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.showNormal()
            self.raise_()
            self.activateWindow()
            # Check if monitoring dialog exists and restore it too
            for widget in QApplication.topLevelWidgets():
                if isinstance(widget, MonitorFolderDialog) and widget.monitoring:
                    widget.showNormal()
                    widget.raise_()
                    widget.activateWindow()

    def open_monitor_dialog(self):
        global transcription_in_progress
        if transcription_in_progress:
            QMessageBox.warning(self, "Transcription in Progress", "Cannot start folder monitoring while a transcription is in progress.")
            return
        dialog = MonitorFolderDialog(self)
        dialog.exec()
        
    def open_frame_snatcher(self):
        """Opens the Frame Snatcher tool in a new window"""
        logger_app.info("Opening Frame Snatcher tool")
        self.frame_snatcher = VideoFrameSnatcher()
        self.frame_snatcher.show()

    def open_youtube_caption_fetcher(self):
        """Opens the YouTube Caption Fetcher tool in a new window"""
        logger_app.info("Opening YouTube Caption Fetcher tool")
        self.youtube_caption_fetcher = YouTubeCaptionFetcher()
        self.youtube_caption_fetcher.show()

    def start_processing(self):
        global transcription_in_progress
        if transcription_in_progress:
            QMessageBox.warning(self, "Transcription in Progress", "Cannot start a new transcription while another is in progress.")
            return
        transcription_in_progress = True
        self.tray_icon.setToolTip("WhisperCreep Control Room - Transcription in Progress")
        run_start_time = datetime.datetime.now()
        # Close any previous run's specific log handler FIRST
        self._close_run_specific_logging() 

        logger_app.info(f"--- User clicked 'Run' at {run_start_time.strftime('%Y-%m-%d %H:%M:%S')} ---")
        self.was_killed_by_user = False 
        
        sel_op_btn = self.button_group.checkedButton()
        if not (sel_op_btn and self.source_file_path):
            QMessageBox.warning(self,"Missing Info","Select operation & source file."); logger_app.warning("Run: op/src missing."); return

        prod_path, log_path = self._generate_output_paths()
        if not (prod_path and log_path):
            logger_app.error("Failed to determine output paths. Aborting."); return
        
        if not self._setup_run_specific_logging(log_path):
            logger_app.error("Failed to setup run-specific logging. Aborting."); return 

        mode = sel_op_btn.objectName(); mdl_name = self.mdl_combo.currentText()
        logger_app.info(f"Run Details (also logged to '{os.path.basename(log_path)}'):")
        logger_app.info(f"  Mode: {mode}, Model: {mdl_name}")
        logger_app.info(f"  Source: {self.source_file_path}")
        logger_app.info(f"  Production Output: {prod_path}")
        logger_app.info(f"  Debug Log For This Run: {log_path}")
        
        self._update_gui_for_processing_state(True) 

        self.whisper_thread = QThread(self) 
        self.whisper_worker = WhisperWorker(mode, self.source_file_path, prod_path, mdl_name)
        self.whisper_worker.moveToThread(self.whisper_thread)
        self.whisper_worker.error.connect(self.handle_worker_error) 
        self.whisper_worker.finished_with_path.connect(self.handle_worker_file_saved_or_issue) 
        self.whisper_thread.started.connect(self.whisper_worker.run)
        self.whisper_worker.finished_with_path.connect(self.whisper_thread.quit)
        self.whisper_worker.error.connect(self.whisper_thread.quit)   
        self.whisper_thread.finished.connect(self.whisper_worker.deleteLater)
        self.whisper_thread.finished.connect(self.whisper_thread.deleteLater)
        self.whisper_thread.finished.connect(self._ensure_gui_finalized_on_thread_end) 
        self.whisper_thread.start(); logger_app.info("Worker thread started.")

    def _ensure_gui_finalized_on_thread_end(self):
        logger_app.debug(f"QThread 'finished' signal. run_btn text: '{self.run_btn.text()}'")
        if self.run_btn.text() == "Processing...": 
            logger_app.warning("Thread finished, but GUI finalize may not have run. Forcing GUI state reset.")
            self._finalize_gui_after_processing(False, "Thread ended; outcome uncertain.", self.was_killed_by_user)

    def handle_worker_error(self, err_msg): 
        logger_app.error(f"Worker error: {err_msg}")
        if self.processing_dialog: self.processing_dialog.stop_animation_and_close(); self.processing_dialog = None
        QMessageBox.critical(self, "Error", err_msg) 
        self._finalize_gui_after_processing(False, f"Error: {err_msg}", self.was_killed_by_user)

    def handle_worker_file_saved_or_issue(self, out_path_msg):
        logger_app.info(f"Worker finished_with_path. Path/Msg: '{out_path_msg}'")
        if self.processing_dialog: self.processing_dialog.stop_animation_and_close(); self.processing_dialog = None

        if self.was_killed_by_user:
            msg = "Process stopped by user."
            file_info = f"Partial output (if any):\n{out_path_msg}" if out_path_msg and os.path.exists(out_path_msg) else "No output file path available or file not created."
            logger_app.warning(f"{msg} {file_info.replace(os.linesep, ' ')}") # Log as single line
            QMessageBox.warning(self, "Process Terminated", f"{msg}\n{file_info}")
            self._finalize_gui_after_processing(False, f"Killed. File: {out_path_msg}", True)
            return 

        if os.path.exists(out_path_msg): 
            QMessageBox.information(self, "Success", f"Operation complete!\nOutput:\n{out_path_msg}")
            logger_app.info(f"SUCCESS. Output: {out_path_msg}")
            self._finalize_gui_after_processing(True, out_path_msg, False)
        else:
            QMessageBox.warning(self, "Process Note", f"Worker finished.\nPath: {out_path_msg}\nFile not created as expected or op not fully successful. Check log.")
            logger_app.warning(f"Worker finished, path '{out_path_msg}' not existing. Check worker logs.")
            self._finalize_gui_after_processing(False, f"File issue: {out_path_msg}", False)

    def _finalize_gui_after_processing(self, success, outcome_message_or_path, killed=False):
        global transcription_in_progress
        transcription_in_progress = False
        self.tray_icon.setToolTip("WhisperCreep Main Window")
        logger_app.info(f"Finalizing GUI. Success:{success}, Killed:{killed}, Path/Msg:'{outcome_message_or_path}'")
        if self.processing_dialog and self.processing_dialog.isVisible(): 
            logger_app.warning("Finalizing GUI, processing dialog still visible. Closing."); self.processing_dialog.stop_animation_and_close(); self.processing_dialog = None
        
        self._update_gui_for_processing_state(False) 
        self._close_run_specific_logging() # Close the run-specific log file
        
        run_end_time = datetime.datetime.now()
        outcome = "Killed by user" if killed else ("Success" if success else "Failure/Error")
        logger_app.info(f"--- GUI 'Run' handling finished at {run_end_time.strftime('%Y-%m-%d %H:%M:%S')}. Outcome: {outcome}, Details: {outcome_message_or_path} ---")

        # Find any paused monitoring dialog and ask user about resuming
        for widget in QApplication.topLevelWidgets():
            if isinstance(widget, MonitorFolderDialog) and widget.monitoring:
                logger_app.info("Manual transcription completed, asking user about monitoring...")
                widget.log_event(f"Manual transcription completed. Outcome: {outcome}. Asking user about monitoring.")
                reply = QMessageBox.question(self, 'Resume Monitoring?',
                    'Your transcription is complete. Would you like to resume folder monitoring?',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes)
                
                if reply == QMessageBox.StandardButton.Yes:
                    logger_app.info("User chose to resume monitoring")
                    widget.log_event("User chose to resume monitoring")
                    # Minimize both windows to tray
                    self.minimize_to_tray()
                    widget.hide()
                else:
                    logger_app.info("User chose to stop monitoring")
                    widget.log_event("User chose to stop monitoring after manual transcription")
                    widget.stop_monitoring()

    def closeEvent(self, event): 
        self.tray_icon.hide()
        shutdown_time = datetime.datetime.now()
        exit_reason = "User killed process" if self.was_killed_by_user else "Normal window close"
        logger_app.info(f"Close event at {shutdown_time.strftime('%Y-%m-%d %H:%M:%S')}. Reason: {exit_reason}. App shutting down.")
        
        if self.processing_dialog: logger_app.debug("Closing active processing dialog."); self.processing_dialog.stop_animation_and_close(); self.processing_dialog = None
        if hasattr(self,'whisper_thread') and self.whisper_thread and self.whisper_thread.isRunning():
            logger_app.info("Stopping worker thread...")
            if hasattr(self,'whisper_worker') and self.whisper_worker: self.whisper_worker.stop()
            self.whisper_thread.quit()
            if not self.whisper_thread.wait(2000): logger_app.warning("Worker didn't stop gracefully. Terminating."); self.whisper_thread.terminate(); self.whisper_thread.wait()
        
        self._close_run_specific_logging() 
        super().closeEvent(event)
        total_uptime = time.time() - self.app_start_time
        logger_app.info(f"--- Application Closed. Final Exit: {exit_reason}. Uptime: {total_uptime:.2f}s. ---")

    def _setup_run_specific_logging(self, log_file_path):
        if not log_file_path: 
            logger_app.error("Attempted to set up run-specific logging with no log_file_path.")
            return False
        try:
            if self.current_run_file_handler:
                logger_app.debug(f"Removing previous run file handler for: {self.current_run_file_handler.baseFilename}")
                logging.getLogger().removeHandler(self.current_run_file_handler)
                self.current_run_file_handler.close()
            
            formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
            file_handler = logging.FileHandler(log_file_path, mode='w', encoding='utf-8') 
            file_handler.setFormatter(formatter)
            file_handler.setLevel(logging.DEBUG) 
            
            logging.getLogger().addHandler(file_handler) 
            self.current_run_file_handler = file_handler
            logger_app.info(f"--- Run-specific file logging initiated to: {log_file_path} ---")
            return True
        except Exception as e:
            logger_app.error(f"Failed to set up run-specific file logging for '{log_file_path}': {e}", exc_info=True)
            QMessageBox.critical(self, "Logging Error", f"Could not create run-specific log file at:\n{log_file_path}\n\nLogging will continue to console only.\nError: {e}")
            self.current_run_file_handler = None 
            return False

    def _close_run_specific_logging(self):
        if self.current_run_file_handler:
            log_path = self.current_run_file_handler.baseFilename
            logger_app.info(f"--- Closing run-specific logging for: {log_path} ---")
            try:
                logging.getLogger().removeHandler(self.current_run_file_handler)
                self.current_run_file_handler.close()
            except Exception as e: logger_app.error(f"Error closing run-specific log for '{log_path}': {e}", exc_info=True)
            finally: self.current_run_file_handler = None
        else: logger_app.debug("No active run-specific log handler to close.")

    def _update_gui_for_processing_state(self, is_processing):
        logger_app.debug(f"Updating GUI for processing state: {is_processing}")
        can_run_now = (self.source_file_path is not None) and (self.button_group.checkedButton() is not None)
        self.run_btn.setEnabled(not is_processing and can_run_now)
        self.run_btn.setText("Processing..." if is_processing else "Run")
        self.rst_btn.setEnabled(not is_processing)
        self.src_btn.setEnabled(not is_processing and (self.button_group.checkedButton() is not None))
        for btn in self.button_group.buttons(): btn.setEnabled(not is_processing)
        self.mdl_combo.setEnabled(not is_processing)
        if is_processing:
            if not self.processing_dialog: self.processing_dialog = ProcessingIndicatorDialog(self); self.processing_dialog.kill_process_requested.connect(self._confirm_kill_process)
            if not self.processing_dialog.isVisible(): self.processing_dialog.start_animation()
        else: 
            if self.processing_dialog: self.processing_dialog.stop_animation_and_close(); self.processing_dialog = None 

    def _update_run_button_state(self): self.run_btn.setEnabled(bool(self.button_group.checkedButton() and self.source_file_path))
    def reset_form(self): 
        logger_app.info("Resetting form.")
        self.source_file_path = None
        self.was_killed_by_user = False
        
        # Uncheck the radio button
        cc = self.button_group.checkedButton()
        if cc:
            self.button_group.setExclusive(False)
            cc.setChecked(False)
            self.button_group.setExclusive(True)
        
        # Reset UI elements
        self.src_btn.setEnabled(False)
        self.src_btn.setText("Browse Source File")
        self.src_file_lbl.setText("No file selected.")
        self.mdl_combo.setCurrentText("base")
        self._update_gui_for_processing_state(False)
        self._update_run_button_state()
        
        # Resume any paused monitoring
        for widget in QApplication.topLevelWidgets():
            if isinstance(widget, MonitorFolderDialog) and widget.monitoring:
                logger_app.info("Form cleared, resuming monitoring...")
                widget.log_event("Monitoring resumed after form clear")
                widget.tray_icon.setToolTip("WhisperCreep Monitor - Active")
        
        logger_app.info("Form reset OK.")
    def enable_file_buttons(self): 
        logger_app.debug("Op selected.")
        self.src_btn.setEnabled(True)
        self._update_run_button_state()
        
        # Check if monitoring is active and show warning if transcription is selected
        for widget in QApplication.topLevelWidgets():
            if isinstance(widget, MonitorFolderDialog) and widget.monitoring:
                sel_op_btn = self.button_group.checkedButton()
                if sel_op_btn and sel_op_btn.objectName() in ["video_transcript", "audio_transcript"]:
                    QMessageBox.information(self, "Monitoring Active",
                        "Note: Starting a transcription will pause the folder monitoring.\n\n"
                        "The monitoring will automatically resume after your transcription is complete, "
                        "plus the scheduled monitoring interval.\n\n"
                        "You can continue with your transcription or stop the monitoring first.")
                    return
    def browse_source_file(self): 
        logger_app.debug("Browse source."); sel_op_btn=self.button_group.checkedButton(); filt="All (*)"
        if sel_op_btn:
            nm=sel_op_btn.objectName()
            if nm=="video_transcript" or nm=="rip_audio": filt="Video (*.mp4 *.mov *.avi *.wmv *.mkv);;All (*)"
            elif nm=="audio_transcript": filt="Audio (*.mp3 *.wav *.m4a);;All (*)"
        f_path_tuple=QFileDialog.getOpenFileName(self,"Select Source",os.path.expanduser("~"),filt); f_path=f_path_tuple[0]
        if f_path:
            self.source_file_path=f_path; btn_txt=f"Src: {os.path.basename(f_path)}"
            self.src_btn.setText(btn_txt if len(btn_txt)<35 else f"Src: ...{os.path.basename(f_path)[-30:]}")
            self.src_file_lbl.setText(f"{os.path.basename(f_path)}"); logger_app.info(f"Src file: {f_path}")
        else:
            if not self.source_file_path: self.src_btn.setText("Browse Source File"); self.src_file_lbl.setText("No file selected.")
            logger_app.debug("Src selection cancelled.")
        self._update_run_button_state()

    def _generate_output_paths(self):
        sel_op_btn = self.button_group.checkedButton()
        if not (sel_op_btn and self.source_file_path): 
            logger_app.error("Cannot generate output paths: operation or source file missing.")
            return None, None
        op_nm = sel_op_btn.objectName()
        src_base_no_ext = os.path.splitext(os.path.basename(self.source_file_path))[0]
        
        dloads_dir = get_download_folder_path() # This now has more robust error handling
        if not dloads_dir: 
            # get_download_folder_path will log critical errors.
            # A QMessageBox is shown here to inform the user directly if path generation fails at this stage.
            QMessageBox.critical(self, "Directory Error", "Could not determine or create the Downloads folder. Please check logs and permissions. Cannot proceed.")
            logger_app.critical("Failed to get a valid Downloads directory path. Output paths cannot be generated.")
            return None, None 
            
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        f_stem = src_base_no_ext.replace(" ","_") # Make filename a bit cleaner
        prod_ext = ""; log_suffix = "_debug.log"

        if op_nm=="video_transcript": f_stem+=f"_transcript_{ts}"; prod_ext=".txt"
        elif op_nm=="rip_audio": f_stem+=f"_rip_{ts}"; prod_ext=".mp3"
        elif op_nm=="audio_transcript": f_stem+=f"_transcript_{ts}"; prod_ext=".txt"
        else: 
            logger_app.error(f"Unknown operation name '{op_nm}' for generating output paths.")
            QMessageBox.critical(self, "Internal Error", f"Unknown operation '{op_nm}' selected. Cannot determine output filenames.")
            return None, None
        
        prod_path = os.path.join(dloads_dir, f_stem + prod_ext)
        log_path = os.path.join(dloads_dir, f_stem + log_suffix)
        
        logger_app.info(f"Generated production output path: {prod_path}")
        logger_app.info(f"Generated debug log path for this run: {log_path}")
        return prod_path, log_path

    def _confirm_kill_process(self): 
        logger_app.info("User initiated kill from dialog.")
        reply = QMessageBox.question(self,'Confirm Kill','Are you sure? This will stop processing and close app.\nOutput may be incomplete.', QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            logger_app.warning("User confirmed KILL. Closing app."); self.was_killed_by_user = True
            if self.processing_dialog: self.processing_dialog.stop_animation_and_close(); self.processing_dialog = None
            self.close() 
        else: logger_app.info("User cancelled kill.")

    def change_app_icon(self):
        """
        Allows user to select a custom icon for the application.
        Copies the selected icon to the application directory and updates config.ini.
        """
        logger_app.info("Opening dialog to change app icon")
        
        # Get icon file from user
        icon_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Application Icon", 
            os.path.expanduser("~"),
            "Icon Files (*.ico *.png)"
        )
        
        if not icon_path:
            logger_app.info("Icon selection cancelled")
            return
            
        try:
            # Create app config directory if it doesn't exist
            script_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(script_dir, "config.ini")
            
            # Copy the selected icon to our directory with a standard name
            icon_filename = "custom_app_icon" + os.path.splitext(icon_path)[1]
            destination_path = os.path.join(script_dir, icon_filename)
            
            # Copy the icon file
            shutil.copy2(icon_path, destination_path)
            logger_app.info(f"Copied icon from {icon_path} to {destination_path}")
            
            # Update config file
            config = configparser.ConfigParser()
            
            # Load existing config if it exists
            if os.path.exists(config_path):
                config.read(config_path)
                
            # Ensure we have the UI section
            if not config.has_section('UI'):
                config.add_section('UI')
                
            # Set the icon path relative to the script directory
            config.set('UI', 'icon_path', icon_filename)
            
            # Save the config
            with open(config_path, 'w') as config_file:
                config.write(config_file)
                
            logger_app.info(f"Updated config file at {config_path}")
            
            # Try to apply the icon immediately
            icon = QIcon(destination_path)
            if not icon.isNull():
                self.setWindowIcon(icon)
                app = QApplication.instance()
                app.setWindowIcon(icon)
                
                # Update tray icon too
                self.tray_icon.setIcon(icon)
                
                logger_app.info("Successfully applied new icon to application")
                QMessageBox.information(
                    self, 
                    "Icon Changed", 
                    "The application icon has been changed.\nThe new icon will be used the next time you start the application."
                )
            else:
                raise Exception("Icon loaded but is null")
                
        except Exception as e:
            logger_app.error(f"Error changing app icon: {e}", exc_info=True)
            QMessageBox.critical(
                self, 
                "Icon Change Failed", 
                f"Could not change the application icon.\nError: {str(e)}"
            )

# --- Main Application Execution ---
if __name__ == "__main__":
    logger_app.info(f"--- Application Starting Up (Console Logging Only Initially). Version: {APP_VERSION}, PID: {os.getpid()} ---")
    if os.name == 'nt': bring_console_to_front() 

    # Tell Qt to use the system's dark theme if available
    app = QApplication(sys.argv)
    app.setAttribute(Qt.ApplicationAttribute.AA_UseStyleSheetPropagationInWidgetStyles, True)
    
    # Set up Windows dark title bar if possible
    if os.name == 'nt':
        try:
            # Use ctypes to access Windows API
            import ctypes
            
            # Try to enable dark mode for title bar
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            hwnd = ctypes.windll.user32.GetActiveWindow()
            
            # Enable dark mode for the app
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(ctypes.c_int(1)),
                ctypes.sizeof(ctypes.c_int)
            )
            logger_app.info("Windows dark mode for title bar enabled")
        except Exception as e:
            logger_app.warning(f"Could not enable Windows dark mode for title bar: {e}")
    
    window = WhisperCreepInterface()
    
    # Apply dark mode to window after creation
    if os.name == 'nt':
        try:
            # Apply dark mode to title bar and window frame
            hwnd = window.winId()
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                int(hwnd), 
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(ctypes.c_int(1)),
                ctypes.sizeof(ctypes.c_int)
            )
            
            # Additional Windows-specific style adjustments to ensure full dark mode
            # Use GetSystemParametersInfo to get system colors to match system theme perfectly
            window.setStyleSheet(window.styleSheet() + """
                QMainWindow {
                    border: none;
                }
            """)
            
            # Force a style update
            window.setStyle(window.style())
        except Exception as e:
            logger_app.warning(f"Could not apply dark mode to window: {e}")
    
    window.show()
    logger_app.info("App window shown. Entering main event loop.")
    sys.exit(app.exec())


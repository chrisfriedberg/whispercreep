import os
import subprocess
import logging
from datetime import datetime
from pathlib import Path
import webbrowser

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QProgressBar
)
from PySide6.QtCore import QThread, Signal

class YouTubeCaptionWorker(QThread):
    log_signal = Signal(str)
    error_signal = Signal(str)
    done_signal = Signal(str)

    def __init__(self, video_url: str, output_dir: Path):
        super().__init__()
        self.video_url = video_url
        self.output_dir = output_dir

    def parse_timestamp(self, ts):
        """Parse SRT timestamp to seconds."""
        h, m, s_ms = ts.split(':')
        s, ms = s_ms.split(',')
        seconds = int(h) * 3600 + int(m) * 60 + int(s) + int(ms) / 1000
        return seconds

    def clean_srt_file(self, srt_path, output_path):
        try:
            lines = srt_path.read_text(encoding="utf-8").splitlines()
            cleaned = []
            seen_text = set()
            buffer = []

            for line in lines:
                if line.strip() == "":
                    if len(buffer) >= 3:
                        timestamp_line = buffer[1]
                        start_str, end_str = timestamp_line.split(' --> ')
                        start_time = self.parse_timestamp(start_str.strip())
                        end_time = self.parse_timestamp(end_str.strip())
                        duration = end_time - start_time
                        if duration >= 0.5:
                            actual_text_content = "\n".join(text_line.strip() for text_line in buffer[2:]).strip()
                            if actual_text_content and actual_text_content not in seen_text:
                                cleaned.extend(buffer)
                                cleaned.append("")
                                seen_text.add(actual_text_content)
                    buffer = []
                else:
                    buffer.append(line)

            # Process the last buffer if necessary
            if len(buffer) >= 3:
                timestamp_line = buffer[1]
                start_str, end_str = timestamp_line.split(' --> ')
                start_time = self.parse_timestamp(start_str.strip())
                end_time = self.parse_timestamp(end_str.strip())
                duration = end_time - start_time
                if duration >= 0.5:
                    actual_text_content = "\n".join(text_line.strip() for text_line in buffer[2:]).strip()
                    if actual_text_content and actual_text_content not in seen_text:
                        cleaned.extend(buffer)
                        cleaned.append("")

            if cleaned:
                output_path.write_text("\n".join(cleaned), encoding="utf-8")
                if srt_path.exists() and srt_path != output_path:
                    srt_path.unlink()
            else:
                self.log_signal.emit(f"Warning: Cleaning resulted in an empty caption file from {srt_path.name}. Original not deleted.")

        except Exception as e:
            self.error_signal.emit(f"File cleaning/processing failed: {str(e)}\nOriginal SRT was: {srt_path.name if srt_path else 'Unknown'}")

    def run(self):
        try:
            import sys 
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            final_cleaned_output_path = self.output_dir / f"YoutubeSubs_{timestamp}.srt"

            yt_dlp_path_str = os.path.join(sys.exec_prefix, "Scripts", "yt-dlp.exe")
            if not os.path.exists(yt_dlp_path_str):
                yt_dlp_path_str = os.path.join(os.environ["USERPROFILE"], "AppData", "Roaming", "Python", "Python312", "Scripts", "yt-dlp.exe")

            if not Path(yt_dlp_path_str).exists():
                self.error_signal.emit(f"yt-dlp.exe not found at expected locations:\n"
                                       f"{os.path.join(sys.exec_prefix, 'Scripts', 'yt-dlp.exe')}\n"
                                       f"{os.path.join(os.environ['USERPROFILE'], 'AppData', 'Roaming', 'Python', 'Python312', 'Scripts', 'yt-dlp.exe')}\n"
                                       f"Please ensure yt-dlp is installed.")
                return

            yt_dlp_output_template = self.output_dir / "%(title)s.%(ext)s"

            cmd = [
                yt_dlp_path_str,
                "--write-auto-sub",
                "--sub-lang", "en",
                "--skip-download",
                "--convert-subs", "srt",
                "-o", str(yt_dlp_output_template),
                self.video_url
            ]

            self.log_signal.emit(f"Running command: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True, errors='replace')
            self.log_signal.emit(f"STDOUT:\n{result.stdout}")
            self.log_signal.emit(f"STDERR:\n{result.stderr}")

            if result.returncode != 0:
                potential_srt_files = sorted(
                    [f for f in self.output_dir.glob("*.en.srt") if f.is_file()],
                    key=lambda f: f.stat().st_mtime,
                    reverse=True
                )
                if not potential_srt_files:
                    self.error_signal.emit(f"yt-dlp failed and no SRT file found. STDERR:\n{result.stderr}")
                    return
                downloaded_srt_file_path = potential_srt_files[0]
                self.log_signal.emit(f"yt-dlp exited with code {result.returncode} but an SRT file '{downloaded_srt_file_path.name}' was found. Attempting to process it.")
            else:
                downloaded_srt_files = sorted(
                    [f for f in self.output_dir.glob("*.en.srt") if f.is_file()],
                    key=lambda f: f.stat().st_mtime,
                    reverse=True
                )
                if not downloaded_srt_files:
                    self.error_signal.emit("SRT file not found after successful yt-dlp download.")
                    return
                downloaded_srt_file_path = downloaded_srt_files[0]

            self.log_signal.emit(f"Processing downloaded SRT file: {downloaded_srt_file_path}")
            self.clean_srt_file(downloaded_srt_file_path, final_cleaned_output_path)
            if final_cleaned_output_path.exists():
                self.done_signal.emit(str(final_cleaned_output_path))

        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.log_signal.emit(f"Exception Traceback:\n{tb}")
            self.error_signal.emit(f"An unexpected error occurred in the worker thread: {str(e)}")

class YouTubeCaptionFetcher(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("YouTube Caption Downloader")
        self.setMinimumHeight(220)
        self.setMinimumWidth(520)
        self.layout = QVBoxLayout(self)

        self.url_label = QLabel("YouTube Video URL:")
        self.url_input = QLineEdit()

        self.download_button = QPushButton("Download Captions")
        self.cancel_button = QPushButton("Cancel")

        self.download_button.setStyleSheet("QPushButton { background-color: green; color: white; font-weight: bold; }")
        self.cancel_button.setStyleSheet("QPushButton { background-color: red; color: white; font-weight: bold; }")

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setVisible(False)

        self.layout.addWidget(self.url_label)
        self.layout.addWidget(self.url_input)
        self.layout.addWidget(self.download_button)
        self.layout.addWidget(self.cancel_button)
        self.layout.addWidget(self.progress_bar)

        self.download_button.clicked.connect(self.handle_download)
        self.cancel_button.clicked.connect(self.handle_cancel)

        self.output_dir = Path(os.path.join(os.environ["USERPROFILE"], "Downloads"))
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.worker = None

    def handle_download(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "Input Error", "Please enter a YouTube URL.")
            return

        self.progress_bar.setVisible(True)
        self.download_button.setEnabled(False)
        self.cancel_button.setEnabled(False) 

        self.worker = YouTubeCaptionWorker(url, self.output_dir)
        self.worker.log_signal.connect(self.log)
        self.worker.error_signal.connect(self.show_error)
        self.worker.done_signal.connect(self.show_success)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.start()

    def handle_cancel(self):
        if self.worker and self.worker.isRunning():
            self.log("Cancel button clicked. Worker will be left to finish or error out; closing UI.")
        self.close()
        
    def on_worker_finished(self):
        self.progress_bar.setVisible(False)
        self.download_button.setEnabled(True)
        self.cancel_button.setEnabled(True)
        self.log("Worker thread finished signal received.")

    def log(self, message):
        print(message) 
        logging.info(message)

    def show_error(self, message):
        self.progress_bar.setVisible(False)
        self.download_button.setEnabled(True)
        self.cancel_button.setEnabled(True)
        QMessageBox.critical(self, "Download Failed", message)
        logging.error(f"Download/Processing Failed: {message}")

    def show_success(self, filepath):
        self.progress_bar.setVisible(False)
        self.download_button.setEnabled(True)
        self.cancel_button.setEnabled(True)
        QMessageBox.information(self, "Download Complete", f"Captions saved to:\n{filepath}")
        logging.info(f"Download Complete. Captions saved to: {filepath}")
        try:
            webbrowser.open(self.output_dir.as_uri()) 
        except Exception as e:
            self.log(f"Could not open output directory: {e}")
            logging.warning(f"Could not open output directory {self.output_dir.as_uri()}: {e}")

if __name__ == "__main__":
    import sys
    from PySide6.QtWidgets import QApplication

    log_dir = Path(os.environ["USERPROFILE"]) / "Downloads"
    log_dir.mkdir(parents=True, exist_ok=True) 
    log_file_path = log_dir / "youtube_caption_downloader.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(threadName)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file_path),
            logging.StreamHandler(sys.stdout) 
        ]
    )
    logging.info("Application started.")

    app = QApplication(sys.argv)
    window = YouTubeCaptionFetcher()
    window.show()
    sys.exit(app.exec())
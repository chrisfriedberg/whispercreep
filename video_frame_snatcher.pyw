import sys
import os
import cv2
import subprocess
import math
import time
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLineEdit, QLabel, QProgressBar, QMessageBox,
    QHBoxLayout, QScrollArea, QGridLayout, QDialog
)
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt, QThread, Signal
import traceback

class FrameExtractor(QThread):
    progress = Signal(int)
    finished = Signal()

    def __init__(self, video_path, output_dir, fps):
        super().__init__()
        self.video_path = video_path
        self.output_dir = output_dir
        self.fps = fps

    def run(self):
        video = cv2.VideoCapture(self.video_path)
        actual_fps = video.get(cv2.CAP_PROP_FPS)
        total_frames = int(video.get(cv2.CAP_PROP_FRAME_COUNT))
        print(f"DEBUG FrameExtractor: Video has {total_frames} frames at {actual_fps} FPS. Target FPS = {self.fps}")

        if actual_fps <= 0 or self.fps <= 0:
            interval = 1
            print(f"DEBUG FrameExtractor: Using default interval=1 (invalid FPS values)")
        else:
            # The original interval calculation
            interval = max(1, int(actual_fps / self.fps))
            print(f"DEBUG FrameExtractor: Using interval={interval} (actual_fps={actual_fps} / target_fps={self.fps})")
            print(f"DEBUG FrameExtractor: This means taking 1 frame every {interval} frames")
            print(f"DEBUG FrameExtractor: Expected output frames ≈ {total_frames / interval}")

        count = 0

        while True:
            ret, frame = video.read()
            if not ret:
                break
            current_frame = int(video.get(cv2.CAP_PROP_POS_FRAMES))
            if current_frame % interval == 0:
                output_file = os.path.join(self.output_dir, f'snapshot_{count:04}.jpg')
                cv2.imwrite(output_file, frame)
                count += 1
            if total_frames > 0:
                self.progress.emit(int(current_frame / total_frames * 100))

        video.release()
        self.finished.emit()

class ReviewDialog(QDialog):
    def __init__(self, image_files):
        super().__init__()
        self.setWindowTitle("Review Output")
        self.setFixedSize(1000, 700)
        self.image_files = image_files
        self.current_index = 0

        self.layout = QVBoxLayout()

        self.position_label = QLabel()
        self.layout.addWidget(self.position_label)

        self.position_progress = QProgressBar()
        self.position_progress.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.position_progress)

        self.image_grid = QGridLayout()
        self.scroll_area = QScrollArea()
        self.image_widget = QWidget()
        self.image_widget.setLayout(self.image_grid)
        self.scroll_area.setWidget(self.image_widget)
        self.scroll_area.setWidgetResizable(True)
        self.layout.addWidget(self.scroll_area)

        nav_layout = QHBoxLayout()

        self.jump_to_start_btn = QPushButton("Jump to Start")
        self.jump_to_start_btn.clicked.connect(self.jump_to_start)
        nav_layout.addWidget(self.jump_to_start_btn)

        self.jump_back_btn = QPushButton("Jump Back 5%")
        self.jump_back_btn.clicked.connect(self.jump_back)
        nav_layout.addWidget(self.jump_back_btn)

        self.prev_btn = QPushButton("Previous")
        self.prev_btn.clicked.connect(self.show_previous)
        nav_layout.addWidget(self.prev_btn)

        self.next_btn = QPushButton("Next")
        self.next_btn.clicked.connect(self.show_next)
        nav_layout.addWidget(self.next_btn)

        self.jump_forward_btn = QPushButton("Jump Forward 5%")
        self.jump_forward_btn.clicked.connect(self.jump_forward)
        nav_layout.addWidget(self.jump_forward_btn)

        self.jump_to_end_btn = QPushButton("Jump to End")
        self.jump_to_end_btn.clicked.connect(self.jump_to_end)
        nav_layout.addWidget(self.jump_to_end_btn)

        self.layout.addLayout(nav_layout)

        close_btn_layout = QHBoxLayout()
        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet("background-color: red; color: white; font-weight: bold;")
        self.close_btn.clicked.connect(self.close)
        close_btn_layout.addStretch()
        close_btn_layout.addWidget(self.close_btn)
        self.layout.addLayout(close_btn_layout)

        self.setLayout(self.layout)

        self.show_images()

    def update_position_info(self):
        start = self.current_index + 1
        end = min(self.current_index + 10, len(self.image_files))
        self.position_label.setText(f"Viewing frames {start}-{end} of {len(self.image_files)}")
        progress = int((self.current_index / max(1, len(self.image_files))) * 100)
        self.position_progress.setValue(progress)

    def show_images(self):
        for i in reversed(range(self.image_grid.count())):
            widget = self.image_grid.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        for i in range(10):
            idx = self.current_index + i
            if idx >= len(self.image_files):
                break
            pixmap = QPixmap(self.image_files[idx]).scaled(200, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            label = QLabel()
            label.setPixmap(pixmap)
            label.mousePressEvent = self.create_mouse_events(self.image_files[idx])
            self.image_grid.addWidget(label, 0, i)

        self.update_position_info()

    def create_mouse_events(self, image_file):
        def handler(event):
            if event.button() == Qt.LeftButton:
                self.open_full_image(image_file)
            elif event.button() == Qt.RightButton:
                self.open_in_system(image_file)
        return handler

    def show_next(self):
        if self.current_index + 10 < len(self.image_files):
            self.current_index += 10
            self.show_images()

    def show_previous(self):
        if self.current_index - 10 >= 0:
            self.current_index -= 10
            self.show_images()

    def jump_forward(self):
        jump_amount = max(1, int(len(self.image_files) * 0.05))
        self.current_index = min(self.current_index + jump_amount, len(self.image_files) - 10)
        self.show_images()

    def jump_back(self):
        jump_amount = max(1, int(len(self.image_files) * 0.05))
        self.current_index = max(self.current_index - jump_amount, 0)
        self.show_images()

    def jump_to_start(self):
        self.current_index = 0
        self.show_images()

    def jump_to_end(self):
        self.current_index = max(0, len(self.image_files) - 10)
        self.show_images()

    def open_full_image(self, image_file):
        dlg = QDialog(self)
        dlg.setWindowTitle(os.path.basename(image_file))
        dlg.setFixedSize(800, 600)
        layout = QVBoxLayout()
        label = QLabel()
        pixmap = QPixmap(image_file).scaled(800, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        label.setPixmap(pixmap)
        layout.addWidget(label)
        dlg.setLayout(layout)
        dlg.exec()

    def open_in_system(self, image_file):
        if sys.platform == "win32":
            os.startfile(image_file)
        elif sys.platform == "darwin":
            subprocess.run(["open", image_file])
        else:
            subprocess.run(["xdg-open", image_file])

class VideoFrameSnatcher(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WhisperCreep Frame Snatcher")
        self.setFixedSize(400, 380)
        self.layout = QVBoxLayout()

        self.video_label = QLabel("No video selected.")
        self.video_label.setTextFormat(Qt.RichText)  # Enable rich text formatting
        self.layout.addWidget(self.video_label)

        self.browse_video_btn = QPushButton("Browse for Video")
        self.browse_video_btn.setStyleSheet("background-color: blue; color: white;")
        self.browse_video_btn.clicked.connect(self.browse_video)
        self.layout.addWidget(self.browse_video_btn)

        self.fps_input = QLineEdit()
        self.fps_input.setPlaceholderText("Enter seconds between snapshots (e.g., 1)")
        self.fps_input.textChanged.connect(self.update_eta)
        self.layout.addWidget(self.fps_input)

        self.eta_label = QLabel("ETA: N/A")
        self.layout.addWidget(self.eta_label)

        self.output_label = QLabel("No output folder selected.")
        self.layout.addWidget(self.output_label)

        self.browse_output_btn = QPushButton("Select Output Folder")
        self.browse_output_btn.setStyleSheet("background-color: blue; color: white;")
        self.browse_output_btn.clicked.connect(self.browse_output)
        self.layout.addWidget(self.browse_output_btn)

        self.process_btn = QPushButton("Process")
        self.process_btn.setStyleSheet("background-color: green; color: white;")
        self.process_btn.clicked.connect(self.process_video)
        self.layout.addWidget(self.process_btn)

        self.review_btn = QPushButton("Review Output")
        self.review_btn.setStyleSheet("background-color: darkgoldenrod; color: white;")
        self.review_btn.clicked.connect(self.review_output)
        self.layout.addWidget(self.review_btn)

        self.find_old_btn = QPushButton("Find/Review Old Output")
        self.find_old_btn.setStyleSheet("background-color: blue; color: white;")
        self.find_old_btn.clicked.connect(self.find_old_output)
        self.layout.addWidget(self.find_old_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.progress_bar)
        
        # Add a red close button at the bottom right
        close_layout = QHBoxLayout()
        close_layout.addStretch()
        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet("background-color: red; color: white;")
        self.close_btn.clicked.connect(self.close)
        close_layout.addWidget(self.close_btn)
        self.layout.addLayout(close_layout)

        self.setLayout(self.layout)

        self.video_path = None
        self.output_dir = None
        self.file_size = None
        self.spinner = None

    def browse_video(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Video File", "", "Video Files (*.mp4 *.avi *.mov)")
        if file:
            self.video_path = file
            self.file_size = os.path.getsize(file)

            # Display formatted file size (already correct)
            if self.file_size < 1024 * 1024:
                size_str = f"{self.file_size / 1024:.1f} KB"
            elif self.file_size < 1024 * 1024 * 1024:
                size_str = f"{self.file_size / (1024 * 1024):.1f} MB"
            else:
                size_str = f"{self.file_size / (1024 * 1024 * 1024):.2f} GB"

            self.video_label.setText(f"Video: {os.path.basename(file)} <span style='color: red;'>({size_str})</span>")

            # Get video properties SAFELY
            try:
                video = cv2.VideoCapture(file)
                fps = video.get(cv2.CAP_PROP_FPS)
                frame_count = video.get(cv2.CAP_PROP_FRAME_COUNT)
                self.video_width = int(video.get(cv2.CAP_PROP_FRAME_WIDTH))
                self.video_height = int(video.get(cv2.CAP_PROP_FRAME_HEIGHT))
                video.release()

                if fps is None or fps <= 0:
                    print(f"⚠ Invalid FPS ({fps}) detected. Duration cannot be calculated.")
                    self.video_duration = None
                else:
                    self.video_duration = frame_count / fps
                    print(f"✅ FPS: {fps}, Frames: {frame_count}, Duration: {self.video_duration} seconds")

            except Exception as e:
                print(f"Error reading video properties: {e}")
                self.video_duration = None
                self.video_width = None
                self.video_height = None

            self.update_eta()

    def browse_output(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if dir:
            self.output_dir = dir
            self.output_label.setText(f"Output: {dir}")

    def update_eta(self):
        try:
            if not self.fps_input.text().strip():
                self.eta_label.setText("ETA: N/A (Missing interval)")
                return

            seconds_between_snapshots = float(self.fps_input.text())
            if seconds_between_snapshots <= 0:
                self.eta_label.setText("ETA: N/A (Invalid interval)")
                return

            print(f"DEBUG UPDATE_ETA: Video properties - duration={getattr(self, 'video_duration', None)}, width={getattr(self, 'video_width', None)}, height={getattr(self, 'video_height', None)}")
            
            if self.video_duration and self.video_width and self.video_height:
                print(f"DEBUG UPDATE_ETA: Calling estimate with valid properties")
                eta_seconds = self.estimate_frame_extraction_eta(
                    self.video_duration,
                    seconds_between_snapshots,
                    self.video_width, 
                    self.video_height
                )
                print(f"DEBUG UPDATE_ETA: Got result: {eta_seconds} seconds")
                
                # Format ETA nicely
                if eta_seconds < 60:
                    eta_str = f"~{int(eta_seconds)} seconds"
                elif eta_seconds < 3600:
                    minutes = int(eta_seconds / 60)
                    seconds = int(eta_seconds % 60)
                    eta_str = f"~{minutes}m {seconds}s"
                else:
                    hours = int(eta_seconds / 3600)
                    minutes = int((eta_seconds % 3600) / 60)
                    eta_str = f"~{hours}h {minutes}m"

                print(f"DEBUG UPDATE_ETA: Setting label to: 'ETA: {eta_str}'")
                self.eta_label.setText(f"ETA: {eta_str}")
            else:
                print(f"DEBUG UPDATE_ETA: Invalid properties, setting N/A")
                self.eta_label.setText("ETA: N/A (Invalid video properties)")
        except Exception as e:
            print(f"Error calculating ETA: {e}")
            print(f"STACK TRACE:", traceback.format_exc())
            self.eta_label.setText("ETA: N/A (calculation error)")

    def estimate_frame_extraction_eta(self, video_duration_seconds, frame_spacing_seconds,
                                     resolution_width, resolution_height,
                                     frames_per_second=5.5):
        """
        Estimate frame extraction time using empirical frames-per-second rate.
        Calibrated based on actual processing time from real-world tests:
        - 20 min video, 10 second intervals, 123 frames, 23 seconds processing time
        
        Parameters:
        - video_duration_seconds: Duration of video in seconds
        - frame_spacing_seconds: Time between snapshots in seconds
        - resolution_width: Width of video in pixels
        - resolution_height: Height of video in pixels
        - frames_per_second: Empirical processing rate (frames processed per second)
        
        Returns:
        - Estimated processing time in seconds
        """
        # Convert inputs to float and validate
        video_duration_seconds = float(video_duration_seconds)
        frame_spacing_seconds = float(frame_spacing_seconds)
        
        if video_duration_seconds <= 0 or frame_spacing_seconds <= 0:
            return 0  # Invalid inputs
            
        # Calculate number of frames that will be extracted
        # For a 20-minute video with 10-second intervals, this is 1200/10 = 120 frames
        num_frames = video_duration_seconds / frame_spacing_seconds
        
        # Use the calibrated processing rate (5.5 FPS) from real-world test
        # 123 frames ÷ 23 seconds ≈ 5.35 FPS, rounded to 5.5 for slight safety margin
        processing_time = num_frames / frames_per_second
        
        # Add fixed overhead time for loading, initializing (calibrated to real-world results)
        overhead_seconds = 3.0
        total_time = processing_time + overhead_seconds
        
        # Ensure at least 1 second for any valid inputs
        return max(1.0, round(total_time))

    def process_video(self):
        if not all([self.video_path, self.output_dir, self.fps_input.text()]):
            QMessageBox.warning(self, "Missing Info", "Please select video, output, and seconds between snapshots.")
            return

        try:
            seconds_between_snapshots = float(self.fps_input.text())
            # Convert to FPS for the FrameExtractor class
            fps = 1.0 / seconds_between_snapshots if seconds_between_snapshots > 0 else 0
            print(f"DEBUG: Converting {seconds_between_snapshots} seconds between snapshots to {fps} FPS")
        except:
            QMessageBox.warning(self, "Invalid Value", "Enter a valid number for seconds between snapshots.")
            return

        self.progress_bar.setValue(0)
        self.process_btn.setEnabled(False)
        self.spinner = FrameExtractor(self.video_path, self.output_dir, fps)
        self.spinner.progress.connect(self.progress_bar.setValue)
        self.spinner.finished.connect(self.process_done)
        self.spinner.start()

    def process_done(self):
        QMessageBox.information(self, "Done", "Frame extraction completed.")
        self.progress_bar.setValue(100)
        self.process_btn.setEnabled(True)
        self.spinner = None

    def review_output(self):
        if not self.output_dir:
            QMessageBox.warning(self, "No Output", "Please select an output folder first.")
            return

        images = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir)
                  if f.lower().endswith(('.jpg', '.png', '.jpeg'))]
        images.sort()

        if not images:
            QMessageBox.warning(self, "No Images", "No snapshots found in the selected output folder.")
            return

        dlg = ReviewDialog(images)
        dlg.exec()

    def find_old_output(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Snapshot Folder")
        if not dir:
            return

        images = [os.path.join(dir, f) for f in os.listdir(dir)
                  if f.lower().endswith(('.jpg', '.png', '.jpeg'))]
        images.sort()

        if not images:
            QMessageBox.warning(self, "No Images", "No snapshots found in the selected folder.")
            return

        dlg = ReviewDialog(images)
        dlg.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VideoFrameSnatcher()
    window.show()
    sys.exit(app.exec())

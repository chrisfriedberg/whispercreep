# whispercreep
This is an all in one video to text, audio to text or video to audio transcription tool
# WhisperCreep  
### Video & Audio Transcription & Monitor App (Paranoid Edition)

**WhisperCreep** is a Windows-first Python app that lets you:
- Transcribe video files into timestamped text using OpenAI's Whisper.
- Extract audio from videos (MP3 format).
- Transcribe audio-only files.
- Monitor a folder for incoming video files and auto-transcribe them.
- Pause when manual transcription is running.
- Robust file handling with retry logic, permissions checks, file lock detection, and paranoid logging.

---

## 🛠 Features  
- GUI built with PySide6 — no CLI.
- Supports video and audio inputs (`.mp4`, `.mov`, `.avi`, `.wmv`, `.mkv`, `.mp3`, `.wav`, `.m4a`).
- Uses `ffmpeg` for audio extraction.
- Uses OpenAI Whisper (`tiny`, `base`, `small`, `medium`, `large` models).
- System Tray integration with minimize, resume, kill options.
- Monitors folders and safely queues files.
- Paranoid-level logging to per-run logs and system logs.
- Handles network shares, checks file locks, permissions, and stability before touching files.
- Uses temp files and only writes the final output if the transcription succeeds.

---

## 💾 Requirements  
- Windows 10 or higher  
- Python 3.10+  
- Installed libraries:
  - `whisper`
  - `torch`
  - `PySide6`
  - `ffmpeg` (must be installed and in your PATH)
  - `pywin32`
  - `shutil`, `tempfile`, `logging` (standard Python modules)

---

## 🚀 How to Use

### 1. Manual Mode (GUI)
- Launch the app.
- Select an operation (Video Transcript, Rip Audio, Audio Transcript).
- Select your source file.
- Click **Run**.
- The result will appear in your Downloads folder with a timestamp.

### 2. Monitor Mode (GUI)
- From the **Tools menu**, select **Monitor Folder for Video**.
- Pick the folder to monitor.
- Pick an output folder.
- Set the time increment between checks.
- Click **Start Monitoring**.
- The app will monitor the folder from the system tray.
- Will pause automatically if you're doing manual transcription.

---

## ⚡ Known Safety Features (Paranoid Additions)
- Files are checked for stability (not actively being written).
- Uses temp files and only moves to final output if successful.
- Handles file locks gracefully.
- Detects lost permissions (network paths, mapped drives).
- Runs retries with exponential backoff if failures happen.
- Warns you before overwriting existing files.

---

## 💡 Notes
- This tool is over-engineered to avoid **half-finished files, race conditions, and permission errors**.
- All output and logs land in your Downloads folder by default.
- Works best on local drives but supports network shares with extra checks.

---

## 🧑‍💻 Credits
Built by **Chris Friedberg**
Fork it, break it, improve it — but respect the paranoia level.


---

WhisperCreep: Overview
WhisperCreep is a Windows-focused application with two main components:
WhisperCreepUI: A comprehensive audio and video transcription tool that:
Transcribes video files to timestamped text using OpenAI's Whisper
Extracts audio from videos in MP3 format
Transcribes audio-only files
Includes a folder monitoring system that watches for new video files and automatically transcribes them
Offers a system tray integration with minimize, resume, and kill options
Provides detailed logging for troubleshooting
Video Frame Snatcher: A utility that:
Extracts frames from videos at specified intervals (FPS)
Allows review of extracted frames with thumbnail previews
Provides navigation controls to browse through the extracted frames
Shows estimated completion time based on file size
Key Features
GUI-Based Interface: Built with PySide6 (Qt for Python), no command-line required
Multiple File Format Support: Handles various video formats (MP4, MOV, AVI, WMV, MKV) and audio formats (MP3, WAV, M4A)
Robust File Handling: Includes checks for file locks, permissions, file stability, and network path accessibility
Temporary File Processing: Uses temp files to ensure only successful transcriptions become final output
Paranoid-Level Logging: Detailed logging to help diagnose issues
System Tray Integration: Allows the app to run in the background while monitoring folders
User-Friendly Features: Progress indicators, error handling, and confirmation dialogs
Technical Implementation
Uses OpenAI's Whisper for speech-to-text conversion
Uses ffmpeg for audio extraction from video files
Uses OpenCV (in Video Frame Snatcher) for frame extraction
Uses threading to prevent UI freezing during processing
Implements exponential backoff retry logic for network operations
Provides detailed error handling and user feedback
The application has been designed with reliability and paranoia in mind, ensuring that file operations are safe, network paths are validated, and that file locks are handled properly to avoid data corruption or incomplete processing.
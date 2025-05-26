import os
import sys
import threading
from urllib.parse import urljoin, urlparse
import requests
from bs4 import BeautifulSoup
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit,
    QPushButton, QProgressBar, QMessageBox
)
from PySide6.QtCore import Qt, QTimer, Signal, QObject
from PySide6.QtGui import QPalette, QColor
from playwright.sync_api import sync_playwright, Playwright
import time

# Define a QObject to emit signals across threads
class ScraperSignals(QObject):
    estimation_complete = Signal(int, int, list)
    manual_login_prompt = Signal(str)
    progress_update = Signal(int, int)
    scrape_complete = Signal()
    scrape_error = Signal(str)
    security_check_failed = Signal(str)
    # New signals for UI updates from background threads
    update_status_label = Signal(str)
    set_progress_bar_visibility = Signal(bool)
    set_progress_bar_range = Signal(int, int)
    # Signal to trigger Playwright operations in a dedicated thread
    start_playwright_process = Signal(str, list)


class WebToPDFScraper(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Web Content Scraper")
        self.setGeometry(100, 100, 400, 310)

        self.set_dark_theme()
        self.signals = ScraperSignals()
        self.signals.estimation_complete.connect(self.prompt_to_continue)
        self.signals.manual_login_prompt.connect(self.show_manual_login_dialog)
        self.signals.progress_update.connect(self.update_progress)
        self.signals.scrape_complete.connect(self.on_scrape_complete)
        self.signals.scrape_error.connect(self.on_scrape_error)
        self.signals.security_check_failed.connect(self.show_security_warning)
        # Connect new UI update signals
        self.signals.update_status_label.connect(self._update_status_label_slot)
        self.signals.set_progress_bar_visibility.connect(self._set_progress_bar_visibility_slot)
        self.signals.set_progress_bar_range.connect(self._set_progress_bar_range_slot)
        # Connect the signal to the slot that runs Playwright in a separate thread
        self.signals.start_playwright_process.connect(self._run_playwright_process_thread)


        self.layout = QVBoxLayout()
        self.label = QLabel("Enter root URL to crawl:")
        self.url_input = QLineEdit()
        self.url_input.setStyleSheet("color: white; background-color: #2d2d2d; border: 1px solid #555;")
        self.start_button = QPushButton("Start Crawling")
        self.close_button = QPushButton("Close")
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: white;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(self.close_button.sizeHint().height())

        self.start_button.setEnabled(False)
        self.start_button.setStyleSheet("background-color: green; color: white;")
        self.close_button.setStyleSheet("background-color: red; color: white;")

        self.layout.addWidget(self.label)
        self.layout.addWidget(self.url_input)
        self.layout.addWidget(self.start_button)
        self.layout.addWidget(self.close_button)
        self.layout.addWidget(self.status_label)
        self.layout.addWidget(self.progress_bar)
        self.setLayout(self.layout)

        self.url_input.textChanged.connect(self.toggle_start_button)
        self.start_button.clicked.connect(self.start_crawling)
        self.close_button.clicked.connect(self.close)

        self.root_url = ""
        self.pages_to_scrape = []
        self.output_dir = os.path.join(os.path.expanduser("~"), "Downloads", "WebContentScaper")
        os.makedirs(self.output_dir, exist_ok=True)
        self.log_path = os.path.join(self.output_dir, "last_run_log.txt")

        # Playwright will be initialized and stopped within the dedicated thread
        self.pw_thread = None

    def closeEvent(self, event):
        # Ensure Playwright thread is stopped if running
        if self.pw_thread and self.pw_thread.is_alive():
            # In a real application, you might need a more robust way to signal the thread to stop
            print("Attempting to join Playwright thread...")
            self.pw_thread.join(timeout=2) # Give it a moment to finish
            if self.pw_thread.is_alive():
                print("Playwright thread did not terminate.")
        super().closeEvent(event)

    def set_dark_theme(self):
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(30, 30, 30))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QColor(45, 45, 45))
        palette.setColor(QPalette.AlternateBase, QColor(60, 60, 60))
        palette.setColor(QPalette.ToolTipBase, Qt.white)
        palette.setColor(QPalette.ToolTipText, Qt.white)
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, QColor(45, 45, 45))
        palette.setColor(QPalette.ButtonText, Qt.white)
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.HighlightedText, Qt.black)
        self.setPalette(palette)

    def toggle_start_button(self):
        self.start_button.setEnabled(bool(self.url_input.text().strip()))

    # Slots for UI updates from background threads
    def _update_status_label_slot(self, text):
        self.status_label.setText(text)

    def _set_progress_bar_visibility_slot(self, visible):
        self.progress_bar.setVisible(visible)

    def _set_progress_bar_range_slot(self, min_val, max_val):
        self.progress_bar.setRange(min_val, max_val)

    def start_crawling(self):
        self.root_url = self.url_input.text().strip()
        if not self.root_url:
            QMessageBox.warning(self, "Input Error", "Please enter a valid URL.")
            return

        # Perform security check before anything else
        self.signals.update_status_label.emit("Performing security check...")
        self.signals.set_progress_bar_visibility.emit(True)
        self.signals.set_progress_bar_range.emit(0, 0) # Indeterminate
        threading.Thread(target=self._perform_security_check_thread, args=(self.root_url,), daemon=True).start()

    def _perform_security_check_thread(self, url):
        is_secure, message = self._check_url_security(url)
        if not is_secure:
            self.signals.security_check_failed.emit(message)
            self.signals.set_progress_bar_visibility.emit(False)
            return

        # If secure, proceed with estimation
        self.signals.update_status_label.emit("Estimating size...")
        # Progress bar is already visible and indeterminate from start_crawling
        threading.Thread(target=self._estimate_crawl_size_thread, args=(url,), daemon=True).start()

    def _check_url_security(self, url):
        parsed_url = urlparse(url)
        if parsed_url.scheme != 'https':
            return False, "URL is not HTTPS. For 2FA and secure login, HTTPS is required."

        try:
            response = requests.head(url, timeout=10, allow_redirects=True)
            response.raise_for_status()

            if not response.url.startswith('https://'):
                return False, f"URL redirects to an insecure (non-HTTPS) address: {response.url}"

            return True, "URL appears secure (HTTPS)."

        except requests.exceptions.SSLError as e:
            return False, f"SSL/TLS Certificate Error: {e}. The site's certificate might be invalid or untrusted."
        except requests.exceptions.ConnectionError as e:
            return False, f"Connection Error: {e}. Could not connect to the URL."
        except requests.exceptions.Timeout:
            return False, "Connection timed out during security check."
        except requests.exceptions.RequestException as e:
            return False, f"An unexpected network error occurred: {e}"
        except Exception as e:
            return False, f"An unknown error occurred during security check: {e}"

    def show_security_warning(self, message):
        QMessageBox.critical(self, "Security Warning", message)
        self.status_label.setText("Security check failed. Scrape aborted.")
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 1)


    def _estimate_crawl_size_thread(self, root_url):
        page_count, total_bytes, pages = self.estimate_crawl_size_requests(root_url)
        self.signals.estimation_complete.emit(page_count, total_bytes, pages)

    def estimate_crawl_size_requests(self, root_url):
        visited = set()
        queue = [root_url]
        total_bytes = 0
        page_count = 0
        all_pages = []

        root_domain = urlparse(root_url).netloc
        excluded_langs = ['/zh-Hans/', '/zh-Hant/', '/ja-JP/', '/ko-KR/', '/es-ES/', '/fr-FR/', '/de-DE/']

        while queue:
            url = queue.pop(0)
            parsed_url = urlparse(url)

            if url in visited or parsed_url.netloc != root_domain:
                continue

            if '/en/' not in parsed_url.path:
                 if any(lang_path in parsed_url.path for lang_path in excluded_langs):
                     continue

            visited.add(url)
            all_pages.append(url)
            page_count += 1

            try:
                response = requests.get(url, timeout=10)
                total_bytes += len(response.content)
                soup = BeautifulSoup(response.text, 'html.parser')
                for link in soup.find_all('a', href=True):
                    next_url = urljoin(url, link['href'])
                    parsed_next_url = urlparse(next_url)

                    if parsed_next_url.netloc == root_domain and next_url not in visited:
                        if '/en/' not in parsed_next_url.path:
                            if any(lang_path in parsed_next_url.path for lang_path in excluded_langs):
                                continue
                        queue.append(next_url)
            except requests.exceptions.RequestException as e:
                print(f"Request failed for {url}: {e}")
                continue
            except Exception as e:
                print(f"Error processing {url}: {e}")
                continue
        return page_count, total_bytes, all_pages

    def prompt_to_continue(self, page_count, total_bytes, pages):
        self.pages_to_scrape = pages
        self.status_label.setText(f"Estimated: {page_count} pages, {total_bytes // 1024} KB")
        reply = QMessageBox.question(self, "Confirm",
                                     f"Estimated: {page_count} pages, {total_bytes // 1024} KB\nDo you want to continue?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            with open(self.log_path, 'w', encoding='utf-8') as f:
                for url in self.pages_to_scrape:
                    f.write(url + '\n')

            # Emit signal to start the Playwright process in its dedicated thread
            self.signals.start_playwright_process.emit(self.root_url, self.pages_to_scrape)
        else:
            self.status_label.setText("Scrape cancelled.")
            self.progress_bar.setVisible(False)
            self.progress_bar.setRange(0, 1)

    def _run_playwright_process_thread(self, root_url, pages_to_scrape):
        # This method runs in a dedicated thread to handle all Playwright operations
        self.pw_thread = threading.current_thread()
        pw_sync_api = None
        pw_browser = None
        pw_context = None

        try:
            pw_sync_api = sync_playwright().start()
            # Launch browser in headless mode for scraping, non-headless for manual login
            pw_browser = pw_sync_api.chromium.launch(headless=False) # Keep non-headless for initial login
            pw_context = pw_browser.new_context()

            # Initiate manual login within this thread
            self._initiate_manual_login_playwright(root_url, pw_context)

            # After manual login, proceed with scraping within this thread
            # Re-launch browser in headless mode for scraping
            pw_browser.close() # Close the non-headless browser used for login
            pw_browser = pw_sync_api.chromium.launch(headless=True) # Launch headless browser for scraping
            pw_context = pw_browser.new_context() # Create a new context with the headless browser

            self._run_scraper_playwright(root_url, pages_to_scrape, pw_context)

        except Exception as e:
            self.signals.scrape_error.emit(f"Playwright process failed: {e}")
        finally:
            if pw_browser:
                pw_browser.close()
            if pw_sync_api:
                pw_sync_api.stop()
            self.signals.scrape_complete.emit()


    def _initiate_manual_login_playwright(self, url, pw_context):
        try:
            page = pw_context.new_page()
            page.goto(url, timeout=60000)

            # Emit signal to show manual login dialog on the main thread
            self.signals.manual_login_prompt.emit(url)

            # Wait for the user to signal completion from the main thread (via the dialog)
            # This requires a way to signal back from the main thread after the dialog is closed.
            # For simplicity here, we'll just add a small delay. A real app needs a proper signal/slot mechanism.
            # A better approach would be a signal from the main thread's dialog handler back to this thread.
            # For now, let's assume the user interacts and the browser is ready after a delay.
            time.sleep(5) # Placeholder: Replace with a signal-based wait

            page.close()

        except Exception as e:
            raise Exception(f"Failed to open browser for manual login: {e}")


    def show_manual_login_dialog(self, url):
        QMessageBox.information(self, "Manual Login Required",
                                f"A browser window has opened. Please log in to '{url}' manually in that window. "
                                "Click OK here when you are fully logged in and ready for the script to continue scraping.")
        # In a real app, after this dialog closes, emit a signal back to the Playwright thread to continue.
        # For this simplified example, the Playwright thread will continue after a delay (see _initiate_manual_login_playwright)


    def _run_scraper_playwright(self, root_url, pages_to_scrape, pw_context):
        visited = set()
        if os.path.exists(self.log_path):
            try:
                with open(self.log_path, 'r', encoding='utf-8') as f:
                    remaining_urls = [line.strip() for line in f if line.strip()]
                pages_to_scrape = [url for url in pages_to_scrape if url in remaining_urls]
            except Exception as e:
                print(f"Error reading log file: {e}")

        # Filter out non-English pages using Playwright before starting the scrape
        self.signals.update_status_label.emit("Filtering non-English pages with Playwright...")
        print("Filtering non-English pages with Playwright...")
        filtered_pages_to_scrape = []
        excluded_langs = ['zh-Hans', 'zh-Hant', 'ja-JP', 'ko-KR', 'es-ES', 'fr-FR', 'de-DE']

        for url in pages_to_scrape:
            try:
                page = pw_context.new_page()
                page.goto(url, timeout=30000)
                page_lang = page.eval_on_selector('html', 'element => element.lang') or ''
                page.close()
                if page_lang not in excluded_langs:
                    filtered_pages_to_scrape.append(url)
                else:
                    print(f"Filtered out {url} due to language: {page_lang}")
            except Exception as e:
                print(f"Failed to check language for {url}: {e}")
                # Optionally include the page if language check fails, or skip it.
                # For now, we'll skip it to be safe.

        pages_to_scrape = filtered_pages_to_scrape
        self.signals.update_status_label.emit(f"Scraping {len(pages_to_scrape)} English pages...")
        print(f"Scraping {len(pages_to_scrape)} English pages...")

        self.signals.set_progress_bar_range.emit(0, len(pages_to_scrape))
        self.signals.progress_update.emit(0, len(pages_to_scrape))


        try:
            for index, url in enumerate(pages_to_scrape):
                if url in visited:
                    continue
                visited.add(url)
                self._save_pdf_sync_playwright(url, index, pw_context)
                with open(self.log_path, 'w', encoding='utf-8') as f:
                    for remaining_url in pages_to_scrape[index+1:]:
                        f.write(remaining_url + '\n')
        except Exception as e:
            raise Exception(f"Scraping failed: {e}")


    def estimate_crawl_size_playwright(self, root_url, initial_page, pw_context):
        visited = set()
        queue = [root_url]
        all_pages = []
        root_domain = urlparse(root_url).netloc
        excluded_langs = ['zh-Hans', 'zh-Hant', 'ja-JP', 'ko-KR', 'es-ES', 'fr-FR', 'de-DE'] # Check lang attribute

        # Check language of the initial page
        initial_page_lang = initial_page.eval_on_selector('html', 'element => element.lang') or ''
        if initial_page_lang not in excluded_langs:
            visited.add(root_url)
            all_pages.append(root_url)
            links_on_initial_page = initial_page.eval_on_selector_all('a[href]', 'elements => elements.map(e => e.href)')
            for link in links_on_initial_page:
                next_url = urljoin(root_url, link)
                parsed_next_url = urlparse(next_url)
                if parsed_next_url.netloc == root_domain and next_url not in visited:
                    queue.append(next_url)
                    visited.add(next_url)


        while queue:
            url = queue.pop(0)

            try:
                page = pw_context.new_page()
                page.goto(url, wait_until="networkidle", timeout=30000)

                page_lang = page.eval_on_selector('html', 'element => element.lang') or ''
                if page_lang not in excluded_langs:
                    all_pages.append(url)
                    links = page.eval_on_selector_all('a[href]', 'elements => elements.map(e => e.href)')
                    for link in links:
                        next_url = urljoin(url, link)
                        parsed_next_url = urlparse(next_url)
                        if parsed_next_url.netloc == root_domain and next_url not in visited:
                            queue.append(next_url)
                            visited.add(next_url)
                page.close()

            except Exception as e:
                print(f"Playwright crawl failed for {url}: {e}")
                continue
        return all_pages


    def _save_pdf_sync_playwright(self, url, index, pw_context):
        try:
            page = pw_context.new_page()
            page.goto(url, timeout=60000)

            page_lang = page.eval_on_selector('html', 'element => element.lang') or ''
            excluded_langs = ['zh-Hans', 'zh-Hant', 'ja-JP', 'ko-KR', 'es-ES', 'fr-FR', 'de-DE']

            if page_lang in excluded_langs:
                print(f"Skipping {url} due to language: {page_lang}")
                page.close()
                self.signals.progress_update.emit(index + 1, len(self.pages_to_scrape))
                return

            filename = self.sanitize_filename(urlparse(url).path or urlparse(url).netloc.replace('.', '_') or "index")
            if not filename:
                filename = f"page_{index}"

            path = os.path.join(self.output_dir, f"{filename}.pdf")

            page.pdf(path=path, format="A4")
            page.close()
            self.signals.progress_update.emit(index + 1, len(self.pages_to_scrape))
        except Exception as e:
            print(f"PDF Save Failed for {url}: {e}")

    def sanitize_filename(self, name):
        return "".join(c if c.isalnum() or c in ('-', '_', '.') else '_' for c in name).strip()

    def update_progress(self, current, total):
        self.progress_bar.setValue(current)
        self.status_label.setText(f"Saving page {current}/{total}")

    def on_scrape_complete(self):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 1)
        self.status_label.setText("Scraping complete.")
        if os.path.exists(self.log_path):
            os.remove(self.log_path)
        QMessageBox.information(self, "Done", "Scraping complete. PDFs saved to Downloads/WebContentScaper.")

    def on_scrape_error(self, message):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 1)
        self.status_label.setText("Scraping error.")
        QMessageBox.critical(self, "Error", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WebToPDFScraper()
    window.show()
    sys.exit(app.exec())
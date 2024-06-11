import sys
import os
import time
import threading
import json
import keyboard
import pyautogui
import logging

from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton,
                             QListWidget, QMainWindow, QMenu, QAction, QActionGroup,
                             QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout, QInputDialog)
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

# Configure logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s | %(levelname)s: %(message)s', 
                    datefmt='%d.%m.%Y %H:%M:%S')

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

icon_path = resource_path("media/esmaabi_icon.ico") 

class WorkerThread(QThread):
    signal_finished = pyqtSignal()

    def __init__(self, messages, delay, mode, repeat_count):
        super().__init__()
        self.messages = messages
        self.delay = delay
        self.mode = mode
        self.repeat_count = repeat_count
        self.stop_event = threading.Event()
    
    def run(self):
        logging.info("Worker thread started")
        current_repeat = 0
        while self.repeat_count == 0 or current_repeat < self.repeat_count:
            for message in self.messages:
                if self.stop_event.is_set():
                    self.handle_stop()
                    return

                if self.mode == "yell":
                    message = "::yell " + message
                elif self.mode == "team_chat":
                    message = "/" + message
                time.sleep(self.delay)
                self.type_message(message)

            current_repeat += 1

        self.signal_finished.emit()
        logging.info("Worker thread finished")

    def type_message(self, message):
        if self.stop_event.is_set():
            self.handle_stop()
            return
        pyautogui.typewrite(message)
        pyautogui.press("enter")

    def handle_stop(self):
        self.signal_finished.emit()

    def stop(self):
        self.stop_event.set()

class AutoTyper(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Auto Typer for RSPS")
        self.messages = []
        self.delay = 5
        self.mode = "default"
        self.worker_thread = None
        self.repeat_count = 0
        self.setWindowIcon(QIcon(icon_path))
        self.setMinimumSize(450, 450)

        self.initUI()

    def initUI(self):
        self.create_menu_bar()
        self.create_central_widget()

        self.setStyleSheet("""
            QWidget {
                background-color: #1c1e26;
                color: #abb2bf;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }
            QLineEdit, QListWidget {
                background-color: #2e303e;
                color: #ffffff;
                border: 1px solid #4b4d5b;
                border-radius: 5px;
                padding: 5px;
                margin: 5px 0;
            }
            QLineEdit:focus, QListWidget:focus {
                border-color: #61afef;
            }
            QPushButton {
                background-color: #61afef;
                color: #ffffff;
                border: none;
                padding: 10px 15px;
                margin: 5px 0;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #56b6c2;
            }
            QPushButton:pressed {
                background-color: #34d058;
            }
            QMenuBar {
                background-color: #2e303e;
                color: #ffffff;
                border: 1px solid #4b4d5b;
            }
            QMenuBar::item {
                background: transparent;
                padding: 5px 10px;
            }
            QMenuBar::item:selected {
                background: #61afef;
                color: #ffffff;
                border-radius: 5px;
            }
            QMenu {
                background-color: #2e303e;
                color: #ffffff;
                border: 1px solid #4b4d5b;
            }
            QMenu::item {
                padding: 5px 20px;
            }
            QMenu::item:selected {
                background-color: #61afef;
                color: #ffffff;
            }
            QMessageBox {
                background-color: #1b1e25;
                color: #ffffff;
                border: 1px solid #4b4d5b;
            }
        """)
    
    def create_menu_bar(self):
        menu_bar = self.menuBar()

        # File Menu
        file_menu = QMenu("File", self)
        save_action = QAction("Save to File", self)
        save_action.triggered.connect(self.save_messages)
        load_action = QAction("Load from File", self)
        load_action.triggered.connect(self.load_messages)
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close) 
        file_menu.addAction(save_action)
        file_menu.addAction(load_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)

        # Mode Menu
        mode_menu = QMenu("Mode", self)
        mode_group = QActionGroup(self)
        mode_group.setExclusive(True)

        self.default_action = QAction("Default", self, checkable=True)
        self.default_action.setChecked(True)
        self.default_action.triggered.connect(lambda: self.set_mode("default"))
        mode_group.addAction(self.default_action)

        self.yell_action = QAction("Yell", self, checkable=True)
        self.yell_action.triggered.connect(lambda: self.set_mode("yell"))
        mode_group.addAction(self.yell_action)

        self.team_chat_action = QAction("Team Chat", self, checkable=True)
        self.team_chat_action.triggered.connect(lambda: self.set_mode("team_chat"))
        mode_group.addAction(self.team_chat_action)

        mode_menu.addAction(self.default_action)
        mode_menu.addAction(self.yell_action)
        mode_menu.addAction(self.team_chat_action)

        # Help Menu
        help_menu = QMenu("Help", self)
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_help)
        help_menu.addAction(about_action)

        menu_bar.addMenu(file_menu)
        menu_bar.addMenu(mode_menu)
        menu_bar.addMenu(help_menu)

    def create_central_widget(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        # Message List
        self.message_list = QListWidget(self)
        self.message_list.setDragDropMode(QListWidget.InternalMove)
        self.message_list.model().rowsMoved.connect(self.messages_reordered)
        layout.addWidget(self.message_list)

        # Input and Buttons
        input_layout = QHBoxLayout()
        self.message_input = QLineEdit(self)
        self.message_input.setPlaceholderText("Enter message...")
        input_layout.addWidget(self.message_input)

        self.add_button = QPushButton("Add", self)
        self.add_button.clicked.connect(self.add_message)
        input_layout.addWidget(self.add_button)

        self.edit_button = QPushButton("Edit", self)
        self.edit_button.clicked.connect(self.edit_message)
        input_layout.addWidget(self.edit_button)

        self.delete_button = QPushButton("Delete", self)
        self.delete_button.clicked.connect(self.delete_message)
        input_layout.addWidget(self.delete_button)
        layout.addLayout(input_layout)

        # Delay Setting
        delay_layout = QHBoxLayout()
        delay_label = QLabel("Message Rate (seconds):", self)
        delay_layout.addWidget(delay_label)
        self.delay_input = QLineEdit(self)
        self.delay_input.setText(str(self.delay))
        delay_layout.addWidget(self.delay_input)
        layout.addLayout(delay_layout)
        self.delay_input.textChanged.connect(self.update_delay)

        # Repeat Count Setting
        repeat_layout = QHBoxLayout()
        repeat_label = QLabel("Repeat Count:", self)
        repeat_layout.addWidget(repeat_label)
        self.repeat_input = QLineEdit(self)
        self.repeat_input.setText(str(self.repeat_count))
        repeat_layout.addWidget(self.repeat_input)
        layout.addLayout(repeat_layout)
        self.repeat_input.textChanged.connect(self.update_repeat_count)

        # Start/Stop Button
        self.start_stop_button = QPushButton("Start", self)
        self.start_stop_button.clicked.connect(self.toggle_typing)
        layout.addWidget(self.start_stop_button)

        # Set layout to central widget
        central_widget.setLayout(layout)

    def set_mode(self, mode):
        self.mode = mode

    def update_delay(self, text):
        """Updates the delay value when the delay_input text changes."""
        try:
            self.delay = int(text)
            logging.info(f"Delay updated to: {self.delay}")
        except ValueError:
            logging.warning("Invalid input for delay. Please enter a number.")

    def update_repeat_count(self, text):
        """Updates the repeat_count when the repeat_input text changes."""
        try:
            self.repeat_count = int(text)
            logging.info(f"Repeat Count updated to: {self.repeat_count}")
        except ValueError:
            logging.warning("Invalid input for repeat count. Please enter a number.")
    
    def messages_reordered(self):
        """Updates self.messages to reflect the current order in the list widget."""
        self.messages = [self.message_list.item(i).text() for i in range(self.message_list.count())]
        logging.info("Message order updated.") 


    def show_help(self):
         QMessageBox.about(self, "About Auto Typer",
                             "<html><head/><body>"
                             "<p>This application automates typing messages with the following features:</p>"
                             "<p><b>Features:</b></p>"
                             "<ul>"
                             "<li>Add, edit, and delete messages</li>"
                             "<li>Set the delay between messages</li>"
                             "<li>Choose from different modes: Default, Yell, Team Chat</li>"
                             "<li>Save and load message lists</li>"
                             "<li>Setting repeat count to 0 enables infinite loop</li>"
                             "</ul>"
                             "<p><b>Hotkey:</b></p>"
                             "<ul><li>F12: Start/Stop typing</li></ul>"
                             "<p>Application made by Esmaabi<br/>"
                             "GitHub: <a href='https://github.com/KristjanPikhof' style='color:#61afef; text-decoration:none;'>github.com/KristjanPikhof</a></p>"
                             "</body></html>")


    def add_message(self):
        message = self.message_input.text()
        if message:
            self.messages.append(message)
            self.update_message_list()
            self.message_input.clear()
            logging.info(f"Message added: {message}")

    def edit_message(self):
        try:
            selected_row = self.message_list.currentRow()
            if selected_row != -1: 
                current_message = self.messages[selected_row]
                new_message, ok = QInputDialog.getText(self, "Edit Message", 
                                                      "Edit the message:", 
                                                      QLineEdit.Normal, current_message)
                if ok and new_message:
                    self.messages[selected_row] = new_message
                    self.update_message_list()
                    logging.info(f'Message edited from "{current_message}" to "{new_message}"')
        except IndexError:
            QMessageBox.warning(self, "No Message Selected", "Please select a message to edit.")

    def delete_message(self):
        try:
            selected_row = self.message_list.currentRow()
            if selected_row != -1:
                logging.info(f"Message deleted: {self.messages[selected_row]}")
                del self.messages[selected_row]
                self.update_message_list()
        except IndexError:
            QMessageBox.warning(self, "No Message Selected", "Please select a message to delete.")

    def update_message_list(self):
        self.message_list.clear()
        for message in self.messages:
            self.message_list.addItem(message)

    def toggle_typing(self):
        if self.worker_thread and self.worker_thread.isRunning():
            self.stop_typing()
        else:
            self.start_typing()

    def start_typing(self):
        if not self.messages:
            QMessageBox.warning(self, "No Messages", "Please add messages to the list.")
            return
        
        self.delay = int(self.delay_input.text())
        self.repeat_count = int(self.repeat_input.text())

        logging.info("Starting typing process")
        self.messages = [self.message_list.item(i).text() for i in range(self.message_list.count())] 
        self.worker_thread = WorkerThread(self.messages.copy(), self.delay, self.mode, self.repeat_count)
        self.worker_thread.signal_finished.connect(self.on_typing_finished)
        self.worker_thread.start()
        self.start_stop_button.setText("Stop")

    def stop_typing(self):
        if self.worker_thread and self.worker_thread.isRunning():
            logging.info("Stopping typing process")
            self.worker_thread.stop()
            self.worker_thread.wait()
            self.activateWindow()
        self.start_stop_button.setText("Start")

    def on_typing_finished(self):
        """Called when the worker thread has finished typing."""
        self.start_stop_button.setText("Start")

    def save_messages(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Messages", "", "JSON Files (*.json)")
        if file_path:
            try:
                data = {
                    "messages": self.messages,
                    "delay": self.delay,
                    "mode": self.mode,
                    "repeat_count": self.repeat_count
                }
                with open(file_path, 'w') as f:
                    json.dump(data, f)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to save to file: {e}")

    def load_messages(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Load Messages", "", "JSON Files (*.json)")
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    self.messages = data.get("messages", [])
                    self.delay = data.get("delay", 5)
                    self.mode = data.get("mode", "default")
                    self.repeat_count = data.get("repeat_count", 1)

                self.delay_input.setText(str(self.delay))
                self.repeat_input.setText(str(self.repeat_count)) 
                self.update_message_list()

                # Update the mode selection based on loaded mode
                if self.mode == "default":
                    self.default_action.setChecked(True)
                elif self.mode == "yell":
                    self.yell_action.setChecked(True)
                elif self.mode == "team_chat":
                    self.team_chat_action.setChecked(True)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load from file: {e}")   

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Monospace", 10))
    window = AutoTyper()
    window.show()
    window.activateWindow()

    keyboard.on_press_key('f12', lambda _: window.toggle_typing())
    logging.info("F12 hotkey registered")

    sys.exit(app.exec_())
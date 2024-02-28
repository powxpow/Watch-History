'''watch_history_app'''
import logging
import os
import platform
import sys
from pathlib import Path, PurePath
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QFileDialog,
    QVBoxLayout,
    QWidget,
    QLabel,
    QPlainTextEdit,
    QPushButton )
from watch_history_core import WatchHistory, Feedback

class ProcessHistoryThread(QThread):
    '''ProcessHistoryThread'''
    parent = None
    thread_status = Signal(str)
    source_file = None
    dest_file = None
    feedback = None

    def __init__(self, parent, feedback, src, dest):
        QThread.__init__(self, parent)
        self.parent = parent #set the parent to keep the thread Signal around
        self.source_file = src
        self.dest_file = dest
        self.feedback = feedback

    def run(self):
        '''run'''
        watch_history = WatchHistory(self.feedback)
        watch_history.process_yt_history(self.source_file, self.dest_file)

class WatchHistoryApp(QMainWindow):
    '''WatchHistoryApp'''
    source_file = None
    dest_folder = os.path.expanduser('~/Downloads')
    dest_file = None
    run_thread = None
    #Keep the feedback object around so we don't reinitialize and get double messages
    run_feedback = Feedback(stream=sys.stdout)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Watch History App")
        self.setGeometry(100, 100, 600, 500)

        #styles
        self.style_path_unset = "border: 1px solid gray; border-radius: 6px;"
        self.style_path_set =  "border: 1px solid lime; border-radius: 6px;"
        self.style_run_ready = "background-color: green; color: white;"
        self.style_console = "background-color: black; color: lightgreen;"
        self.style_console += "font-family: 'Cascadia Code', monospace;"
        self.style_console += "border: 1px solid white; border-radius: 6px;margin-left: 2px"

        # Create widgets
        self.source_label = QLabel("Google Takeout File:")
        self.source_path_label = QLabel()
        self.source_path_label.setStyleSheet(self.style_path_unset)
        self.source_button = QPushButton("Select Google Takeout File")
        self.source_button.clicked.connect(self.pick_source_file)

        self.dest_label = QLabel("Destination Folder:")
        home =os.path.expanduser('~')
        self.dest_path_label = QLabel(self.dest_folder.replace(home, '~'))
        self.dest_path_label.setWordWrap(True)
        self.dest_path_label.setStyleSheet(self.style_path_set)
        self.dest_path_button = QPushButton("Select Destination Folder")
        self.dest_path_button.clicked.connect(self.pick_destination_folder)
        self.dest_open_button = QPushButton("Open Destination Folder")
        self.dest_open_button.clicked.connect(self.open_destination_folder)

        self.run_button = QPushButton("Process Takeout")
        self.run_button.setEnabled(False)
        self.run_button.clicked.connect(self.run_history_thread)
        self.run_console = QPlainTextEdit()
        self.run_console.setReadOnly(True)
        self.run_console.setStyleSheet(self.style_console)

        # Layout
        box = QVBoxLayout()
        box.addWidget(self.source_label)
        box.addWidget(self.source_path_label)
        box.addWidget(self.source_button)
        box.addWidget(self.dest_label)
        box.addWidget(self.dest_path_label)
        box.addWidget(self.dest_path_button)
        box.addWidget(self.run_button)
        box.addWidget(self.run_console)
        box.addWidget(self.dest_open_button)

        central_widget = QWidget()
        central_widget.setLayout(box)
        self.setCentralWidget(central_widget)

    def pick_source_file(self):
        '''pick_source_file'''
        file, _ = QFileDialog.getOpenFileName(self, "Select Google Takeout", "", "All Files (*)")
        if file and Path(file).suffix in ['.zip', '.html', '.json']:
            self.source_path_label.setText(Path(file).name)
            self.source_path_label.setStyleSheet(self.style_path_set)
            self.source_file = file
            self.run_button.setEnabled(True)
            self.run_button.setStyleSheet(self.style_run_ready)
            self.dest_open_button.setStyleSheet(None)
            self.run_console.clear()

    def pick_destination_folder(self):
        '''pick_destination_folder'''
        dest = self.dest_folder
        folder_name = QFileDialog.getExistingDirectory(self, "Destination Folder", dest)
        if folder_name:
            self.dest_folder = folder_name
            home =os.path.expanduser('~')
            self.dest_path_label.setText(folder_name.replace(home,'~'))
            self.dest_path_label.setStyleSheet(self.style_path_set)

    def open_destination_folder(self):
        '''open_destination_folder'''
        cmd = 'open'
        match platform.system():
            case 'Windows':
                cmd = 'explorer'
            case 'Linux':
                cmd = 'xdg-open'
        os.system(f'{cmd} {Path(self.dest_folder)}', )  #Open folder in default file explorer

    def run_history_thread(self):
        '''run_history_thread'''
        self.run_console.clear()
        self.run_button.setStyleSheet(None)
        self.run_button.setEnabled(False)
        if self.source_file is not None:
            logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s\t%(message)s')
            self.thread_start()

    def thread_start(self):
        '''thread_start'''
        if self.run_thread is None:
            xlsx_file = Path(self.source_file).name.replace(Path(self.source_file).suffix, '.xlsx')
            dest_file = PurePath(self.dest_folder, xlsx_file)
            self.dest_file = dest_file
            self.run_thread = ProcessHistoryThread(self, self.run_feedback,
                self.source_file, self.dest_file)
            self.run_feedback.signal = self.run_thread.thread_status
            self.run_thread.thread_status.connect(self.thread_update)
            self.run_thread.finished.connect(self.thread_finished)
            self.run_thread.start()

    def thread_update(self, run_feedback):
        '''thread_update'''
        self.run_console.appendPlainText(run_feedback)

    def thread_finished(self):
        '''thread_finished'''
        self.run_thread = None
        self.dest_open_button.setStyleSheet(self.style_run_ready)
        #clear "source" so it is ready for another file
        self.source_file = None
        self.source_path_label.setText("")
        self.source_path_label.setStyleSheet(self.style_path_unset)

if __name__ == "__main__":
    if platform.system() == 'Windows':
        sys.argv += ['-platform', 'windows:darkmode=2']
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = WatchHistoryApp()
    window.show()
    sys.exit(app.exec())

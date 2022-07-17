from PyQt5.QtWidgets import *
from PyQt5.QtCore import QUrl
from PyQt5.QtMultimedia import QMediaContent
from pathlib import Path


song, _ = QFileDialog.getOpenFileName( "Open Song", "~", "Sound Files (*.mp3 *.ogg *.wav *.m4a)")
print(song)
url = QUrl.fromLocalFile(song)
filename = Path(song).name

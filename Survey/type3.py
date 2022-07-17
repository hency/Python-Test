import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class Time(QWidget):
    def __init__(self):
        super(Time, self).__init__()
        self.setWindowTitle("动态显示时间")

        self.label = QLabel("当前时间")
        self.start = QPushButton("Start")
        self.end = QPushButton("End")

        layout = QGridLayout()

        self.timer = QTimer()
        self.timer.timeout.connect(self.showTime)

        layout.addWidget(self.label, 0, 0, 1, 2)
        layout.addWidget(self.start, 1, 0)
        layout.addWidget(self.end, 1, 1)

        self.start.clicked.connect(self.startTimer)
        self.end.clicked.connect(self.endTimer)

        self.setLayout(layout)

    def showTime(self):
        time = QDateTime.currentDateTime()
        timeDisplay = time.toString("yyyy-MM-dd hh:mm:ss dddd")
        self.label.setText(timeDisplay)
        print(self.label.text())

    def startTimer(self):
        self.timer.start(10000)
        self.start.setEnabled(False)
        self.end.setEnabled(True)

    def endTimer(self):
        self.timer.stop()
        self.start.setEnabled(True)
        self.end.setEnabled(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = Time()
    main.show()
    sys.exit(app.exec_())
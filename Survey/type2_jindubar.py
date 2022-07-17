# from PyQt5.Qt import *
#
# class MyQProgressBar(QProgressBar):
#     def timerEvent(self, evt):
#         value = self.value()
#         if value <= self.maximum()-1:
#             self.setValue(value+1)
#         else:
#             self.killTimer(evt.timerId())
#
# class MyWindow(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("进度条")
#         self.resize(500,500)
#         self.setup_ui()
#
#     def setup_ui(self):
#         self.qpb = MyQProgressBar(self)
#         self.qpb.resize(300,30)
#         self.qpb.setValue(20)
#
#         # self.qpb.startTimer(1000,Qt.VeryCoarseTimer)
#
#         self.btn = QPushButton("开始",self)
#         self.btn.move(0,50)
#
#         self.btn.clicked.connect(self.change_progressbar)
#
#     def change_progressbar(self):
#         if self.btn.text() == "开始":
#             self.btn.setText("结束")
#             self.qpb_time_id = self.qpb.startTimer(1000,Qt.VeryCoarseTimer)
#         else:
#             self.btn.setText("开始")
#             self.qpb.killTimer(self.qpb_time_id)
#
# if __name__ == '__main__':
#     import sys
#     app = QApplication(sys.argv)
#     window = MyWindow()
#     window.show()
#     sys.exit(app.exec())
from PyQt5.Qt import *

class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("进度条")
        self.resize(500,500)
        self.setup_ui()

    def setup_ui(self):
        self.qpb = QProgressBar(self)
        self.qpb.resize(300,30)
        self.qpb.setValue(0)

        # self.qpb.startTimer(1000,Qt.VeryCoarseTimer)

        self.btn = QPushButton("开始",self)
        self.btn.move(0,50)
        self.timer = QTimer(self.qpb)

        #Qtimer类当定时间隔到了之后会发射一个timerout信号，这样就可以变更进度条的值
        self.timer.timeout.connect(self.change_value)

        #这里用按钮来启动一个定时器（光创建一个定时器并不会起作用），同时这个按钮也可以选择停止一个定时器
        self.btn.clicked.connect(self.change_progressbar)

    def change_value(self):
        if self.qpb.value() < self.qpb.maximum():
            self.qpb.setValue(self.qpb.value()+2)
        else:
            self.timer.stop()

    def change_progressbar(self):
        if self.btn.text() == "开始":
            self.btn.setText("结束")
            self.timer.start(10)
        else:
            self.btn.setText("开始")
            self.timer.stop()

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())
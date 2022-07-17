from PyQt5.QtWidgets import *
import sys
import qt5_applications
import os
class QlineEditEchoMode(QWidget):
    def __init__(self):
        super(QlineEditEchoMode,self).__init__()
        self.initUI()
    def initUI(self):
        self.setWindowTitle('文本输入框的回显模式')
        FormLayout=QFormLayout()
        normalLineEdit=QLineEdit()
        noEchoLineEdit=QLineEdit()
        passwordLineEdit=QLineEdit()
        passwordechoLineEdit=QLineEdit()
        FormLayout.addRow("Normal",normalLineEdit)
        FormLayout.addRow("noEcho",noEchoLineEdit)
        FormLayout.addRow("password",passwordLineEdit)
        FormLayout.addRow("passwordonecho",passwordechoLineEdit)
        normalLineEdit.setPlaceholderText("Normal")
        noEchoLineEdit.setPlaceholderText("noEcho")
        passwordLineEdit.setPlaceholderText("password")
        passwordechoLineEdit.setPlaceholderText("passwordonecho")
        normalLineEdit.setEchoMode(QLineEdit.Normal)
        noEchoLineEdit.setEchoMode(QLineEdit.NoEcho)
        passwordLineEdit.setEchoMode(QLineEdit.Password)
        passwordechoLineEdit.setEchoMode(QLineEdit.PasswordEchoOnEdit)
        self.setLayout(FormLayout)

if __name__ == '__main__':
    app=QApplication(sys.argv)
    main=QlineEditEchoMode()
    main.show()
    sys.exit(app.exec_())
